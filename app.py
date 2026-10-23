import csv
import io
import json
import os
import re
import secrets
import ssl
import urllib.error
import urllib.request
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

try:
    import openpyxl
except ImportError:
    openpyxl = None


ROOT = Path(__file__).parent
STATIC = ROOT / "static"
MAX_UPLOAD = 20 * 1024 * 1024
DEFAULT_ZENTAO_ENDPOINT = "https://hub-zentao.lhftx.com/index.php?m=testcase&f=showImport&productID=18&branch=0&zin=1"
DEFAULT_ZENTAO_SUITE_ENDPOINT = "https://hub-zentao.lhftx.com/index.php?m=testcase&f=browse&productID=18&branch=0&browseType=bySuite&param=78&zin=1"
DEFAULT_FORM_VALUES = {"requirement": "项目管理(#413)", "module": "/项目管理(#222)", "product": "2.0CIT内部平台(#18)"}
PENDING_IMPORTS = {}

ALIASES = {
    "case_number": ["用例编号", "编号", "case id", "case number", "id"],
    "title": ["用例标题", "标题", "用例名称", "case title", "title", "name"],
    "product": ["所属产品", "产品", "product", "产品名称"],
    "module": ["所属模块", "模块", "module", "模块名称"],
    "scenario": ["所属场景", "场景", "scenario", "场景名称"],
    "requirement": ["相关研发需求", "研发需求", "关联需求", "requirement", "story"],
    "release": ["用例版本", "版本", "发布", "release", "version"],
    "suite": ["所属套件", "套件", "suite", "测试套件"],
    "keywords": ["关键词", "关键字", "keywords", "keyword"],
    "precondition": ["前置条件", "前提条件", "precondition", "前置"],
    "steps": ["操作步骤", "步骤", "测试步骤", "steps", "step"],
    "expected": ["预期结果", "期望结果", "预期", "expected result", "expected"],
    "priority": ["优先级", "priority", "级别"],
    "type": ["用例类型", "类型", "case type", "type"],
    "stage": ["适用阶段", "阶段", "stage"],
}


def norm(value):
    return re.sub(r"\s+", "", str(value or "").strip().lower())


def parse_csv(data):
    text = None
    for encoding in ("utf-8-sig", "gb18030", "gbk"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = data.decode("utf-8", errors="replace")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel
    parsed = list(csv.reader(io.StringIO(text), dialect=dialect))
    if not parsed:
        return []
    headers = [str(value or "").strip() for value in parsed[0]]
    # Some Excel exports leave quotes inside multiline cells unescaped. The
    # csv reader then emits the continuation as a new record. Reattach such
    # records to the previous case so complete content is not lost.
    repaired = []
    for record in parsed[1:]:
        values = list(record)
        first = str(values[0] if values else "").strip().strip("\t")
        is_case = bool(
            re.match(r"^[A-Za-z][A-Za-z0-9_]*-[A-Za-z0-9_.-]*\d+$", first)
            or re.fullmatch(r"\d+", first)
            or (len(values) >= 6 and sum(bool(str(value or "").strip()) for value in values[1:6]) >= 3)
        )
        if is_case or not repaired:
            repaired.append(values)
            continue
        if not any(str(value or "").strip() for value in values):
            continue
        previous = repaired[-1]
        while len(previous) < max(len(headers), 7):
            previous.append("")
        # In the affected exports the split occurs inside the expected-result
        # column; keep the fragment as a new line and recover its priority.
        fragment = first
        if fragment:
            previous[5] = "\n".join(x for x in (str(previous[5] or "").strip(), fragment) if x)
        if not str(previous[6] or "").strip():
            for value in values[1:]:
                candidate = str(value or "").strip()
                if candidate in {"1", "2", "3", "4"}:
                    previous[6] = candidate
                    break
    return [dict(zip(headers, row)) for row in repaired]


def parse_xlsx(data):
    if openpyxl is None:
        raise ValueError("当前 Python 环境缺少 openpyxl，无法读取 Excel 文件")
    book = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheet = book.active
    values = list(sheet.values)
    if not values:
        return []
    headers = [str(x or "").strip() for x in values[0]]
    return [dict(zip(headers, row)) for row in values[1:] if any(x not in (None, "") for x in row)]


def parse_file(filename, data):
    ext = Path(filename).suffix.lower()
    if ext == ".csv":
        return parse_csv(data)
    if ext in (".xlsx", ".xlsm"):
        return parse_xlsx(data)
    raise ValueError("仅支持 .xlsx、.xlsm 或 .csv 文件")


def find_value(row, key):
    lookup = {norm(k): v for k, v in row.items() if k is not None}
    for alias in ALIASES[key]:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    return ""


def source_defaults(rows):
    if not rows:
        return {}
    first = rows[0]
    return {
        "product": str(find_value(first, "product") or "").strip(),
        "module": str(find_value(first, "module") or "").strip(),
        "scenario": str(find_value(first, "scenario") or "").strip(),
        "requirement": str(find_value(first, "requirement") or "").strip(),
        "release": str(find_value(first, "release") or "").strip(),
        "type": str(find_value(first, "type") or "功能测试").strip(),
        "stage": str(find_value(first, "stage") or "系统测试").strip(),
    }


def transform(rows, variables):
    result = []
    errors = []
    for index, row in enumerate(rows, start=2):
        suite = str(find_value(row, "suite") or "").strip()
        source_precondition = str(find_value(row, "precondition") or "").strip()
        precondition = "\n".join(x for x in (suite, source_precondition) if x)
        item = {
            "case_number": str(find_value(row, "case_number") or "").strip(),
            "title": str(find_value(row, "title") or "").strip(),
            "product": str(variables.get("product") or find_value(row, "product") or "").strip(),
            "module": str(variables.get("module") or find_value(row, "module") or "").strip(),
            "scenario": str(variables.get("scenario") or find_value(row, "scenario") or "").strip(),
            "precondition": precondition,
            "steps": str(find_value(row, "steps") or "").strip(),
            "expected": str(find_value(row, "expected") or "").strip(),
            "priority": str(find_value(row, "priority") or variables.get("priority") or "3").strip(),
            "type": str(find_value(row, "type") or variables.get("type") or "功能测试").strip(),
            "stage": str(find_value(row, "stage") or variables.get("stage") or "系统测试").strip(),
            "keywords": str(find_value(row, "keywords") or find_value(row, "case_number") or "").strip(),
            "requirement": str(variables.get("requirement") or find_value(row, "requirement") or "").strip(),
            "release": str(variables.get("release") or find_value(row, "release") or "").strip(),
            "project": variables.get("project", "").strip(),
        }
        missing = [label for key, label in (("title", "标题"), ("steps", "操作步骤"), ("expected", "预期结果")) if not item[key]]
        if missing:
            errors.append({"row": index, "message": "缺少必填字段：" + "、".join(missing)})
        result.append(item)
    return result, errors


def validate_for_zentao(cases, variables):
    errors = []
    requirement = extract_id(variables.get("requirement"))
    module = extract_id(variables.get("module"))
    if requirement and not requirement.isdigit():
        errors.append({"row": 0, "message": "相关研发需求需要包含禅道需求 ID，例如 项目管理(#413)"})
    if module and not module.isdigit():
        errors.append({"row": 0, "message": "所属模块需要包含禅道模块 ID，例如 /项目管理(#222)"})
    seen = set()
    for index, case in enumerate(cases, start=2):
        key = case.get("keywords", "").strip()
        if key and key in seen:
            errors.append({"row": index, "message": f"关键词/用例编号重复：{key}"})
        if key:
            seen.add(key)
        if case.get("priority", "3") not in {"1", "2", "3", "4"}:
            errors.append({"row": index, "message": "优先级只能是 1、2、3 或 4"})
    return errors


def make_xlsx(cases):
    if openpyxl is None:
        raise ValueError("当前 Python 环境缺少 openpyxl，无法生成 Excel 文件")
    book = openpyxl.Workbook()
    sheet = book.active
    headers = ["用例名称", "所属产品", "所属模块", "所属场景", "前置条件", "步骤", "预期", "优先级", "用例类型", "适用阶段", "相关研发需求", "用例版本"]
    sheet.append(headers)
    for case in cases:
        sheet.append([case.get("title", ""), case.get("product", ""), case.get("module", ""), case.get("scenario", ""), case.get("precondition", ""), case.get("steps", ""), case.get("expected", ""), case.get("priority", "3"), case.get("type", "功能测试"), case.get("stage", "系统测试"), case.get("requirement", ""), case.get("release", "")])
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def make_import_csv(cases):
    headers = ["所属产品", "所属模块", "相关研发需求", "所属场景", "用例名称", "前置条件", "步骤", "预期", "实际情况", "关键词", "优先级", "用例类型", "适用阶段", "用例状态", "B", "R", "S", "执行人", "执行时间", "结果", "由谁创建", "创建日期", "最后修改者", "修改日期", "用例版本", "相关用例", "附件"]
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(headers)
    for case in cases:
        writer.writerow([case.get("product", ""), case.get("module", ""), case.get("requirement", ""), case.get("scenario", ""), case.get("title", ""), case.get("precondition", ""), case.get("steps", ""), case.get("expected", ""), "", case.get("case_number", ""), case.get("priority", "3"), case.get("type", "功能测试"), case.get("stage", "系统测试"), "", "", "", "", "", "", "", "", "", "", "", case.get("release", ""), "", ""])
    return output.getvalue().encode("utf-8-sig")


def extract_id(value):
    match = re.search(r"#(\d+)", str(value or ""))
    return match.group(1) if match else str(value or "").strip()


def zentao_type(value):
    value = str(value or "").lower()
    return {"功能测试": "feature", "性能测试": "performance", "安全测试": "security", "配置相关": "config", "安装部署": "install", "其他": "other"}.get(value, value or "feature")


def zentao_stage(value):
    value = str(value or "").lower()
    return {"单元测试": "unittest", "集成测试": "integration", "系统测试": "system", "验收测试": "acceptance", "回归测试": "feature", "feature": "feature"}.get(value, value or "feature")


def make_multipart(fields):
    boundary = "----CodexZentao" + secrets.token_hex(12)
    chunks = []
    for name, value in fields:
        value = str(value or "")
        chunks.append((f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"\r\n\r\n{value}\r\n").encode("utf-8"))
    return boundary, b"".join(chunks) + f"--{boundary}--\r\n".encode("utf-8")


def post_multipart(endpoint, body, content_type, cookie):
    request = urllib.request.Request(endpoint, data=body, method="POST", headers={"Content-Type": content_type, "Accept": "*/*", "X-Requested-With": "XMLHttpRequest", "Cookie": cookie})
    context = ssl.create_default_context()
    opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=context))
    with opener.open(request, timeout=30) as response:
        return response.status, response.read().decode("utf-8", errors="replace")


def clean_endpoint(value):
    endpoint = str(value or "").strip().strip('"\'')
    endpoint = endpoint.replace('^"', '').replace('"^', '').replace('^&', '&')
    if not re.match(r"^https?://", endpoint, re.IGNORECASE):
        raise ValueError("禅道地址格式错误，请使用完整的 http:// 或 https:// 地址，不要粘贴 cURL 命令。")
    return endpoint


def build_suite_endpoint(variables):
    endpoint = clean_endpoint(variables.get("endpoint") or DEFAULT_ZENTAO_SUITE_ENDPOINT)
    parsed = urlparse(endpoint)
    params = dict(parse_qsl(parsed.query, keep_blank_values=True))
    params.update({
        "m": "testcase",
        "f": "browse",
        "productID": extract_id(variables.get("product") or "18") or "18",
        "branch": "0",
        "browseType": "bySuite",
        "param": str(variables.get("suite") or "78").strip() or "78",
        "zin": "1",
    })
    return urlunparse(parsed._replace(query=urlencode(params)))


def build_run_case_endpoint(variables):
    endpoint = clean_endpoint(variables.get("endpoint") or DEFAULT_ZENTAO_SUITE_ENDPOINT)
    parsed = urlparse(endpoint)
    params = {
        "m": "testtask",
        "f": "runCase",
        "runID": str(variables.get("runId") or "0"),
        "caseID": str(variables.get("caseId") or "").strip(),
        "version": str(variables.get("version") or "1"),
    }
    if not params["caseID"]:
        raise ValueError("请选择要执行的禅道用例")
    return urlunparse(parsed._replace(query=urlencode(params)))


def build_bug_create_endpoint(variables):
    endpoint = clean_endpoint(variables.get("endpoint") or DEFAULT_ZENTAO_SUITE_ENDPOINT)
    parsed = urlparse(endpoint)
    bug_fields = variables.get("bugFields") if isinstance(variables.get("bugFields"), dict) else {}
    product_id = str(variables.get("productID") or variables.get("product") or bug_fields.get("product") or "18").strip()
    extras = ",".join([
        f"productID={product_id}",
        f"moduleID={str(variables.get('moduleID') or variables.get('module') or bug_fields.get('module') or '').strip()}",
        f"projectID={str(variables.get('projectID') or variables.get('project') or bug_fields.get('project') or '').strip()}",
        f"executionID={str(variables.get('executionID') or variables.get('execution') or bug_fields.get('execution') or '').strip()}",
        "regionID=",
        "allBuilds=",
        "allUsers=on",
    ])
    params = {"m": "bug", "f": "create", "productID": product_id, "branch": "", "extras": extras}
    return urlunparse(parsed._replace(query=urlencode(params)))


def create_pending(kind, endpoint, data):
    token = secrets.token_urlsafe(24)
    PENDING_IMPORTS[token] = {"kind": kind, "endpoint": endpoint, "fields": [], "data": data, "result": {"status": "pending"}}
    return token


def submit_zentao_query(variables):
    endpoint = build_suite_endpoint(variables)
    token = create_pending("query", endpoint, {"suite": str(variables.get("suite") or "78")})
    launch_url = endpoint.split("#", 1)[0] + "#codex-zentao-query=" + token
    return {"mode": "extension-query", "message": "正在通过当前禅道会话查询套件用例。", "data": {"launchUrl": launch_url, "token": token}}


def submit_zentao_execution(variables):
    endpoint = build_run_case_endpoint(variables)
    steps = variables.get("steps") if isinstance(variables.get("steps"), list) else []
    token = create_pending("execute", endpoint, {"caseId": str(variables.get("caseId") or ""), "version": str(variables.get("version") or "1"), "runId": str(variables.get("runId") or "0"), "steps": steps})
    launch_url = endpoint.split("#", 1)[0] + "#codex-zentao-execute=" + token
    return {"mode": "extension-execute", "message": "正在通过当前禅道会话同步执行结果。", "data": {"launchUrl": launch_url, "token": token}}


def submit_zentao_execution_and_bug(variables):
    case_id = str(variables.get("caseId") or "").strip()
    if not case_id:
        raise ValueError("请选择要执行的禅道用例")
    execution_endpoint = build_run_case_endpoint(variables)
    bug_endpoint = build_bug_create_endpoint(variables)
    steps = variables.get("steps") if isinstance(variables.get("steps"), list) else []
    bug_fields = variables.get("bugFields") if isinstance(variables.get("bugFields"), dict) else {}
    token = create_pending("execute-and-bug", execution_endpoint, {
        "caseId": case_id,
        "version": str(variables.get("version") or "1"),
        "runId": str(variables.get("runId") or "0"),
        "steps": steps,
        "bugEndpoint": bug_endpoint,
        "bugFields": bug_fields,
        "bugTitle": str(variables.get("bugTitle") or "").strip(),
        "bugSteps": str(variables.get("bugSteps") or ""),
        "story": str(variables.get("story") or "").strip(),
    })
    launch_url = execution_endpoint.split("#", 1)[0] + "#codex-zentao-execute-bug=" + token
    return {"mode": "extension-execute-bug", "message": "正在同步执行结果并提交 Bug。", "data": {"launchUrl": launch_url, "token": token}}


def submit_zentao_image_upload(variables):
    file_data = variables.get("file") if isinstance(variables.get("file"), dict) else {}
    if not str(file_data.get("fileId") or "").strip() or not str(file_data.get("data") or "").strip():
        raise ValueError("图片预上传缺少有效的图片数据")
    bug_endpoint = build_bug_create_endpoint(variables)
    token = create_pending("image-upload", bug_endpoint, {
        "bugEndpoint": bug_endpoint,
        "file": file_data,
    })
    return {"mode": "extension-image-upload", "message": "正在后台预上传图片。", "data": {"token": token}}


def submit_zentao_case_details(variables):
    endpoint = build_run_case_endpoint(variables)
    token = create_pending("case-details", endpoint, {"caseId": str(variables.get("caseId") or ""), "version": str(variables.get("version") or "1"), "runId": str(variables.get("runId") or "0")})
    bridge_endpoint = clean_endpoint(variables.get("bridgeEndpoint") or DEFAULT_ZENTAO_SUITE_ENDPOINT)
    launch_url = bridge_endpoint.split("#", 1)[0] + "#codex-zentao-case=" + token
    return {"mode": "extension-case", "message": "正在读取禅道用例详情。", "data": {"launchUrl": launch_url, "token": token}}


def build_web_confirm_fields(cases, variables):
    product_id = extract_id(variables.get("product"))
    default_module = extract_id(variables.get("module"))
    story_id = extract_id(variables.get("requirement"))
    fields = []
    for index, case in enumerate(cases, start=1):
        fields.extend([
            (f"rawID[{index}]", ""),
            (f"product[{index}]", ""),
            (f"title[{index}]", case.get("title")),
            (f"module[{index}]", extract_id(case.get("module")) or default_module or "0"),
            (f"story[{index}]", story_id),
            (f"type[{index}]", zentao_type(case.get("type"))),
            (f"pri[{index}]", case.get("priority", "3")),
            (f"precondition[{index}]", case.get("precondition")),
            (f"keywords[{index}]", case.get("keywords", "")),
            (f"stage[{index}][]", ""),
            (f"steps[{index}]", case.get("steps")),
            (f"expects[{index}]", case.get("expected")),
        ])
    fields.extend([("isEndPage", "1"), ("pagerID", "1"), ("insert", "")])
    return fields


def submit_web_confirm(cases, variables):
    endpoint = clean_endpoint(variables.get("endpoint"))
    fields = build_web_confirm_fields(cases, variables)
    token = secrets.token_urlsafe(24)
    PENDING_IMPORTS[token] = {"endpoint": endpoint, "fields": fields, "result": {"status": "pending"}}
    launch_url = "https://hub-zentao.lhftx.com/index.php?m=index&f=index#codex-zentao-import=" + token
    return {"mode": "extension-submit", "message": "已生成一次性导入任务，正在由禅道同源桥接器提交。", "data": {"launchUrl": launch_url}}


def submit_web_import(cases, variables):
    endpoint = clean_endpoint(variables.get("endpoint"))
    cookie = (variables.get("cookie") or "").strip()
    if not cookie:
        raise ValueError("网页导入需要登录 Cookie。请从浏览器复制 Cookie 到本地输入框，不要把 Cookie 写入代码。")
    boundary = "----CodexZentao" + secrets.token_hex(12)
    csv_data = make_import_csv(cases)
    chunks = []
    chunks.append((f"--{boundary}\r\nContent-Disposition: form-data; name=\"file[]\"; filename=\"zentao-testcases.csv\"\r\nContent-Type: text/csv\r\n\r\n").encode() + csv_data + b"\r\n")
    chunks.append((f"--{boundary}\r\nContent-Disposition: form-data; name=\"encode\"\r\n\r\nutf-8\r\n").encode())
    body = b"".join(chunks) + f"--{boundary}--\r\n".encode()
    try:
        status, raw = post_multipart(endpoint, body, f"multipart/form-data; boundary={boundary}", cookie)
        if status >= 400:
            raise ValueError(f"禅道网页导入返回 HTTP {status}: {raw[:500]}")
        return {"mode": "web-preview", "message": "文件已上传到禅道网页导入入口；该接口通常只负责解析预览，是否还需在禅道页面确认请查看返回内容。", "data": {"status": status, "preview": raw[:500]}}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"禅道网页导入返回 HTTP {exc.code}: {detail[:500]}")
    except urllib.error.URLError as exc:
        raise ValueError("无法连接禅道网页导入地址：" + str(exc.reason) + "（地址：" + repr(endpoint) + "）")


def zentao_submit(cases, variables):
    endpoint = clean_endpoint(variables.get("endpoint") or os.getenv("ZENTAO_API_ENDPOINT", "")) if (variables.get("endpoint") or os.getenv("ZENTAO_API_ENDPOINT", "")) else ""
    token = (variables.get("token") or os.getenv("ZENTAO_API_TOKEN", "")).strip()
    if not endpoint:
        return {"mode": "export", "message": "未配置禅道 API 地址，已生成导入文件。", "data": None}
    if "m=testcase" in endpoint and ("f=import" in endpoint or "f=showImport" in endpoint):
        if "f=showImport" in endpoint:
            return submit_web_confirm(cases, variables)
        return submit_web_import(cases, variables)
    payload = {"cases": cases, "product": variables.get("product", ""), "project": variables.get("project", ""), "release": variables.get("release", ""), "requirement": variables.get("requirement", "")}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(endpoint, data=body, method="POST", headers={"Content-Type": "application/json", "Accept": "application/json"})
    if token:
        request.add_header("Token", token)
        request.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            raw = response.read().decode("utf-8", errors="replace")
            return {"mode": "api", "message": "禅道接口提交成功。", "data": json.loads(raw) if raw else {}}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"禅道接口返回 HTTP {exc.code}: {detail[:500]}")
    except urllib.error.URLError as exc:
        raise ValueError("无法连接禅道接口：" + str(exc.reason))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_args):
        return

    def send_json(self, payload, status=HTTPStatus.OK):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/health":
            self.send_json({"ok": True})
            return
        if path == "/api/config":
            self.send_json({"endpoint": DEFAULT_ZENTAO_ENDPOINT, "suiteEndpoint": DEFAULT_ZENTAO_SUITE_ENDPOINT, "cookie": os.getenv("ZENTAO_COOKIE", ""), "defaults": DEFAULT_FORM_VALUES})
            return
        if path.startswith("/api/pending/"):
            parts = path.strip("/").split("/")
            token = parts[2] if len(parts) > 2 else ""
            if len(parts) > 3 and parts[3] == "result":
                payload = PENDING_IMPORTS.get(token)
                if payload is None:
                    self.send_json({"error": "导入任务不存在"}, HTTPStatus.NOT_FOUND)
                else:
                    self.send_json(payload.get("result", {"status": "pending"}))
                return
            payload = PENDING_IMPORTS.get(token)
            if payload is None:
                self.send_json({"error": "导入任务不存在或已被读取"}, HTTPStatus.NOT_FOUND)
            else:
                self.send_json({"kind": payload.get("kind", "import"), "endpoint": payload["endpoint"], "fields": payload.get("fields", []), "data": payload.get("data", {})})
            return
        target = STATIC / ("index.html" if path in ("/", "") else path.lstrip("/"))
        if target.exists() and target.is_file() and STATIC in target.parents:
            data = target.read_bytes()
            content_type = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8"}.get(target.suffix.lower(), "text/plain; charset=utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        pending_prefix = "/api/pending/"
        request_path = urlparse(self.path).path
        if request_path.startswith(pending_prefix) and request_path.endswith("/result"):
            token = request_path[len(pending_prefix):-len("/result")].strip("/")
            payload = PENDING_IMPORTS.get(token)
            if payload is None:
                self.send_json({"error": "导入任务不存在"}, HTTPStatus.NOT_FOUND)
                return
            length = int(self.headers.get("Content-Length", "0"))
            try:
                payload["result"] = json.loads(self.rfile.read(length).decode("utf-8"))
                self.send_json({"ok": True})
            except Exception as exc:
                self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if request_path in ("/api/zentao/query", "/api/zentao/case", "/api/zentao/execute", "/api/zentao/execute-and-bug", "/api/zentao/upload-image"):
            length = int(self.headers.get("Content-Length", "0"))
            try:
                variables = json.loads(self.rfile.read(length).decode("utf-8"))
                if request_path.endswith("/query"):
                    submit = submit_zentao_query(variables)
                elif request_path.endswith("/case"):
                    submit = submit_zentao_case_details(variables)
                elif request_path.endswith("/execute-and-bug"):
                    submit = submit_zentao_execution_and_bug(variables)
                elif request_path.endswith("/upload-image"):
                    submit = submit_zentao_image_upload(variables)
                else:
                    submit = submit_zentao_execution(variables)
                self.send_json({"ok": True, "submit": submit})
            except Exception as exc:
                self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if request_path != "/api/process":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD:
            self.send_json({"error": "文件不能超过 20 MB"}, HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return
        try:
            raw = self.rfile.read(length)
            import cgi
            form = cgi.FieldStorage(fp=io.BytesIO(raw), headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")})
            upload = form["file"] if "file" in form else None
            if upload is None or not getattr(upload, "filename", ""):
                raise ValueError("请先选择测试用例文件")
            variables = json.loads(form.getfirst("variables", "{}"))
            rows = parse_file(upload.filename, upload.file.read())
            action = variables.get("action", "preview")
            override = form.getfirst("cases", "")
            if action == "submit" and override:
                cases = json.loads(override)
                errors = []
            else:
                cases, errors = transform(rows, variables)
            if errors:
                self.send_json({"error": "文件校验未通过", "errors": errors, "count": len(cases)}, HTTPStatus.UNPROCESSABLE_ENTITY)
                return
            if action == "submit":
                submit_errors = validate_for_zentao(cases, variables)
                if submit_errors:
                    self.send_json({"error": "禅道格式校验未通过", "errors": submit_errors, "count": len(cases)}, HTTPStatus.UNPROCESSABLE_ENTITY)
                    return
            submit = zentao_submit(cases, variables) if action == "submit" else {"mode": "preview", "message": "预览生成成功。", "data": None}
            file_data = make_xlsx(cases)
            csv_data = make_import_csv(cases)
            import base64
            self.send_json({"ok": True, "count": len(cases), "cases": cases, "defaults": source_defaults(rows), "download": "data:text/csv;base64," + base64.b64encode(csv_data).decode(), "downloadXlsx": "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + base64.b64encode(file_data).decode(), "submit": submit})
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)


def main():
    port = int(os.getenv("ZENTAO_WORKSPACE_PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"禅道测试用例工作区已启动：http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
