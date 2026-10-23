---
name: zentao-testcase-importer
description: 将用户上传的 Excel 或 CSV 测试用例转换为禅道导入格式，并在用户提供禅道 API 地址和 Token 时提交到禅道。用于批量导入测试用例、关联研发需求和版本、校验字段以及生成禅道导入文件。
---

# 禅道测试用例导入工作区

使用本目录的 `app.py` 启动本地工作区。工作区支持上传 `.xlsx`、`.xlsm`、`.csv`，页面只要求填写相关研发需求、所属模块、所属产品和登录 Cookie；版本、场景、类型、阶段、优先级等从文件中沿用。

## 启动

在本目录执行：

```powershell
.\start.ps1
```

然后打开 `http://127.0.0.1:8765`。网页导入地址已内置为当前禅道实例，也可通过代码中的 `DEFAULT_ZENTAO_ENDPOINT` 调整。

## 执行规则

1. 先让用户上传原始文件并填写四个输入项。
2. 执行“转换”，检查前 50 条结果及必填字段错误，并下载生成的禅道 CSV。
3. 只有用户点击“提交到禅道”时才发送网页表单；网页导入优先由当前浏览器直接提交，使用已登录的禅道会话，避免本地后台进程访问内网失败；Cookie 只通过表单或环境变量传递，不写入文件。
4. 记录接口返回错误；不要在失败时重复提交。

## 字段识别

脚本会按中英文别名识别标题、产品、模块、场景、研发需求、版本、前置条件、步骤、预期结果、优先级、类型和阶段；如果存在“所属套件”，会按原转换脚本的规则合并到“前置条件”。必填字段是标题、操作步骤和预期结果。

## 禅道适配注意事项

示例文件使用 GBK 编码，脚本会自动尝试 UTF-8、GB18030 和 GBK。禅道各版本的 API 路径和字段可能不同。配置 JSON API 地址时，工作区发送包含 `cases`、`product`、`project`、`release` 和 `requirement` 的 JSON。对于网页导入地址 `m=testcase&f=showImport`，工作区会按浏览器表单提交 `rawID[i]`、`product[i]`、`title[i]`、`module[i]`、`story[i]`、`type[i]`、`pri[i]`、`precondition[i]`、`keywords[i]`、`stage[i][]`、`steps[i]`、`expects[i]`、`isEndPage`、`pagerID` 和 `insert`。浏览器 Network 中的 `Cookie`（例如 `zentaosid`）是登录会话，不等同于 API Token，不要把它写入 Skill 或提交到代码仓库。
