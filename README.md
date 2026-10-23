# 禅道测试用例工作台

这是一个本地运行的禅道测试用例工作台，支持：

- 上传 CSV、XLSX、XLSM 测试用例并转换为禅道导入格式
- 在转换结果预览中编辑用例并校验必填字段
- 查询指定禅道套件中的用例
- 查看用例前置条件、步骤和预期结果
- 提交步骤执行结果、实际情况和图片
- 同步执行结果并提交 Bug
- 查看当前用例的执行状态、关联 Bug 和操作日志

## 环境要求

- Windows 10/11
- Python 3.10 或更高版本
- Microsoft Edge 或 Chromium 内核浏览器
- 已登录目标禅道站点

## 安装依赖

在本目录打开 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 启动工作台

```powershell
.\start.ps1
```

启动后打开：<http://127.0.0.1:8765/>

也可以不自动打开浏览器：

```powershell
.\start.ps1 -NoBrowser
```

检查服务：

```powershell
Invoke-RestMethod http://127.0.0.1:8765/api/health
```

正常返回 `{"ok":true}`。

## 安装浏览器扩展

1. 在 Edge 地址栏打开 `edge://extensions/`。
2. 开启“开发人员模式”。
3. 点击“加载解压缩的扩展”，选择本目录下的 `edge-extension` 文件夹。
4. 打开禅道并保持登录状态，再返回工作台操作。

扩展用于在当前登录的禅道标签页中执行查询、同步执行结果、上传图片和提交 Bug。Cookie、Token 和登录会话只应在本机使用，不要写入代码或提交到仓库。

## 配置说明

默认禅道地址和示例字段值位于 `app.py` 顶部。若要连接其他禅道实例，需要同时调整：

- `app.py` 中的默认导入地址和套件查询地址
- `edge-extension/manifest.json` 中的 `host_permissions` 和 `content_scripts.matches`
- `edge-extension/background.js`、`content.js` 中的目标站点匹配规则

建议将实际禅道地址、Cookie 和 Token 放在本机配置或环境变量中，不要公开到 GitHub。

## 目录结构

```text
app.py                 本地 HTTP 服务和转换接口
static/index.html      工作台页面
edge-extension/        Edge 浏览器扩展
start.ps1              PowerShell 启动脚本
start.cmd              双击启动入口
requirements.txt       Python 依赖
```
