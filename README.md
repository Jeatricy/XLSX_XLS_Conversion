# Excel 批量格式转换工具（xlsx ⇄ xls）

调用本机 Microsoft Excel 的 COM 接口，把一个目录（含子目录）下的 Excel 文件在 `.xlsx` 与 `.xls` 两种格式之间批量互转。

## 功能

- 双向转换：`xlsx → xls` 与 `xls → xlsx` 二选一，通过界面单选项切换。
- 转换强制使用 Excel COM 引擎（`win32com.client` 调用独立的 `Excel.Application` 实例，隐藏窗口、关闭提示）；不会连接或关闭用户已打开的 Excel 会话。点击开始前会先提示格式兼容性风险，并检测 Microsoft Excel 与 pywin32 是否可用。
- 递归扫描源目录下待转换格式的文件，输出时保持原有的相对目录结构，并在输出目录下先建立一层以「源文件夹名」命名的根文件夹。
- 同名文件三种处理策略：跳过、覆盖、自动重命名（在文件名后依次追加 `(1)`、`(2)` …）。
- CustomTkinter 图形界面：源目录、输出目录、转换方向、同名策略、运行日志，界面字体统一为微软雅黑；转换在后台线程中执行，点「停止」会在当前文件转换完成后中断。
- 日志逐条记录每个文件的转换结果，失败时打印具体错误原因；结束时汇总「成功 / 失败 / 跳过」数量并弹窗提示。
- 输出格式写死为 Excel 常量 51（xlOpenXMLWorkbook，即 `.xlsx`）与 56（xlExcel8，即 `.xls`）；窗口图标尝试从当前工作目录读取 `icon.ico`，文件不存在时静默忽略、不影响运行。

## 环境要求

- Python 3.10+（`resolve_conflict` 的返回类型标注使用了 `Path | None` 语法，且文件未启用 `from __future__ import annotations`）
- 依赖：customtkinter、pywin32（提供 `win32com.client`）
- 本机需安装 Microsoft Excel，转换实际由 Excel 完成

## 使用方式

### 方式一：从源码运行

```bat
pip install customtkinter pywin32
python Excel新旧格式转换.py
```

运行后选择源目录与输出目录，选择转换方向和同名文件策略，点击「开始转换」，日志区会显示每个文件的处理结果。可选的窗口图标 `icon.ico` 需放在运行程序时的当前工作目录下，本目录未随附该文件。

## 文件说明

| 文件 | 说明 |
|---|---|
| `Excel新旧格式转换.py` | 主程序（也是唯一脚本）。包含 CustomTkinter 界面、路径规划、同名文件策略与 Excel COM 转换逻辑 |
| `.gitignore` | Git 忽略规则（Python 通用模板），不影响程序运行 |

本目录不包含 `.exe` 等打包产物，也没有随附示例工作簿：转换对象由运行时选择的源目录决定，输出写入所选输出目录。
