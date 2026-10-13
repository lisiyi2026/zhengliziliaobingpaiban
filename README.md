# 考研文件整理

本项目默认入口是夸克扫描王导出的 Word：将基础 DOCX 放入 `input/word/` 后执行结构识别、排版和 DOCX 兼容清理。项目不调用夸克 API，也不保存夸克任务缓存。PDF/图片直读仅作为备用入口，不参与默认 Word 流程。

## 当前完整流程

`夸克 DOCX → 提取段落、公式、符号、图片、表格和布局索引 → 本地规则结构判断 → 固定排版 → 分页/留白/图片风险检查 → DOCX OOXML Sanitizer → 结构校验 → 输出兼容 DOCX → 按需执行 Word→PDF sample/full 视觉检查`。

默认流程不使用 AI、MinerU 或 PDF-Extract-Kit，也不重新扫描原 PDF。原生 OMML 公式保留；不可靠公式不强行改写，避免错误 OCR 覆盖原内容。

PDF 直读仍保留 `--pdf-engine` 备用参数，但不应替代夸克 DOCX 主流程。需要 PDF 直读时才显式选择对应引擎；日常处理不启动 MinerU/PDF-Extract-Kit 桥接器。

## 使用方式

把待处理的 `.docx` 或 `.doc` 放入 `input/word/`，运行：

```powershell
python scripts/quark_pdf_to_word.py
```

直接处理 PDF（推荐先用两页预览验证环境）：

```powershell
python scripts/quark_pdf_to_word.py `
  --input input/pdfs/机械原理习题集.pdf `
  --output output/word/机械原理习题集.docx `
  --pdf-engine auto `
  --max-pages 2
```

可用 `--pdf-engine mineru` 强制 MinerU，或 `--pdf-engine office` 跳过 MinerU。若 MinerU 不在项目虚拟环境，可通过 `--mineru-python C:\\path\\to\\python.exe` 指定解释器；`--mineru-backend pipeline` 和 `--mineru-method auto` 是默认设置。

如果主脚本是在 Codex 沙箱中运行，请先在普通 Windows PowerShell 启动 MinerU 桥接器并保持窗口打开：

```powershell
.\scripts\start_mineru_bridge.cmd
```

公式候选检测还可以使用已安装的 PDF-Extract-Kit（YOLOv8 MFD 权重）。在普通 Windows PowerShell 中单独测试时启动：

```powershell
.\scripts\start_pdfkit_bridge.cmd
```

`run_all.cmd` 默认只处理 `input/word/` 中的夸克 DOCX，不自动打开 MinerU、PDF-Extract-Kit 或 LibreOffice 桥接器。兼容清理在本地直接完成；只有 `.doc` 转换或显式视觉检查才需要 LibreOffice。

桥接器使用 `.venv-mineru\\Scripts\\python.exe` 在普通用户账户中执行 MinerU，主脚本只通过 `tmp/mineru-bridge/` 交换任务和结果。这样可以绕过 Codex 沙箱对 Torch DLL 的加载限制；桥接器不可用时，`--pdf-engine auto` 仍会回退到 LibreOffice。

日常测试和处理可以直接使用固定全流程入口：

```powershell
.\scripts\run_all.cmd
```

它优先处理 `input/word/` 中的夸克 DOCX；如果该目录没有 DOCX/DOC，才处理 `input/pdfs/` 中的 PDF。默认不执行 Word→PDF 渲染，只做 DOCX/XML 层面的分页、留白、图片 anchor 和对象安全检查。首次测试可附加 `--max-pages 2`；需要真实页面抽查时附加 `--render-layout`，或组合使用 `--compare sample --visual-compare`。

结果写入 `output/word/`，结构化检查报告写入 `output/reports/`。Word-first 流程不启动公式模型；兼容层记录 anchor、定位节点、隐藏分页和 section 规范化结果。处理过程使用临时目录 `tmp/word/`，结束后自动清理。

也可以指定单个文件：

```powershell
python scripts/quark_pdf_to_word.py `
  --input input/word/机械原理习题集.docx `
  --output output/word/机械原理习题集_整理版.docx
```

## 可选的原 PDF 视觉诊断

默认流程不再把 Word 转成 PDF 做视觉对照，验收依据是 Word 文字、公式、表格和图片对象的完整性以及本地确定性检查。原 PDF 仍可作为可选诊断来源；显式添加 `--visual-compare` 时，脚本才会调用 LibreOffice 渲染 Word 并进行页面版式比较：

```powershell
python scripts/quark_pdf_to_word.py `
  --input input/word/机械原理习题集.docx `
  --reference input/reference/机械原理习题集.pdf `
  --compare sample `
  --visual-compare
```

`--compare sample` 检查首、中、末页，速度较快；`full` 检查全部页面。原 PDF 只用于可选诊断，不会修改。视觉分数不再是默认输出的必要条件。

注意：使用 `--max-pages N` 时，为了按排版后的真实页数截取，仍需要临时执行一次 Word→PDF；完整处理且不指定页数时不会为默认校验转换 PDF。

不再提供 AI 审阅或模型配置参数。Word-first 默认不从 PDF 重新识别公式；原有公式对象和图片只做兼容清理。

## 固定排版规则

- 中文统一宋体（SimSun），拉丁字母和数字统一 Times New Roman，文字颜色为纯黑。
- 标题字号统一且大于正文；正文统一字号，清除原转换器产生的左右缩进。
- 前言、目录、章节、年份试题和“同步训练题答案”等大标题自动分页并居中；目录条目按普通正文顶格，文字和页码保持同一行，不使用引导线或目录表格。
- 清除页眉和页脚，避免 WPS/Word 字段差异和空白页。
- 公式、图片作为独立对象保留。公式不做语义推理；不可靠的局部公式/符号按安全规则自动替换为小图，禁止整页图片化。

## 依赖

主流程只需要本地 Python 文档处理依赖和 LibreOffice（用于 `.doc` 转换及可选视觉比对）。PDF-Extract-Kit/PaddleOCR 已安装在隔离环境 `.venv-pdfkit`，不会覆盖 MinerU 的 `.venv-mineru`；对应锁定依赖见 `requirements-pdfkit.txt`。

```powershell
python -m pip install -r requirements.txt
```

如需重新建立公式/版面检测环境：

```powershell
python -m venv .venv-pdfkit
.\.venv-pdfkit\Scripts\python.exe -m pip install -r requirements-pdfkit.txt
```

LibreOffice 可执行文件默认从系统 PATH 或 `C:\Program Files\LibreOffice\program` 查找。

## LibreOffice 启动错误 `0xc0000022`

如果 LibreOffice 在启动阶段返回 `0xc0000022`，先处理 Windows 中检测到的旧迅雷驱动 `XLGuard.sys`。请用管理员 PowerShell 运行项目脚本：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
& .\scripts\repair_windows_lo.ps1
```

脚本会停止并删除 `XLGuard` 服务/驱动，然后运行 DISM 和 SFC。完成后重启 Windows，再运行文档处理脚本。日志保存在 `tmp/windows-repair.log`。如果脚本仍提示拒绝访问，说明当前 PowerShell 没有以管理员身份运行。

## 目录

- `input/word/`：待整理的 Word 文件
- `input/reference/`：可选的原始 PDF/图片，仅用于比对
- `output/word/`：整理后的 Word
- `output/reports/`：文字、视觉和格式化检查报告
- `scripts/quark_pdf_to_word.py`：本地 Word 处理入口
- `scripts/docx_sanitizer.py`：最终 OOXML 兼容清理和结构校验
- `tmp/word/`：运行期间的临时文件（自动清理）
