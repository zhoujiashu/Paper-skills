# Paper Skills

Three self-contained ZCode skills for research-writing and document generation. Each skill is a `SKILL.md` prompt plus dependency-free PowerShell scripts that build `.docx` packages directly — no Python, pandoc, or `python-docx` required.

| Skill | What it does | Output |
|---|---|---|
| [`ruanzhu`](ruanzhu/) | 软件著作权说明书生成与版式复刻：读模板 docx → 写 md → 代码↔文稿审计 → 生成 docx | 软著说明书 docx |
| [`zhuanli`](zhuanli/) | 发明专利/实用新型申请文件撰写：先权利要求后说明书，代码↔文稿审计，标准版式或模板复刻 | 说明书 + 权利要求书（+ 摘要）docx |
| [`lunwen`](lunwen/) | AGU/WRR 期刊论文骨架与填空：标题/作者/Key Points/Abstract/正文/附录全就位，公式与图表留位 | 论文骨架 docx |

## 共同特点

- **零依赖**：直接构造 OOXML 包，只用 Windows PowerShell 5.1+ 与 .NET 自带压缩库
- **公式可批处理**：公式一律做成灰色 LaTeX 占位，成稿时用 MathType 的 Toggle TeX（Alt+\）逐条转换
- **版式可复刻**：ruanzhu / zhuanli 支持"模板包替换法"，继承既有 docx 的 styles、页眉页脚与 sectPr
- **审计优先**：ruanzhu 与 zhuanli 内置代码↔文稿一致性清单；lunwen 内置公式推导七步协议与 `check_fill.ps1` 占位检查
- **纯 ASCII 脚本**：中文全部放在 UTF-8 数据文件里，避免 PowerShell 编码坑

## 环境要求

- Windows 10 / 11
- PowerShell 5.1 或更高
- Microsoft Word（可选，仅用于构建后的 COM 试开验证）
- MathType（可选，仅用于把 LaTeX 占位转成公式对象）

## 安装

把需要的 skill 文件夹复制到你的 agent 技能目录，例如：

```
C:\Users\<you>\.agents\skills\ruanzhu
C:\Users\<you>\.agents\skills\zhuanli
C:\Users\<you>\.agents\skills\lunwen
```

重新加载技能后，用 `/ruanzhu`、`/zhuanli`、`/lunwen` 触发。

## 快速开始

### ruanzhu —— 软著说明书

```powershell
powershell -ExecutionPolicy Bypass -File ruanzhu\scripts\build_manual_docx.ps1 `
  -Template "C:\path\模板.docx" -Content "C:\path\content.txt" `
  -Out "C:\path\输出.docx" -Title "软件名 V1.0"
```

### zhuanli —— 专利申请文件

标准版式（不需要模板）：

```powershell
powershell -ExecutionPolicy Bypass -File zhuanli\scripts\build_patent_docx.ps1 `
  -Content "C:\path\content.txt" -OutDir "C:\path\out" -BaseName "一种xxx方法"
```

复刻既有专利稿：

```powershell
powershell -ExecutionPolicy Bypass -File zhuanli\scripts\build_patent_docx.ps1 `
  -Template "C:\path\专利模板.docx" -Content "C:\path\content.txt" `
  -OutDir "C:\path\out" -BaseName "一种xxx方法"
```

先解析模板结构：

```powershell
powershell -ExecutionPolicy Bypass -File zhuanli\scripts\inspect_docx.ps1 -Path "C:\path\专利模板.docx"
```

### lunwen —— 期刊论文骨架

```powershell
# 生成填空骨架
powershell -ExecutionPolicy Bypass -File lunwen\scripts\build_paper_skeleton.ps1 `
  -Out "C:\path\paper_skeleton.docx"

# 填完检查剩余占位、GUIDE 段与字数
powershell -ExecutionPolicy Bypass -File lunwen\scripts\check_fill.ps1 `
  -Path "C:\path\paper_filled.docx"
```

`lunwen` 默认按 AGU/WRR 投稿版式生成：Letter 纸、1 英寸边距、连续行号、Times New Roman 12 pt、双倍行距、页脚页码；待填内容用黄色 `[[...]]` 高亮，写作提示用灰色 `GUIDE` 段，公式符号解释用顶格小写的 `where ...`。

## 说明

- 仓库只包含 skill 提示词、脚本与空白模板，**不包含任何论文、专利或软著正文内容**。
- 生成文件里的公式是 LaTeX 占位，不是可直接提交的公式对象；需用 MathType 转换。
- `zhuanli` 与 `lunwen` 不替代专利代理师、同行评审或新颖性检索；正式提交前请人工核对著录项、数据与声明。

## License

[MIT](LICENSE)
