---
name: ruanzhu
description: 软件著作权（软著）说明书生成与版式复刻工作流。根据求解器代码（MATLAB/Python 等）撰写软著说明书 md，逐项审计代码与文稿一致性（控制方程、数值格式、默认参数、结果数字），再以既有软著 docx 为模板包复刻版式生成新 docx（styles/标题层级/公式段制表位/三线表/题注/参考文献），公式以灰色 LaTeX 占位供 MathType 批量转换。当用户提到：软著、软件著作权、说明书、代码写说明书、说明书md转docx、版式对齐、MathType、公式占位、ruanzhu、code-to-manual 时使用。
---

# ruanzhu —— 软著说明书生成工作流

从求解器代码产出一份"与代码严格对应、版式与既有软著 word 一致、公式可批量转 MathType"的说明书 docx。

## 工作流总览（五步）

1. **读模板**：解析既有软著 docx（模板）内部结构，拿到样式 ID、版式参数
2. **写 md**：按标准章节结构撰写说明书 md，公式编号 (2-1)~(2-n)、(3-1)，图表占位与模板对齐
3. **审计**：逐项核对 md 声明 ↔ 代码实现，输出对应表，修正不一致处
4. **转换**：写内容 DSL 文件 + 调 `scripts/build_manual_docx.ps1` 生成 docx（模板包替换法）
5. **交接**：告知用户剩余人工步骤（MathType 转 36 条公式、贴 22 处截图）

## 第一步：解析模板 docx（不要用 COM 摸样式）

把模板 docx 复制一份解压（zip），读取 `word/document.xml` 与 `word/styles.xml`，确认：

| 需要确认的要素 | 典型值（中文 Word/WPS 模板） |
|---|---|
| 标题 1/2/3 样式 ID | `1` / `2` / `ad`（`ad` 的 name 可能是 Title/标题3） |
| 无编号章节标签 | `a8`（List Paragraph，常配 numPr numId） |
| 公式段样式 | `af`（"对齐"：居中制表位 4536 + 右制表位 9072） |
| 表格样式 | `a9`（Table Grid），三线表需覆盖 insideV=none |
| 参考文献样式 | `EndNoteBibliography`（或普通段落 + TNR/宋体 sz21） |
| 正文默认 | Normal：TNR+宋体、sz 24（12pt）、行距、首行缩进 |
| 页面 sectPr | 直接从模板 body 尾部原样复制（含页眉页脚 rId） |
| 公式实现方式 | MathType OLE 对象（w:object + Equation.DSMT4）或 OMML |

**核心决策**：公式是 MathType OLE → 新文档公式一律做成"af 样式段 + 左右制表位 + 编号"，中间填灰色 LaTeX 占位文本，用户用 MathType Toggle TeX 转换。绝不要试图用 COM 复刻 MathType 对象。

## 第二步：说明书 md 结构标准（与模板逐节对应）

```
标题（软件名 V1.0）
设计背景（无编号，2-4 段 + 图1-1 占位 + 引文标记[1][2][3][4]）
2. 设计流程
   2.1 模型基本假定（概化导语 + (1)~(5)条）
   2.2 xx坐标系下渗流模型（2.2.1 质量守恒 / 2.2.2 达西定律 / 2.2.3 控制方程 / 2.2.4 速度场）
   2.3 yy运移控制方程（2.3.1 守恒方程 / 2.3.2 弥散系数 / 2.3.3 滞留模型 + 耦合方程组）
   2.4 工程应用场景描述（井源项、初始条件、边界条件逐条列公式）
   2.5 数值离散与求解流程（2.5.1 隐式离散 / 2.5.2 源项边界 / 2.5.3 运移格式 / 2.5.4 CFL判据 / 质量台账）
   2.6 模型求解流程图（图2-1）
3. 软件计算精度验证（3.1 解析解描述+式(3-1)+表3-1 / 3.2 图3-1+结果段）
4. 软件应用算例（4.1 工程叙事+表4-1+图4-1+结果图占位4-x(a-d) / 4.2 三段式总结）
5. 运行环境（5.1 硬件 / 5.2 软件）
6. 软件功能及操作方法（6.1 参数配置+图6-1 / 6.2 仿真引擎流程 / 6.3 可视化+图6-2）
参考文献：
```

规则：正文用连续学术语段（不用 bullet 列表堆砌）；每个公式独立编号行；"式中，…"逐符号解释段紧跟公式。

## 第三步：代码↔文稿一致性审计（核心价值）

逐项核对并输出对应表（✅/❌），必查清单：

- [ ] 每条控制方程 ↔ 代码头注注释/实现（符号、量纲、正负号）
- [ ] 离散格式 ↔ 实际实现（隐式/显式、限制器公式方向、调和/算术平均、保守形式通量差）
- [ ] 隐式解的代数推导 ↔ 代码分母分子逐项核对（易漏乘系数）
- [ ] CFL 判据各项 ↔ 代码 min() 列表
- [ ] 边界条件 ↔ 代码实际做法（Dirichlet vs 通量型源，务必看代码注释里的"【修正】"段）
- [ ] 表格参数 ↔ GUI **启动默认值**（不是复位按钮的值！两套默认可能不同）
- [ ] 单位换算链（GUI 显示单位 × 换算系数 = solver 内部单位，如 g/cm³→kg/m³、1/h→1/s）
- [ ] 结果数字（质量误差、残差、峰值浓度等）必须来自实际运行，不得凭空写
- [ ] 非均质场类型/参数 ↔ 生成器代码（谱指数维数约定等）

## 第四步：md → docx（模板包替换法）

**原理**：复制模板 docx → 在 ZipArchive Update 模式下仅替换 `word/document.xml`（可选更新 `docProps/core.xml` 标题）。styles.xml/numbering/页眉页脚/字体全部原样继承，版式对齐是构造性的。

**内容 DSL**（UTF-8 文本文件，`|` 分隔，行首标签）：

```
TITLE|软件名 V1.0        居中加粗 sz32 大标题
H1A|设计背景             无编号一级章节（a8 + numPr numId=1 + 加粗 sz30）
H1|2. 设计流程           样式"1"
H2|2.1 xx                样式"2"
H3|2.2.1 xx              样式"ad"
BLAB|水头边界条件         a8 + numPr numId=8（自动编号小标签）
BLIT|2.xx                a8 无编号（字面编号）
B|正文段落               行距400exact+首行480+szCs24；内联标记：
                         [i]..[/i]斜体 [sb]..[/sb]下标 [sp]..[/sp]上标 [b]..[/b]加粗
EQN|(2-1)|latex          公式段：tab + 灰色斜体CambriaMath的LaTeX占位 + tab + 编号
                         latex 可省略 → 纯留白；内容文件里 LaTeX 不要含 | 字符
FIG|图1-1 xx             空居中段（贴图占位）+ 居中题注 sz21
TCAP|表3-1 xx            表注（ind 420 + 居中 + sz21）
TBLH|a##b##c##d          表头行；TBLR|.. 数据行；列分隔符 ##；单元格支持内联标记
REFH|参考文献：            样式"2"
REF|1.Wu, J., ...        EndNoteBibliography 样式、TNR/宋体 sz21、左对齐（无域代码）
```

**调用**（模板 docx 路径、DSL 路径、输出路径、文档标题）：

```powershell
powershell -ExecutionPolicy Bypass -File <skill>\scripts\build_manual_docx.ps1 `
  -Template "C:\path\模板.docx" -Content "C:\path\content.txt" `
  -Out "C:\path\输出.docx" -Title "三维xx仿真平台V1.0"
```

脚本自动：提取模板根标签与 sectPr → 构建正文 XML → [xml] 良构校验 → 替换包内 document.xml/core.xml → Word COM 试开验证（段落/表格数、关键文本）。

## 关键坑位速查（血泪教训）

| 坑 | 规则 |
|---|---|
| PowerShell .ps1 含中文 | **脚本内容必须纯 ASCII**；无 BOM UTF-8 的中文 .ps1 会被按 ANSI 读出乱码导致解析错乱。中文一律放独立 UTF-8 数据文件，或用 `[char]0xXXXX` 编码 |
| Word COM 设样式 | 中文版 Word 下英文名样式（"Heading 1"/"Table Grid"）报"项不存在"；**放弃 COM 设样式，直接 XML 构建** |
| pandoc/python-docx | 不可假设已安装；模板包替换法零依赖 |
| LaTeX 占位 | 禁用 amsmath 环境（`aligned`/`cases`，MathType 转换器不支持）；用 `\begin{array}{l}` / `{ll}`；无 `$` 定界符 |
| OMML 手写 | 除非必要不走 OMML（公式对象统一交给 MathType） |
| GUI 双套默认值 | 启动 createComponents 的默认值 ≠ 复位按钮的默认值，文档以启动值为准 |
| 表格三线表 | tblStyle 用模板的 Table Grid ID + insideV=none；表头 top12+bottom12、首数据行 top12、末行 bottom12、首列 left=nil；格内 sz21 居中 |
| 输出文件被占用 | Word 开着文档时 Copy-Item 报 IOException；提示用户关闭后重试 |
| MATLAB 中文路径 | bash 命令行传中文 OK；MATLAB -batch 参数含空格会被拆断 → 用 bat 文件承载；路径中文用 `char([36719 33879])` 式编码；8.3 短路径可能未启用 |
| 编号沿用模板列表 | H1A/BLAB 引用模板 numbering.xml 的 numId（1、8 等），复用模板包即自动生效 |

## 交付话术

生成后告知用户：
1. N 条公式 LaTeX 占位（灰色斜体）→ MathType Toggle TeX（Alt+\）逐条转换
2. M 处图占位空行（题注已就位）→ 粘贴软件截图
3. 文件首次保存后 Word 会自动清理模板遗留的未引用媒体，体积会变小

## 本次会话验证过的实例

- 模板：2D 软著说明书 docx（MathType 公式、三线表、EndNote 文献）
- 产出：3D 说明书 docx（36 公式 LaTeX 占位、2 张表、22 图占位、453 段落）
- 脚本：`scripts/build_manual_docx.ps1`；DSL 示例：`scripts/content_template.txt`
