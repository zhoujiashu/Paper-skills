---
name: lunwen
description: 期刊论文骨架与填空工作流（默认 AGU/WRR 投稿版式，英文）。按目标期刊搭出 Title/Authors/Key Points/Abstract/Plain Language Summary/Introduction/Governing Equations/Methodology/Results/Discussion/Conclusions/References/Appendix 的 docx 骨架，所有待填内容以黄色 [[...]] 占位、写作提示以灰色 GUIDE 段给出，作者只需逐项替换；公式推导与数值/解析求解部分按 ruanzhu 的推导协议补全（假定—守恒—本构—耦合消元—无量纲化—离散/解析—验证—代码审计），公式以灰色 LaTeX 占位供 MathType 转换。当用户提到：论文、paper、manuscript、期刊论文、AGU、WRR、Water Resources Research、论文骨架、论文模板、填空、公式推导、Appendix、投稿稿 时使用。
---

# lunwen —— 期刊论文骨架与填空工作流

目标：先用目标期刊版式把论文"搭好架子"，作者只填内容；涉及公式推导与求解的部分，按 ruanzhu 的严格度补全，而不是留一句"详见附录"。

**分工**：`paper-writer` 负责把材料写成正文；`zhuanli` 负责专利版式；本 skill 负责**期刊骨架 + 版式 + 公式推导/求解 + 回填检查**。三者可串联：先 `lunwen` 搭骨架，用 `paper-writer` 填正文，公式段落按 ruanzhu 协议写。

## 工作流（七步）

1. **定叙事**：确定研究问题、贡献点、目标期刊（默认 AGU/WRR）
2. **搭骨架**：`build_paper_skeleton.ps1` 生成填空 docx（黄色 `[[...]]` 占位 + 灰色 GUIDE）
3. **填正文**：按章节规范逐段替换；每个 `[[...]]` 替换后去掉双方括号与高亮
4. **补公式**：按"公式推导协议"逐步推导，填入 `EQ` 行；数值解/解析解各自补齐验证
5. **审计**：公式 ↔ 代码/推导、符号 ↔ 单位、数字 ↔ 实际结果逐项核对
6. **回填**：用填好的 content.txt 重新生成 docx
7. **检查**：`check_fill.ps1` 报告剩余占位、GUIDE 段、字数超限

## 第一步：生成填空骨架

```powershell
powershell -ExecutionPolicy Bypass -File <skill>\scripts\build_paper_skeleton.ps1 `
  -Out "C:\path\paper_skeleton.docx"
```

- 默认读取 `scripts/content_skeleton.txt`；也可 `-Content "自己的 content.txt"`
- `-Journal "Water Resources Research"` 控制页眉
- `-SkipValidation` 跳过 Word 试开
- `[[...]]` 自动黄色高亮；以 `GUIDE` 开头的段落为灰色斜体提示
- 生成后输出公式数、表格数、Key Points 数、摘要/PLS 字数

### 内容 DSL（UTF-8，`|` 分隔，脚本纯 ASCII）

```
TITLE|标题                         14pt 加粗，左对齐
AUTH|作者[sp]1,2[/sp]               作者行（支持上标）
AFF|1单位                           单位行
CORR|Corresponding author: ...      通讯作者行
H1|2 Governing Equations            一级标题：12pt 加粗
H2|2.1 Governing equations          二级标题：12pt 非加粗、左缩进 0.5"
B|正文段落                          双倍行距、两端对齐、首行缩进 0.5"
BQ|无缩进段落                       用于摘要式短段或公式解释
W|where ...                        公式符号解释段：顶格、无首行缩进、where 小写
KEY|Key point                        Key Points 样式（3 条，每条 ≤100 字符）
ABS|摘要正文                        Abstract 样式（≤250 words）
PLS|通俗摘要                        Plain Language Summary（100–200 words）
NOTE|GUIDE: 写作提示                灰色斜体，提交前删除
EQ|(1)|latex                        公式段：居中制表位 + 灰色 LaTeX + 右侧编号
FIG|Fig. 1. 图题                    空行（贴图）+ 图题
TCAP|Table 1. 表题                  表题
TBLH|a##b##c                       表头；TBLR|.. 数据行；## 分列
REF|作者, A. A. (年). 题名...       参考文献，悬挂缩进
PAGEBREAK                           分页
```

行内标记：`[i]..[/i]` 斜体、`[sb]..[/sb]` 下标、`[sp]..[/sp]` 上标、`[b]..[/b]` 加粗、`[[..]]` 黄色占位。LaTeX 中不要使用 `aligned`/`cases` 等 amsmath 环境，用 `\begin{array}{l}`。

## 第二步：章节写作规范

| 章节 | 段落/篇幅 | 必须包含 |
|---|---|---|
| Title | ≤15 words | 系统 + 方法 + 结果，避免 "Study on" 式空话 |
| Key Points | 3 条，每条 ≤100 字符 | 结论性发现，不写方法 |
| Abstract | 1 段，≤250 words | 背景 1–2 句、缺口、方法、定量结果、意义；不引用文献 |
| Plain Language Summary | 100–200 words | 无术语、无缩写、无引用；回答"为什么普通人要在意" |
| 1 Introduction | 5–6 段 | 背景 → 已知 → 缺口 → 2–3 个技术挑战 → 贡献（对应结果节）→ 路线图 |
| 2 Governing Equations | 4 小节 | 假定、守恒、本构、耦合消元、边界/初始、无量纲化；每个公式后紧跟符号解释 |
| 3 Methodology | 3 小节 | 可复现的离散/损失函数、算法参数、稳定性/收敛判据、参考解与误差指标 |
| 4 Results | 3–4 小节 | 先验证后应用；每个结论对应图或表；报告误差量级 |
| 5 Discussion | 2 小节 | 物理解释、与已有工作对比、局限与敏感性 |
| 6 Conclusions | 4 段 | 做了什么、验证了什么、发现了什么、下一步；不引入新结果 |
| Appendix | 按需 | 详细推导、尺度分析、符号表；公式编号 (A1)… |

## 第三步：公式推导协议（按 ruanzhu 要求）

推导不是把结果抄上去，而是让读者能从假定走到方程。每一步落到 `EQ` 行，公式后紧跟 "where ..." 符号解释。

### D1 假定
逐条列出并说明理由：维度与几何、饱和/非饱和、小应变/大应变、等温、材料相数、是否考虑吸附/降解。每条假定对应后文一个可检验的极限或适用范围。

### D2 守恒方程
写出质量守恒、动量/平衡、溶质质量守恒的积分或微分形式；明确通量定义与控制体。

### D3 本构关系
达西定律、胡克定律、Fick 定律、吸附等温式等逐条给出，并说明系数含义与量纲。

### D4 耦合与消元
把本构代入守恒，消去辅助场（如位移），得到降维/降场系统。**中间步骤不能跳**：积分、代入、整理各占一行公式；附录给完整推导。

### D5 无量纲化
列出特征尺度（H、T_ref、p0、c0），代入后收集无量纲群；全文此后统一使用无量纲记号。

### D6 求解
- **数值解**：离散格式（FDM/FEM/PINN）、时间推进、稳定性或 CFL 条件、网格/网络规模、迭代与收敛判据、边界与初始条件的实现方式
- **解析解**：分离变量/Laplace 变换/特征函数展开的每一步；级数收敛性；退化为已知极限的检验
- **学习型方法**：残差形式、损失函数各项权重、配点数、训练阶段、接口条件（域分解时）

### D7 验证
解析解对比、网格收敛阶、质量/能量守恒误差、与文献 benchmark 的偏差；给出误差定义式与数值。结果数字必须来自实际运行，不得编造。

### 公式书写规则

- 每个独立公式一行，编号连续：(1)…(n)；附录用 (A1)…
- 公式后的符号解释单独成段，用 `W|where ...`：**顶格（无首行缩进）**，`where` **一律小写**；脚本会自动把句首 `Where` 改成 `where`
- 公式后紧跟一段 "where ..." 逐符号解释，含单位
- 符号表与正文一致；同一符号不得两种含义
- 公式中不使用未定义缩写；下标用 `[sb]`，上标用 `[sp]`
- 先用灰色 LaTeX 占位，成稿时 MathType Toggle TeX（Alt+\）转换

## 第四步：代码 ↔ 文稿审计（沿用 ruanzhu）

- [ ] 每条控制方程 ↔ 代码头注/实现（符号、量纲、正负号）
- [ ] 离散格式 ↔ 实际实现（隐式/显式、限制器方向、平均方式、通量差）
- [ ] 隐式解的代数推导 ↔ 代码分子分母逐项核对
- [ ] CFL/稳定性判据各项 ↔ 代码 `min()` 列表
- [ ] 边界条件 ↔ 代码实际做法（Dirichlet vs 通量型）
- [ ] 表格参数 ↔ 代码**启动默认值**（不是复位值）
- [ ] 单位换算链（显示单位 × 系数 = 内部单位）
- [ ] 结果数字（误差、峰值、收敛阶）↔ 实际运行输出
- [ ] 图/表编号 ↔ 正文引用顺序
- [ ] 摘要/结论中的数字 ↔ 结果节

## 第五步：回填与检查

填完后重新生成：

```powershell
powershell -ExecutionPolicy Bypass -File <skill>\scripts\build_paper_skeleton.ps1 `
  -Content "C:\path\filled_content.txt" -Out "C:\path\paper_filled.docx"
```

检查剩余占位与字数：

```powershell
powershell -ExecutionPolicy Bypass -File <skill>\scripts\check_fill.ps1 `
  -Path "C:\path\paper_filled.docx"
```

输出剩余 `[[...]]` 数、GUIDE 段数、Key Points 条数与最大字符数、Abstract/PLS 词数，并对超限项给出 WARNING。

## AGU/WRR 版式规格（内置）

| 要素 | 值 |
|---|---|
| 纸张 | Letter（8.5×11 in） |
| 页边距 | 四边 1 in |
| 行号 | 连续编号，每行 |
| 页眉 | `manuscript submitted to <Journal>`，10pt 居中 |
| 正文 | Times New Roman 12pt，双倍行距，两端对齐，首行缩进 0.5" |
| 标题 | 14pt 加粗，左对齐 |
| 一级标题 | 12pt 加粗，段前 12pt |
| 二级标题 | 12pt 非加粗，左缩进 0.5" |
| 公式 | 居中制表位 4680 + 右制表位 9360，编号右对齐 |
| 图表题 | 12pt，表题在上、图题在下 |
| 参考文献 | 12pt，悬挂缩进 0.5"，作者-年份，字母序 |

## 关键坑位

| 坑 | 规则 |
|---|---|
| 把骨架当正文 | `[[...]]` 与 `GUIDE` 段提交前必须全部清除，用 `check_fill.ps1` 兜底 |
| 公式只写结果 | 每条公式前有推导动机、后有符号解释；附录补完整步骤 |
| 符号不一致 | 正文、公式、图表、附录共用一张符号表；同一符号只表示一个量 |
| 编造结果 | 摘要与结论的数字必须能在结果节和运行日志中找到 |
| 直接复用他人稿件 | 版式可复刻，文字与数据不可照搬；本 skill 只使用版式参数 |
| MathType | LaTeX 占位为灰色斜体，成稿时逐条 Toggle TeX；不要手写 OMML |
| 字数超限 | Key Points 每条 ≤100 字符；Abstract ≤250 words；PLS 100–200 words |
| 图表顺序 | 正文首次引用顺序 = 图/表编号顺序 |
| 页眉期刊名 | 用 `-Journal` 修改，默认 Water Resources Research |

## 交付话术

1. 生成了填空骨架 docx：黄色 `[[...]]` 为待填项，灰色 `GUIDE` 为写作提示，提交前删除
2. 公式为灰色 LaTeX 占位，成稿时 MathType Toggle TeX；附录公式编号 (A1)…
3. 已按 AGU/WRR 版式：Letter、1 in 边距、连续行号、TNR 12pt、双倍行距
4. 填完后跑 `check_fill.ps1`，确认 0 个占位、0 个 GUIDE、字数合规
5. 本 skill 不替代同行评审与数据可得性审查；投稿前仍需核对作者贡献、基金与伦理声明

## 本次会话验证过的实例

- 版式来源：用户提供的 WRR/AGU 投稿稿（526 段、56 个 MathType 公式、6 表、连续行号）
- 实测版式：Letter 612×792 pt、四边 72 pt、行号开启、页眉正确、标题 14 pt 加粗、Text 12 pt 双倍行距
- 骨架构建：165 段、19 个公式位、2 张表、3 条 Key Points、13 条 GUIDE，Word 实际打开成功
- `check_fill.ps1`：骨架报 143 个占位 / 13 条 GUIDE；干净稿报 `OK: no placeholders or guides left`
- 脚本：`scripts/build_paper_skeleton.ps1`、`scripts/check_fill.ps1`
- 骨架：`scripts/content_skeleton.txt`（含 14 个主公式位、5 个附录公式位、图表位与参考文献位）
