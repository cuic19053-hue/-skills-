# 大学生申报书/申请书制作 Skill 集合

[![Agent Skills](https://img.shields.io/badge/Agent_Skills-Compatible-6e3bf0)](https://agentskills.io/specification)
[![Platform](https://img.shields.io/badge/Platform-Claude_Code_|_Codex_|_Cursor_|_Copilot_|_Gemini_CLI-blue)]()
[![License](https://img.shields.io/badge/License-MIT-green)](./LICENSE)
[![Install](https://img.shields.io/badge/install-npx_skills_add_cuic19053--hue/-skills---black)](https://github.com/evaisse/npx-skills)
[![Stars](https://img.shields.io/github/stars/cuic19053-hue/-skills-?style=social)](https://github.com/cuic19053-hue/-skills-)
[![Subskills](https://img.shields.io/badge/Subskills-35-blue)](#能写哪些申报书)
[![Content](https://img.shields.io/badge/Content-2.4MB_|_42000+_lines-orange)]()

> 📊 **项目数据看板**
>
> | 指标 | 数值 | 说明 |
> |---|---|---|
> | 覆盖赛道 | **35 个** | 8 大类，从大创到公派留学全覆盖 |
> | 总内容量 | **2,400 KB** | 42,000+ 行结构化领域知识 |
> | 平均深度 | **68 KB / 赛道** | 最深 116 KB（大创创新训练） |
> | 最近更新 | 2026-07 | v2.1 版本 |
> | Star 目标 | 🎯 **1,000** | 当前向目标推进中 |
>
> ⭐ **如果这个项目帮到了你，点个 Star 让更多同学看到。Star 数越多，越能吸引赛道专家贡献内容，最终受益的是所有使用者。**

一套覆盖中国大学生常见申报场景的 Agent Skills。**35 个子 skill**，从大创立项到入党申请书，从奖学金到保研推免，从应征入伍到公派留学，让 AI 编程助手变成你身边那个"帮学弟学妹改过几十份申报书"的学长。

> 🔔 **不需要克隆、不需要装环境、不需要会编程。**
> 找到你需要的 [子 skill](./subskills/README.md)，复制 `SKILL.md` 内容粘贴到豆包/DeepSeek/Claude，就能用。下方有[保姆级教程](#零基础使用指南文科生--不会编程也能用)。

---

## 为什么做这个

中国大学生每年要写多少份申报书？大创、奖学金、三下乡、入党、保研——每一份都要求格式规范、语气恰当、事实准确。但大多数人要么对着空白模板发愁，要么去网上搜一些不知所云的范文。

ChatGPT 可以帮你写，但通用模型不知道"大创"是什么、不知道国家奖学金评审看什么、不知道入党申请书有查重。

这个项目把**高校申报书的领域知识**打包成了 Agent Skills——不是给聊天框里贴一段 prompt，而是让你的 AI 编程助手（Claude Code、Codex、Cursor 等）在需要时自动加载对应的专业知识。

## 能写哪些申报书

**8 大类 · 35 个子 skill · 2,400 KB 结构化领域知识**

> 状态说明：✅ 已完备（可直接使用）｜ 🔄 持续打磨中｜ 📅 Star 解锁（达到对应 Star 数后启动）

### 🔬 项目类（要钱、要立项）— 8 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 1 | 大创 · 创新训练项目 | `innovation_research` | 116 KB | ✅ |
| 2 | 大创 · 创业训练项目 | `entrepreneurship_training` | 112 KB | ✅ |
| 3 | 大创 · 创业实践项目 | `entrepreneurship_practice` | 52 KB | ✅ |
| 4 | 校级科研立项 | `university_research` | 101 KB | ✅ |
| 5 | 院级科研立项 | `college_research` | 98 KB | ✅ |
| 6 | 挑战杯 · 课外学术科技作品 | `challenge_cup` | 102 KB | ✅ |
| 7 | 互联网+ · 商业计划书 | `internet_plus` | 93 KB | ✅ |
| 8 | 互联网+ · 红旅赛道 | `internet_plus_red_tour` | 79 KB | ✅ |

### 🏆 评优类（要荣誉）— 12 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 9 | 国家奖学金（8000 元） | `national_scholarship` | 39 KB | ✅ |
| 10 | 国家励志奖学金（5000 元） | `motivation_scholarship` | 43 KB | ✅ |
| 11 | 校级奖学金 | `university_scholarship` | 40 KB | ✅ |
| 12 | 企业专项奖学金 | `enterprise_scholarship` | 47 KB | ✅ |
| 13 | 单项奖学金 | `single_scholarship` | 44 KB | ✅ |
| 14 | 国家助学金 | `grant_application` | 75 KB | ✅ |
| 15 | 优秀毕业生 | `outstanding_graduate` | 46 KB | ✅ |
| 16 | 优秀学生 / 三好学生 | `outstanding_student` | 58 KB | ✅ |
| 17 | 优秀学生干部 | `outstanding_cadre` | 87 KB | ✅ |
| 18 | 文明大学生 / 优秀团员 | `civilized_student` | 60 KB | ✅ |
| 19 | 优秀班集体 | `class_collective` | 66 KB | ✅ |
| 20 | 优秀毕业设计 / 论文 | `outstanding_thesis` | 73 KB | ✅ |

### 🌾 活动类 — 5 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 21 | 三下乡 · 社会调查 | `social_survey` | 46 KB | ✅ |
| 22 | 三下乡 · 支教 | `volunteer_teaching` | 48 KB | ✅ |
| 23 | 三下乡 · 政策宣讲 | `policy_lecture` | 54 KB | ✅ |
| 24 | 三下乡 · 科技服务 | `tech_service` | 63 KB | ✅ |
| 25 | 西部计划 | `western_plan` | 74 KB | ✅ |

### 🚩 政治身份类 — 4 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 26 | 入党申请书 | `party_application` | 58 KB | ✅ |
| 27 | 思想汇报 | `thought_report` | 51 KB | ✅ |
| 28 | 转正申请书 | `party_full_member` | 90 KB | ✅ |
| 29 | 入团申请书 | `youth_league_application` | 72 KB | ✅ |

### 🎓 升学类 — 2 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 30 | 保研推免申请 | `graduate_recommendation` | 43 KB | ✅ |
| 31 | 选调生申请 | `selected_graduate` | 68 KB | ✅ |

### ✈️ 公派留学 / 交流类 — 2 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 32 | CSC 国家公派留学 | `csc_scholarship` | 65 KB | ✅ |
| 33 | 交流项目申请 | `exchange_program` | 89 KB | ✅ |

### 📋 其他 — 2 个

| # | 赛道 | 编码 | 内容量 | 状态 |
|---|---|---|---|---|
| 34 | 转专业申请 | `major_transfer` | 78 KB | ✅ |
| 35 | 应征入伍申请书 | `military_enlistment` | 73 KB | ✅ |

---

### 📅 Star 解锁路线图

> 以下赛道已规划，达到对应 Star 数后启动开发。**点 Star = 催更。**

| Star 里程碑 | 解锁赛道 | 说明 |
|---|---|---|
| ⭐ 100 | 讲莲杯、创青春、三创赛 | 竞赛类扩展 |
| ⭐ 300 | 节能减排大赛、生命科学竞赛 | 学科专项类 |
| ⭐ 500 | 服务外包大赛、商业精英挑战赛 | 商科类扩展 |
| ⭐ 1000 | 师范生教学技能、英语类竞赛 | 教师资格 + 外语类 |

## 项目结构

```
college-application-doc/
├── SKILL.md                 # 路由入口：识别类型 -> 分流到子 skill
├── AGENT_PROMPT.md          # Agent 使用提示词与工作流
├── CASE_ANALYSIS.md         # 案例分析文档
├── SUBAGENT_TEMPLATE.md     # 子代理模板
├── index.json               # 机器可读子 skill 索引
├── version.json             # 项目版本元数据
├── README.md                # 你正在看的
├── LICENSE                  # MIT
├── subskills/               # 35 个子 skill，每个独立文件夹
│   ├── innovation_research/SKILL.md           # 大创-创新训练
│   ├── entrepreneurship_training/SKILL.md     # 大创-创业训练
│   ├── entrepreneurship_practice/SKILL.md     # 大创-创业实践
│   ├── university_research/SKILL.md           # 校级科研立项
│   ├── college_research/SKILL.md              # 院级科研立项
│   ├── challenge_cup/SKILL.md                 # 挑战杯
│   ├── internet_plus/SKILL.md                 # 互联网+
│   ├── internet_plus_red_tour/SKILL.md        # 互联网+红旅赛道
│   ├── national_scholarship/SKILL.md          # 国家奖学金
│   ├── motivation_scholarship/SKILL.md        # 国家励志奖学金
│   ├── university_scholarship/SKILL.md        # 校级奖学金
│   ├── enterprise_scholarship/SKILL.md        # 企业奖学金
│   ├── single_scholarship/SKILL.md            # 单项奖学金
│   ├── grant_application/SKILL.md             # 国家助学金
│   ├── outstanding_graduate/SKILL.md          # 优秀毕业生
│   ├── outstanding_student/SKILL.md           # 优秀学生
│   ├── outstanding_cadre/SKILL.md             # 优秀学生干部
│   ├── civilized_student/SKILL.md             # 文明大学生
│   ├── class_collective/SKILL.md              # 优秀班集体
│   ├── outstanding_thesis/SKILL.md            # 优秀毕业设计
│   ├── social_survey/SKILL.md                 # 三下乡-社会调查
│   ├── volunteer_teaching/SKILL.md            # 三下乡-支教
│   ├── policy_lecture/SKILL.md                # 三下乡-政策宣讲
│   ├── tech_service/SKILL.md                  # 三下乡-科技服务
│   ├── western_plan/SKILL.md                  # 西部计划
│   ├── party_application/SKILL.md             # 入党申请书
│   ├── thought_report/SKILL.md                # 思想汇报
│   ├── party_full_member/SKILL.md             # 转正申请书
│   ├── youth_league_application/SKILL.md      # 入团申请书
│   ├── graduate_recommendation/SKILL.md       # 保研推免
│   ├── selected_graduate/SKILL.md             # 选调生申请
│   ├── major_transfer/SKILL.md                # 转专业申请
│   ├── military_enlistment/SKILL.md           # 应征入伍申请书
│   ├── csc_scholarship/SKILL.md              # CSC 公派留学
│   └── exchange_program/SKILL.md              # 交流项目申请
├── references/              # 共享知识库（所有子 skill 按需引用）
│   ├── types.md             # 8 类申报书栏目结构详解
│   ├── writing_guide.md     # 撰写规范 + 格式标准 + 信息采集清单
│   ├── review_criteria.md   # 评审打分维度
│   └── pitfalls.md          # 常见错误与避坑指南
├── scripts/
│   └── build_docx.py        # 生成 Word 文档的辅助脚本
├── utils/                   # 工程化能力模块（v2.1 新增）
│   ├── dispatcher.py        # 分流决策树（790+行）
│   ├── docx_common.py       # 通用 docx 工具
│   ├── school_template.py   # 学校模板管理
│   ├── pdf_export.py        # PDF 导出
│   ├── plagiarism_checker.py # 查重检测
│   ├── review_simulator.py  # 评审模拟
│   └── schools/             # 学校模板（清华/北大/浙大/武大）
└── examples/
    ├── demos/               # demo Word 文档（v2.1 新增）
    ├── innovation_project_example.md
    └── scholarship_example.md
```

### 路由设计

根目录的 `SKILL.md` 是**路由器**，不直接生成内容。它的职责：

1. 识别用户要写哪类申报书（关键词匹配 + 决策树）
2. 不确定时列出候选让用户确认
3. 移交到对应子 skill

每个子 skill 是独立的，有完整的 YAML frontmatter（name、description、触发条件），Agent 可以根据描述自动匹配。

## 核心设计原则

### 1. 诚实底线

这不是"AI 帮你编一份看起来光鲜的申报书"的工具。所有子 skill 都内置了诚实约束：

- **只写用户能提供证据的事实** —— 问"你拿过什么奖"，不替你列奖项
- **不放大、不润色** —— "班级第二"就是"班级第二"，不是"成绩优异名列前茅"
- **不留模糊占位** —— 拿不准的信息追问，不写"获得多项荣誉"这种空话
- **提醒复核** —— 生成后明确提示用户逐项核对真实性

大学生申报书是面向学校/教育主管部门的正式材料。造假会被记入档案甚至触发学籍处分——这是红线，Skill 里写了，README 里再说一遍。

### 2. 渐进式加载

遵循 Agent Skills 规范的三层加载机制：

- **元数据层**（~100 tokens）：启动时只加载 name 和 description
- **指令层**：Agent 匹配到对应 skill 后，才加载 SKILL.md 正文
- **资源层**：执行过程中按需加载 references/ 中的深度文档

不把 35 个 skill 全塞进上下文。

### 3. 学校差异兼容

每个学校的模板有微调（栏目顺序、字号、页边距）。Skill 的策略是：

- 优先用用户提供的学校模板
- Skill 提供**内容质量参考**（怎么写好、怎么避坑）
- 格式标准以 `references/writing_guide.md` 为准

---

## 零基础使用指南（文科生 / 不会编程也能用）

> **不需要装环境、不需要敲命令、不需要 GitHub 账号。**

### 方式 A：网页 AI 直接用（零门槛，推荐文科生）

1. 打开任意 AI 对话工具（[豆包](https://doubao.com)、[DeepSeek](https://chat.deepseek.com)、[Claude](https://claude.ai)、[ChatGPT](https://chat.openai.com)）
2. 在本仓库找到你需要的子 skill，打开对应的 `SKILL.md` 文件（如 `subskills/national_scholarship/SKILL.md`）
3. 复制文件全部内容，粘贴到 AI 对话框，发送
4. 然后跟 AI 说："我要申请国家奖学金，请按这个 skill 帮我写"

**就这么简单，不需要任何技术操作。**

> ⚠️ **注意**：AI 在聊天框里输出的是纯文本，没有 Word 排版（字体/页边距/表格线等）。如果需要格式完整的文档，建议用下方的**方式 B 直接下载模板**。

### 方式 B：直接下载成品模板（零门槛，推荐）

> 不用复制粘贴，不用跟 AI 对话，直接下载完整的 Word 文档，填入自己的信息即可。

[examples/demos/](https://github.com/cuic19053-hue/-skills-/tree/main/examples/demos) 目录存放了 10 个已经排版好的 demo 文档（含字体、表格、页边距、页眉页脚等完整格式）：

| 下载 | 类型 | 说明 |
|---|---|---|
| [📄 国家奖学金申请书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_national_scholarship.docx) | 奖学金 | 8000 元国奖，前 10% GPA |
| [📄 入党申请书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_party_application.docx) | 政治身份 | 4000 字，含入党志愿+认识 |
| [📄 挑战杯申报书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_challenge_cup.docx) | 竞赛 | 自然科学类学术论文 |
| [📄 互联网+商业计划书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_internet_plus.docx) | 竞赛 | 创业大赛，2025 新评审维度 |
| [📄 大创创新训练.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_innovation_research.docx) | 科研 | 国家级，研究报告/论文产出 |
| [📄 优秀毕业生申请书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_outstanding_graduate.docx) | 评优 | 省级+校级，四年综合表现 |
| [📄 院级科研立项.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_college_research.docx) | 科研 | SRTP，含参考文献 |
| [📄 应征入伍申请书.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_military_enlistment.docx) | 征兵 | 2025 上半年应征 |
| [📄 CSC 公派留学.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_csc_scholarship.docx) | 公派留学 | 联合培养博士研究生 |
| [📄 大创创业训练.docx](https://github.com/cuic19053-hue/-skills-/raw/main/examples/demos/demo_entrepreneurship_training.docx) | 科研 | 消防无人机项目（含 18 张表格+技术路线图） |

**使用方法**：点击链接 → 浏览器下载 → 用 Word 打开 → 把内容替换成你自己的信息。

### 方式 C：AI 编程助手一键装（低门槛）

如果你电脑上装了 [Trae](https://www.trae.cn/)、[Cursor](https://cursor.sh/)、[Claude Code](https://claude.ai/code) 等 AI 编程工具：

```bash
npx skills add cuic19053-hue/-skills-
```

装完后直接跟 AI 说"帮我写入党申请书"，它会自动加载对应 skill。

### 方式 D：直接看模板手动套用（纯手工）

1. clone 或下载本仓库
2. 进入 `subskills/` 找到对应目录
3. 看 `SKILL.md` 里的结构说明和 `examples/` 里的示例
4. 手动套用到自己的申报书上

---

## 安装（高级用户 / 开发者）

> 💡 **普通用户不需要看这一节。** 直接用上方的[零基础使用指南](#零基础使用指南文科生--不会编程也能用)即可。
> 以下内容适合：AI 编程助手用户（Trae/Cursor/Claude Code）、贡献者、需要批量生成 Word 的高级用户。

本项目遵循 [Agent Skills 开放标准](https://agentskills.io/specification)，支持 Claude Code、Codex、Cursor、GitHub Copilot、Gemini CLI、Windsurf 等所有兼容 SKILL.md 的 Agent。

### 推荐方式：npx skills（一键安装）

[`npx skills`](https://github.com/evaisse/npx-skills) 是 Agent Skills 生态的官方 CLI 工具（Vercel Labs 出品），一行命令搞定安装、更新、卸载。

```bash
# 安装全部 35 个 skill（项目级，提交到 Git 可与团队共享）
npx skills add cuic19053-hue/-skills-

# 全局安装（所有项目都能用）
npx skills add cuic19053-hue/-skills- -g

# 只装你需要的几个
npx skills add cuic19053-hue/-skills- --skill national-scholarship --skill party-application

# 指定安装到特定 Agent
npx skills add cuic19053-hue/-skills- -a claude-code -a cursor

# CI/CD 或脚本中静默安装
npx skills add cuic19053-hue/-skills- -g -a claude-code -y
```

安装后，`npx skills` 会自动把 skill 链接到你的 Agent 目录（`.claude/skills/`、`.codex/skills/` 等），Agent 启动时即可自动发现和加载。

**查看已安装的 skill：**

```bash
npx skills list              # 项目级
npx skills ls -g             # 全局
npx skills ls -a claude-code # 按 Agent 过滤
```

**更新：**

```bash
npx skills update            # 更新全部
npx skills update national-scholarship  # 更新单个
```

### 手动安装

如果你不想用 CLI 工具，也可以手动复制：

```bash
# 克隆仓库
git clone https://github.com/cuic19053-hue/-skills-.git

# Claude Code（项目级）
cp -r -skills-/subskills/* .claude/skills/

# Codex
cp -r -skills-/subskills/* .codex/skills/

# Cursor
cp -r -skills-/subskills/* .cursor/skills/

# 全局安装以 Claude Code 为例
cp -r -skills-/subskills/* ~/.claude/skills/
```

> Cursor 同时支持 `.cursor/skills/`（skills）和 `.cursor/rules/`（rules），不要把 SKILL.md 放到 rules 里，会导致重复加载。

---

## 使用方式

Skill 安装后，**不需要记命令**。直接跟 Agent 说人话就行：

```
"帮我写一份国家奖学金申请书，我是计算机专业大三的"
"我要申报大创创新训练项目，课题是无人机路径规划"
"帮我写一份入党申请书"
```

路由 skill 会自动识别类型。如果不确定，Agent 会反问你是哪种。

生成的内容是 Markdown，如果需要 Word 文档，Agent 会调用 docx 生成工具。

---

## 不是什么东西

- 不是"一键生成完美申报书"的魔法按钮 —— 你需要提供真实经历和数据
- 不是政策咨询工具 —— 不解读学校的评审规则变化
- 不保证中标 —— 申报结果取决于你的实际条件和评审偏好
- 不是学术不端工具 —— 不会帮你编造论文、奖项、经历
  

---

## 免责声明

**本项目按"现状"提供，仅供学习和参考使用。**

1. **非官方出品**：本项目与任何高校、教育主管部门、奖学金评审委员会无关。Skill 中的评审标准、格式规范基于公开信息和社区经验整理，不构成官方指导意见。
2. **内容真实性由使用者负责**：Skill 生成的申报书内容取决于你提供的信息。使用本项目产生的任何申报材料，其真实性和准确性由使用者本人承担全部责任。用虚假信息申报奖学金、科研立项等，可能面临学籍处分乃至法律责任。
3. **不保证结果**：本项目不承诺、不保证使用后一定能通过评审、获得立项、拿到奖学金。评审结果受多种因素影响，包括但不限于：你的实际条件、竞争激烈程度、评审专家的主观判断、学校政策变化。
4. **格式以学校模板为准**：不同学校下发的模板在栏目顺序、字数要求、排版规范上可能存在差异。Skill 提供的内容参考不能替代学校官方模板，提交前请对照本校要求逐项检查。
5. **政治类文书的特殊性**：入党申请书、思想汇报等政治身份类文书有严格的政治表述规范。Skill 中的参考内容不构成政治立场建议。申请人对文书中政治表述的准确性承担全部责任。
6. **责任豁免**：在法律允许的最大范围内，本项目作者及贡献者不对因使用或无法使用本项目而产生的任何直接、间接、附带、特殊或后果性损害承担责任，包括但不限于申报失败、奖学金落选、项目被拒、学籍处分或其他损失，即使已被告知可能发生此类损害。
7. **合规使用**：使用者必须遵守所在国家/地区以及所在学校的相关法律法规和规章制度。如使用行为违反相关规定，请立即停止使用并删除本项目。

**使用本项目即表示你已阅读、理解并同意上述声明。** 如果你不同意，请勿使用。

---

## 致谢

感谢 [@richyhu](https://github.com/richyhu) 完成了本项目绝大部分的编写工作。从大创申报书到入党申请书的每一个栏目、每一条避坑指南、每一处格式规范，都来自他对高校申报体系的深入理解和大量实践积累。

这些 skill 凝结了巨大的工作量——不只是写 prompt，而是把一套庞杂的、隐性的、散落在各高校通知文件里的领域知识，系统性地编码为 Agent 可理解、可执行的结构化指令。这在国内 Agent Skills 生态中是非常稀缺的工作。

---

## 贡献

欢迎提 Issue 和 PR。

如果你发现：

- 某个学校的模板格式有特殊要求，Skill 没有覆盖
- 某个申报类型的评审标准有更新（比如互联网+ 2026 新规）
- Skill 输出的内容有事实性错误

请告诉我们。

贡献前请阅读 `references/writing_guide.md` 了解撰写规范，确保你的修改符合项目的诚实底线。

---

## 许可

MIT License —— 随便用、随便改，但请保留原作者署名和免责声明。

详见 [LICENSE](./LICENSE)。

---

*如果你用这个项目写出了一份好的申报书，或者踩了什么坑，欢迎提 Issue 分享。让后面的人少走弯路。*
