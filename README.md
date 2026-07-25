# 大学生申报书/申请书制作 Skill 集合

[![Agent Skills](https://img.shields.io/badge/Agent_Skills-Compatible-6e3bf0)](https://agentskills.io/specification)
[![Platform](https://img.shields.io/badge/Platform-Claude_Code_|_Codex_|_Cursor_|_Copilot_|_Gemini_CLI-blue)]()
[![License](https://img.shields.io/badge/License-MIT-green)](./LICENSE)
[![Install](https://img.shields.io/badge/install-npx_skills_add_cuic19053--hue/-skills---black)](https://github.com/evaisse/npx-skills)

一套覆盖中国大学生常见申报场景的 Agent Skills。18 个子 skill，从大创立项到入党申请书，从奖学金到保研推免，让 AI 编程助手变成你身边那个"帮学弟学妹改过几十份申报书"的学长。

---

## 为什么做这个

中国大学生每年要写多少份申报书？大创、奖学金、三下乡、入党、保研——每一份都要求格式规范、语气恰当、事实准确。但大多数人要么对着空白模板发愁，要么去网上搜一些不知所云的范文。

ChatGPT 可以帮你写，但通用模型不知道"大创"是什么、不知道国家奖学金评审看什么、不知道入党申请书有查重。

这个项目把**高校申报书的领域知识**打包成了 Agent Skills——不是给聊天框里贴一段 prompt，而是让你的 AI 编程助手（Claude Code、Codex、Cursor 等）在需要时自动加载对应的专业知识。

## 能写哪些申报书

### 项目类（要钱、要立项）

- **大创创新训练项目** —— 学术研究型，文献综述 + 技术路线 + 预期成果
- **大创创业训练项目** —— 商业计划书型，市场分析 + 商业模式
- **大创创业实践项目** —— 实际运营型，团队 + 财务 + 运营数据
- **校级/院级科研立项** —— 比大创轻量，侧重文献和方法
- **挑战杯作品申报书** —— 课外学术科技作品
- **互联网+商业计划书** —— 含 2025 新评审标准

### 评优类（要荣誉）

- **国家奖学金**（8000 元） —— 学习 + 综合素质
- **国家励志奖学金**（5000 元） —— 学习 + 家庭经济
- **校级奖学金** —— 纯学业导向
- **企业专项奖学金** —— 学习 + 行业认知
- **单项奖学金** —— 科研/社工/文体单项突出
- **优秀毕业生** —— 省级/校级，四年综合
- **优秀学生/三好学生** —— 学年综合
- **文明大学生/优秀团员** —— 道德文明侧重

### 活动类

- **三下乡社会实践** —— 含安全保障预案、每日行程、宣传计划

### 政治身份类

- **入党申请书** —— 标准 4000 字，含党章引用规范
- **思想汇报** —— 季度汇报模板

### 升学类

- **保研推免申请** —— 个人陈述 + 科研经历 + 推荐信

## 项目结构

```
college-application-doc/
├── SKILL.md                 # 路由入口：识别类型 -> 分流到子 skill
├── README.md                # 你正在看的
├── LICENSE                  # MIT
├── subskills/               # 18 个子 skill，每个独立文件夹
│   ├── innovation_research/SKILL.md      # 大创-创新训练
│   ├── entrepreneurship_training/SKILL.md # 大创-创业训练
│   ├── entrepreneurship_practice/SKILL.md # 大创-创业实践
│   ├── college_research/SKILL.md         # 科研立项
│   ├── challenge_cup/SKILL.md            # 挑战杯
│   ├── internet_plus/SKILL.md            # 互联网+
│   ├── national_scholarship/SKILL.md     # 国家奖学金
│   ├── motivation_scholarship/SKILL.md   # 国家励志奖学金
│   ├── university_scholarship/SKILL.md   # 校级奖学金
│   ├── enterprise_scholarship/SKILL.md   # 企业奖学金
│   ├── single_scholarship/SKILL.md       # 单项奖学金
│   ├── outstanding_graduate/SKILL.md     # 优秀毕业生
│   ├── outstanding_student/SKILL.md      # 优秀学生
│   ├── civilized_student/SKILL.md        # 文明大学生
│   ├── social_survey/SKILL.md            # 三下乡
│   ├── party_application/SKILL.md        # 入党申请书
│   ├── thought_report/SKILL.md           # 思想汇报
│   └── graduate_recommendation/SKILL.md  # 保研推免
├── references/              # 共享知识库（所有子 skill 按需引用）
│   ├── types.md             # 8 类申报书栏目结构详解
│   ├── writing_guide.md     # 撰写规范 + 格式标准 + 信息采集清单
│   ├── review_criteria.md   # 评审打分维度
│   └── pitfalls.md          # 常见错误与避坑指南
├── scripts/
│   └── build_docx.py        # 生成 Word 文档的辅助脚本
└── examples/
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

不把 18 个 skill 全塞进上下文。

### 3. 学校差异兼容

每个学校的模板有微调（栏目顺序、字号、页边距）。Skill 的策略是：

- 优先用用户提供的学校模板
- Skill 提供**内容质量参考**（怎么写好、怎么避坑）
- 格式标准以 `references/writing_guide.md` 为准

---

## 安装

本项目遵循 [Agent Skills 开放标准](https://agentskills.io/specification)，支持 Claude Code、Codex、Cursor、GitHub Copilot、Gemini CLI、Windsurf 等所有兼容 SKILL.md 的 Agent。

### 推荐方式：npx skills（一键安装）

[`npx skills`](https://github.com/evaisse/npx-skills) 是 Agent Skills 生态的官方 CLI 工具（Vercel Labs 出品），一行命令搞定安装、更新、卸载。

```bash
# 安装全部 18 个 skill（项目级，提交到 Git 可与团队共享）
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
