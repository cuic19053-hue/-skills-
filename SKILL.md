---
name: college-application-doc
description: >-
  大学生申报书/申请书制作路由 skill。当用户提到"大学生申报书""申请书""立项书""申报材料"等模糊词时触发，负责识别具体类型并分流到 18 个细分子 skill。子 skill 覆盖：大创（创新训练/创业训练/创业实践）、科研立项、奖学金（国奖/励志/校级/企业/单项）、三下乡、评优（优秀毕业生/优秀学生/文明大学生）、入党（申请书/思想汇报）、保研推免、学科竞赛（挑战杯/互联网+）。每个子 skill 都有万字级 SKILL.md，详细到每个栏目、每段话、每个字。即使用户只说"帮我写个申报书"也应当触发并先确认具体类型。
---

# 大学生申报书制作（路由）

本 skill 是**路由入口**，不直接生成申报书。它的职责是：识别用户要写哪一类申报书，然后分流到对应的子 skill。

## 何时触发

- 用户说"帮我写个申报书 / 申请书 / 立项书"——必触发，先做类型识别
- 用户已经说了具体类型（如"国家奖学金申请书"）——直接路由到对应子 skill
- 用户给了一份空白模板或旧版申报书要改——识别类型后路由

## 18 个子 skill 索引

### 项目类（要钱、要立项）

| 子 skill | 中文名 | 触发关键词 |
|----------|--------|------------|
| `dachuang-innovation-training` | 大创创新训练项目 | "大创""创新训练""学术研究项目""国家级大创" |
| `dachuang-entrepreneurship-training` | 大创创业训练项目 | "创业训练""商业计划书项目""大创创业类" |
| `dachuang-entrepreneurship-practice` | 大创创业实践项目 | "创业实践""实际注册公司""大创实践类" |
| `college-research-project` | 校级/院级科研立项 | "科研立项""课题申报""SRTP""学生科研" |
| `competition-challenge-cup` | 挑战杯作品申报书 | "挑战杯""课外学术科技作品" |
| `competition-internet-plus` | 互联网+商业计划书 | "互联网+""创新创业大赛""创业大赛" |

### 评优类（要荣誉）

| 子 skill | 中文名 | 触发关键词 |
|----------|--------|------------|
| `scholarship-national` | 国家奖学金 | "国家奖学金""国奖""8000元奖学金" |
| `scholarship-inspirational` | 国家励志奖学金 | "国家励志""励志奖学金""5000元" |
| `scholarship-university` | 校级奖学金 | "校级奖学金""校奖""一等奖学金" |
| `scholarship-enterprise` | 企业专项奖学金 | "企业奖学金""专项奖学金""XX公司奖学金" |
| `scholarship-single` | 单项奖学金 | "科研单项""社会工作单项""文体单项" |
| `honor-outstanding-graduate` | 优秀毕业生 | "优秀毕业生""省优毕业生""校优毕业生" |
| `honor-outstanding-student` | 优秀学生/三好学生 | "三好学生""优秀学生""优秀学生标兵" |
| `honor-civilized-student` | 文明大学生/优秀团员 | "文明大学生""优秀团员""优秀团干" |

### 活动类

| 子 skill | 中文名 | 触发关键词 |
|----------|--------|------------|
| `social-practice-three-rural` | 三下乡社会实践 | "三下乡""暑期社会实践""返家乡" |

### 政治身份类

| 子 skill | 中文名 | 触发关键词 |
|----------|--------|------------|
| `party-membership-application` | 入党申请书 | "入党申请书""申请入党""志愿加入中国共产党" |
| `party-thought-report` | 思想汇报 | "思想汇报""入党积极分子汇报""季度汇报" |

### 升学类

| 子 skill | 中文名 | 触发关键词 |
|----------|--------|------------|
| `postgrad-recommendation` | 保研推免申请 | "保研""推免""保研申请""推免表" |

## 类型识别决策树

```
用户要"申报 / 申请"什么？
│
├─ 要钱、要立项？
│   ├─ 大创训练计划
│   │   ├─ 偏学术研究 → dachuang-innovation-training
│   │   ├─ 偏商业计划（不实际运营） → dachuang-entrepreneurship-training
│   │   └─ 实际注册公司运营 → dachuang-entrepreneurship-practice
│   ├─ 校级/院级科研课题 → college-research-project
│   └─ 创业大赛
│       ├─ 学术科技作品 → competition-challenge-cup
│       └─ 创业计划书 → competition-internet-plus
│
├─ 要荣誉？
│   ├─ 奖学金
│   │   ├─ 国家级 8000 元，纯学业+综合 → scholarship-national
│   │   ├─ 国家级 5000 元，家庭经济困难 → scholarship-inspirational
│   │   ├─ 校级（一二三等） → scholarship-university
│   │   ├─ 企业/社会捐赠 → scholarship-enterprise
│   │   └─ 单项（科研/社工/文体） → scholarship-single
│   ├─ 优秀毕业生（省/校级） → honor-outstanding-graduate
│   ├─ 三好学生/优秀学生 → honor-outstanding-student
│   └─ 文明大学生/优秀团员 → honor-civilized-student
│
├─ 要做活动？
│   └─ 三下乡/寒暑假社会实践 → social-practice-three-rural
│
├─ 政治身份？
│   ├─ 第一次递交入党申请 → party-membership-application
│   └─ 已是积极分子，季度汇报 → party-thought-report
│
└─ 要升学？
    └─ 保研推免 → postgrad-recommendation
```

## 路由流程

1. **识别**：根据用户原话匹配上表关键词
2. **确认**：如果不确定，列出 2~3 个候选让用户选
3. **移交**：明确告诉用户"正在为你调用 [子 skill 名称]，它专门处理 [该类型]"
4. **执行**：调用子 skill 的 SKILL.md，按其工作流执行

## 通用诚实底线（所有子 skill 共享）

大学生申报书是面向学校/教育主管部门的正式材料，**造假会被记入档案甚至触发学籍处分**。所有子 skill 必须遵守：

1. **只写用户能提供证据的事实**——问"你拿过什么奖"，不替用户列奖项
2. **不放大、不润色**——"班级第二"不能写成"成绩优异名列前茅"
3. **不留模糊占位**——拿不准的信息要追问，不要写"获得多项荣誉"这种空话
4. **提醒用户复核**——生成 docx 后必须明确提示用户："以下内容基于你提供的信息生成，提交前请逐项核对真实性"

## 通用格式标准（所有子 skill 共享）

- 纸张：A4，页边距 2.5cm
- 正文：宋体小四，1.5 倍行距，首行缩进 2 字符
- 一级标题：黑体三号，居中，段前段后 12pt
- 二级标题：黑体小三，左对齐，段前段后 6pt
- 三级标题：宋体四号加粗
- 表格：宋体五号，居中
- 英文/数字：Times New Roman

具体到每个类型的栏目顺序、字数要求、撰写要点，由对应子 skill 的 SKILL.md 详细规定。

## 与其他 skill 的协作

- **docx skill**：实际生成 Word 文档时调用
- **pdf skill**：用户要求 PDF 输出时调用
- **charts skill**：申报书里要画技术路线图 / 进度甘特图 / 经费饼图时调用
- **web-search skill**：用户要求"参考同类项目"时调用

## 反模式

- ❌ 不识别类型就直接写——会拿大创模板写奖学金申请书
- ❌ 路由后还在本 skill 里写内容——应该完全移交给子 skill
- ❌ 替用户编造任何可核查事实——红线
