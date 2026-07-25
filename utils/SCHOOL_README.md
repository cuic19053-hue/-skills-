# 学校模板适配层使用说明（SCHOOL_README）

> 让 22+ 个子 skill 适配不同学校的页眉/印章/字段名/签字栏/字体/页边距差异。

## 一、设计目标

不同高校的学生申报书在以下方面存在差异：

| 差异项 | 示例 | 影响 |
| --- | --- | --- |
| 页眉文字 / 校徽 | "清华大学" + 紫色校徽 vs "北京大学" + 红色校徽 | 视觉规范 |
| 印章位置与大小 | 清华右下角 Ø4.2cm / 北大居中 Ø4.0cm | 打印校对 |
| 字段名 | "学习成绩" vs "GPA" vs "绩点" vs "平均学分绩点" | 表格列头 |
| 签字栏 | 默认 4 方 / 清华加"导师" / 北大改"班主任" | 表单结构 |
| 字体偏好 | 多数用宋体，清华要求仿宋 | 正文字体 |
| 页边距 | 多数 2.5cm，北大左右 2.8cm | 排版边界 |
| 申请理由字数 | 清华 200 字 vs 北大 250 字 | 内容生成 |

本适配层把这些差异从 22+ 个 `build.py` 中抽离，集中到 `schools/template_<name>.json`，让 build.py 调用统一 API 即可生成符合任何学校规范的 docx。

## 二、目录结构

```
utils/
├── school_template.py        # 核心适配层（SchoolTemplate 类 + API）
├── SCHOOL_README.md          # 本文件
└── schools/
    ├── template_default.json # 默认模板（所有字段兜底值）
    ├── template_tsinghua.json
    ├── template_pku.json
    ├── template_zju.json
    └── template_whu.json
```

## 三、快速上手

### 3.1 在 build.py 中调用

```python
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "utils"))

from docx import Document
from school_template import load_template, apply_template, build_signature_table

# 1. 加载学校模板（school_name 可由用户参数传入）
tpl = load_template("tsinghua")  # 或 "清华" / "清华大学" / "THU"

# 2. 创建 Document 并应用模板（页边距/页眉/页脚/印章）
doc = Document()
apply_template(doc, tpl)

# 3. 用学校特定字段名写表格列头
gpa_label = tpl.get_field("gpa")        # -> "GPA"（清华） / "绩点"（北大）
rank_label = tpl.get_field("rank")      # -> "本专业排名"
# 写入表格时用 gpa_label 而非写死 "学习成绩"

# 4. 用学校特定字体写正文
body_font = tpl.body_font()             # -> "仿宋"（清华） / "宋体"（北大）
heading_font = tpl.heading_font()       # -> "黑体"

# 5. 在落款处生成学校特定签字栏
build_signature_table(doc, tpl)

# 6. 校验申请理由字数
check = tpl.validate_apply_reason(reason_text)
if not check["ok"]:
    print(f"申请理由{check['message']}")
```

### 3.2 命令行自检

```bash
$ python3 utils/school_template.py
============================================================
SchoolTemplate 自检
============================================================
支持学校数量: 5
  - [file    ] default      -> 默认
  - [file    ] pku          -> 北京大学
  - [file    ] tsinghua     -> 清华大学
  - [file    ] whu          -> 武汉大学
  - [file    ] zju          -> 浙江大学

字段映射对照表（gpa / rank / major / class）：
  学校        gpa        rank       major      class
  清华大学    GPA        本专业排名  专业方向   行政班
  北京大学    绩点       年级排名    专业       班级
  浙江大学    平均学分绩点 专业排名  主修专业   所在班级
  武汉大学    学习成绩   专业排名    专业       班级

字体偏好对照表：
  默认        body=宋体   heading=黑体 title=22.0pt
  ...
```

## 四、API 一览

### 4.1 `load_template(school_name="default") -> SchoolTemplate`

加载学校配置。`school_name` 支持以下任一形式（统一规范化为配置文件名）：

| 输入 | 规范化为 |
| --- | --- |
| `"清华"` / `"清华大学"` / `"THU"` / `"tsinghua"` | `tsinghua` |
| `"北大"` / `"北京大学"` / `"PKU"` / `"pku"` | `pku` |
| `"浙大"` / `"浙江大学"` / `"ZJU"` | `zju` |
| `"武大"` / `"武汉大学"` / `"WHU"` | `whu` |

未识别的输入会原样小写化作为文件名查找。

### 4.2 `list_supported_schools() -> List[dict]`

返回所有可用学校配置：

```python
[
    {"id": "default",  "school_name": "默认",       "source": "file"},
    {"id": "tsinghua", "school_name": "清华大学",  "source": "file"},
    {"id": "pku",      "school_name": "北京大学",  "source": "file"},
    {"id": "zju",      "school_name": "浙江大学",  "source": "file"},
    {"id": "whu",      "school_name": "武汉大学",  "source": "file"},
]
```

### 4.3 `apply_template(doc, template, *, apply_seal=True) -> None`

把模板应用到 `docx.Document`，会执行：

1. 设置所有 section 的页边距（`margins`）
2. 设置所有 section 的页眉（`header.text` + `header.logo`，logo 路径相对于 `schools/`）
3. 设置所有 section 的页脚（`footer.text` + 页码域）
4. （可选）在文档末尾添加印章占位段落

`template` 可以是 `SchoolTemplate` 实例、学校名字符串、或配置字典。

### 4.4 `get_field_mapping(school_name, field) -> str`

便捷函数，等价于 `load_template(school_name).get_field(field)`。

### 4.5 `register_template(school_name, config_dict) -> SchoolTemplate`

运行时注册新学校，无需写 JSON 文件：

```python
from school_template import register_template, load_template

register_template("fudan", {
    "school_name": "复旦大学",
    "fonts": {"body": "仿宋"},
    "margins": {"top_cm": 2.8, "bottom_cm": 2.8,
                "left_cm": 2.5, "right_cm": 2.5},
    "apply_reason_chars": {"min": 200, "max": 220},
})

tpl = load_template("fudan")
```

### 4.6 `build_signature_table(doc, template, *, date_str=...) -> None`

按学校模板的 `signature_blocks` 生成 N 行 2 列签字栏表格。

### 4.7 `SchoolTemplate` 实例方法

| 方法 | 返回 | 说明 |
| --- | --- | --- |
| `get_field(field, default=None)` | `str` | 取字段名映射 |
| `remap_record(dict)` | `dict` | 把字典 key 按字段映射重命名 |
| `body_font()` | `str` | 正文字体（如 "宋体" / "仿宋"） |
| `heading_font()` | `str` | 标题字体 |
| `title_size_pt()` | `float` | 标题字号（磅） |
| `apply_reason_min()` | `int` | 申请理由字数下限 |
| `apply_reason_max()` | `int` | 申请理由字数上限 |
| `validate_apply_reason(text)` | `dict` | 字数校验，返回 `{ok, length, min, max, message}` |
| `to_dict()` / `to_json()` | `dict` / `str` | 序列化 |

## 五、JSON 配置字段规范

`template_<name>.json` 必须包含以下字段（缺失的字段会从 `template_default.json` 自动合并）：

```jsonc
{
  "school_name": "清华大学",        // 必填
  "header": {                       // 页眉
    "enabled": true,
    "text": "清华大学",              // 为空时用 school_name
    "logo": "logos/tsinghua_logo.png",  // 相对 schools/ 目录
    "align": "center"               // left/center/right
  },
  "footer": {                       // 页脚
    "enabled": true,
    "text": "清华大学学生事务处",
    "page_number": true             // 是否显示页码
  },
  "seal": {                         // 印章占位
    "enabled": true,
    "position": "right_bottom",     // right_bottom/left_bottom/center_bottom
    "size_cm": 4.2,
    "transparency": 30              // 仅文档说明用，不实际渲染
  },
  "field_mapping": {                // 字段名映射
    "gpa": "GPA",
    "rank": "本专业排名",
    "major": "专业方向",
    "class": "行政班",
    "name": "姓名",
    "student_id": "学号",
    "college": "所在院系",
    "grade": "年级",
    "phone": "联系电话",
    "date": "日期"
  },
  "fonts": {                        // 字体偏好
    "body": "仿宋",                 // 正文（多数学校用宋体）
    "heading": "黑体",
    "title_size": "22"              // 标题字号（磅值字符串或数字）
  },
  "margins": {                      // 页边距（cm）
    "top_cm": 2.8,
    "bottom_cm": 2.5,
    "left_cm": 2.5,
    "right_cm": 2.5
  },
  "signature_blocks": [             // 签字栏顺序与角色
    "申请人",
    "导师",                         // 清华特有
    "辅导员",
    "院系负责人",
    "学校负责人"
  ],
  "apply_reason_chars": {           // 申请理由字数限制
    "min": 180,
    "max": 200
  }
}
```

## 六、新增一所学校的步骤

1. 复制 `template_default.json` 为 `template_<school_id>.json`
   （`<school_id>` 用拼音或英文缩写，如 `fudan` / `sjtu` / `nju`）
2. 按学校实际规范修改字段
3. （可选）把校徽图片放到 `schools/logos/<school_id>_logo.png`，并在 `header.logo` 引用
4. （可选）在 `school_template.py` 的 `_ALIAS_NORMALIZE_RULES` 列表追加中文别名规则
5. 运行 `python3 utils/school_template.py` 自检，确认新学校出现在列表中

## 七、与现有 build.py 的集成路径

本适配层采用**零侵入**设计，现有 22+ 个 build.py 无需立即改动：

- **短期**：仅在新文档场景调用本模块（如调用方传入 `school` 参数时启用页眉/印章/字段名映射）
- **中期**：在 build.py 中把硬编码的 "学习成绩" / "排名" / "申请人签字" 等替换为 `tpl.get_field(...)` 调用
- **长期**：build.py 顶层加 `--school <name>` 参数，未传时默认 `default`

示例 patch（最小侵入）：

```python
# build.py 顶部
from school_template import load_template
TPL = load_template(getattr(args, "school", "default") or "default")

# build.py 内部，把所有写死的字段名替换为：
# "学习成绩" -> TPL.get_field("gpa")
# "排名"     -> TPL.get_field("rank")
# "申请人签字：" -> TPL.signature_blocks[0] + "签字："
```

## 八、4 所学校示例配置对比

| 学校 | gpa 字段 | body 字体 | 签字栏特色 | 页边距（左右） | 申请理由字数 |
| --- | --- | --- | --- | --- | --- |
| 默认 | 学习成绩 | 宋体 | 4 方（学生/辅导员/院系/学校） | 2.5cm | 180-200 |
| 清华 | GPA | 仿宋 | 5 方（加"导师"） | 2.5cm | 180-200 |
| 北大 | 绩点 | 宋体 | 4 方（"辅导员"改"班主任"） | 2.8cm | 230-250 |
| 浙大 | 平均学分绩点 | 宋体 | 4 方（"院系"改"学院"） | 2.5cm | 200-220 |
| 武大 | 学习成绩 | 宋体 | 4 方（"院系"改"学院"） | 3.0cm / 2.5cm | 200-220 |

## 九、常见问题

**Q1：印章是否会被真实画到 docx 里？**
A1：不会。当前实现仅在文档末尾添加一行占位文字"〔此处加盖学校公章 Ø4.2cm〕"用于打印校对。真实红章由学校人工加盖。如需程序化绘制红色圆形印章，可在 `_apply_seal_placeholder` 中扩展，用 `docx.oxml` 绘制椭圆 shape。

**Q2：校徽图片不存在会怎样？**
A2：`_set_page_header` 会自动跳过图片，仅显示文字页眉，不会报错。

**Q3：load_template 失败时的兜底策略？**
A3：若 `template_<name>.json` 不存在，会回退到 `template_default.json` 并把 `school_name` 替换为用户传入的名字。若 `template_default.json` 也不存在，抛 `FileNotFoundError`。

**Q4：能否在不修改 JSON 的情况下覆盖某个字段？**
A4：可以，用 `register_template(name, {...})` 注入运行时配置，它会与 default 模板深合并。

**Q5：字段映射如何处理"申请理由"等多字段场景？**
A5：`field_mapping` 是开放字典，可任意扩展 key，调用方按需取用。`SchoolTemplate.remap_record(record_dict)` 可一次性翻译整个字典。

---

*最后更新：T37 工程化任务 / 学校模板适配层*
