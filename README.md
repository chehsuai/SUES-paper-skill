# Academic Paper Skill

> 从 Word 模板自动生成符合国内学术规范的论文，引用真实文献，带交叉引用功能。

## 这是什么

一个给 **Claude Code** 用的 skill。给它一个论文模板（.docx），它会：
1. 读取模板中的选题、格式、字数等要求
2. 在知网/Google Scholar 上搜索**真实可查**的参考文献
3. 撰写 5000-8000 字的学术论文
4. 生成带**上标交叉引用**的 Word 文档（Ctrl+点击引用可跳转到文献）

## 安装

### 方法一：作为 Claude Code Skill

将 `skill.md` 放到你的 Claude Code skills 目录：

```bash
# Windows
copy skill.md %USERPROFILE%\.claude\skills\academic-paper.md

# macOS / Linux
cp skill.md ~/.claude/skills/academic-paper.md
```

然后在 Claude Code 中输入 `/academic-paper` 即可调用。

### 方法二：直接使用生成脚本

如果你只需要 Word 文档生成功能（不需要文献搜索和论文撰写环节），可以直接使用 `generate_paper.py`：

```bash
pip install python-docx lxml
python generate_paper.py
```

## 文件说明

| 文件 | 作用 |
|------|------|
| `skill.md` | Skill 指令文件，告诉 Claude Code 如何完成整个论文生成任务 |
| `generate_paper.py` | Word 文档生成脚本，处理排版和交叉引用 |
| `README.md` | 本文件 |

## 生成文档格式

- **正文**：宋体小四（12pt），1.5 倍行距，首行缩进 2 字符
- **标题**：顶格，宋体小四，不加粗，分三级
  - 一级：一、二、三…
  - 二级：（一）（二）（三）…
  - 三级：1. 2. 3.…
- **引用**：上标显示，带书签超链接，Ctrl+点击跳转到文末参考文献
- **参考文献**：宋体五号（10.5pt）
- **页面**：上下 2.54cm，左右 3.18cm

## Python API 快速参考

```python
from generate_paper import PaperGenerator

gen = PaperGenerator()

# 封面
gen.set_cover_info(course='课程名', school='学院', student_id='学号',
                   author='姓名', advisor='导师')

# 标题 & 摘要
gen.set_title('主标题', subtitle='副标题')
gen.set_inner_header('完整标题', '学号 姓名')
gen.set_abstract('摘要正文...', keywords='关键词1；关键词2')

# 正文
gen.add_heading('一、绪论')           # 一级标题
gen.add_heading('（一）研究背景', level=2)  # 二级标题
gen.add_body('这里有引用[1][2]...')   # 自动处理引用

# 参考文献
gen.add_reference_section_title()
gen.add_reference('作者. 题名[J]. 刊名, 年份.', ref_num=1)

gen.save('output.docx')
```

## 交叉引用原理

正文中的 `[1]` → Word 超链接 `<w:hyperlink anchor="ref1">` → 参考文献书签 `<w:bookmarkStart name="ref1">`

```
正文 [1]  ──── Ctrl+点击 ────▶  参考文献 [1] 作者. 题名...
  (上标)                           (书签锚点)
```

## 依赖

- Python 3.8+
- `python-docx` — Word 文档生成
- `lxml` — XML 操作（书签和超链接）

```bash
pip install python-docx lxml
```

## 注意事项

- **参考文献必须真实**：所有文献都通过 WebSearch 在知网/Google Scholar 上确认可查后再写入
- **不虚构信息**：学号、姓名等个人信息从模板占位符保留，由用户自行替换
- **模板括号说明要删除**：生成完成后，模板中原有的括号内格式注解一律删除

## License

MIT
