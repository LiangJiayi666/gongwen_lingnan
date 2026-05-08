---
name: gongwen-lingnan-2-est
description: 生成符合岭南学院公文格式规范的 Word 文档。智能估计页数，偶数页不加分页符、奇数页加分页符。版记始终保留。
---
# 公文 Word 文档生成器（智能分页）

基于 python-docx / lxml 实现岭南学院公文格式规范。

与 `gongwen-lingnan-2` 的区别：生成版记前先估计文档页数。若估计为偶数页，使用**连续分节符**（不新增空白页），版记排在末页底部；若估计为奇数页，使用**下一页分节符**（新增空白页承载版记）。版记始终保留。

## 工作目录规范

1. **存放中间产物**：所有调试文件、中间产物（如临时 Markdown、JSON、脚本日志等）以及最终输出的 Word 文档，都必须放在该临时目录中。
2. **告知用户路径**：在回复用户时，必须明确告诉用户结果的储存位置，例如“生成的文档已保存在 `gongwen-lingnan-2-temp/20250415-143052/输出.docx`”。

## 操作步骤

**第一步：创建临时目录**

在每次执行本 Skill 时，首先在 `gongwen-lingnan-2-temp/` 下建立子目录，命名格式：

```
gongwen-lingnan-2-temp/<原始文件名（特殊字符化为_）>-<YYYYMMDD-HHMMSS>/
```

- `原始文件名`：用户提供的 .docx 文件名（不含路径），所有非字母数字汉字字符替换为 `_`
- `YYYYMMDD-HHMMSS`：当前时间戳
- 若无法获取原始文件名（如纯命令行参数输入），则使用 `gongwen-<YYYYMMDD-HHMMSS>`

示例：

```bash
# 原始文件: 通知-2026.docx → sanitized: 通知_2026
filename="通知_2026"
timestamp=$(date +%Y%m%d-%H%M%S)
mkdir -p "gongwen-lingnan-2-temp/${filename}-${timestamp}"
```

**第二步：读取现有 docx 文件**

如果用户提供了现有的 `.docx` 文件作为参考或需要读取其内容，**必须先使用 `/markitdown` 命令将其转换为 Markdown** 后再进行后续处理。

**第三步：准备 Markdown 文件**

```markdown
---
doc_number: 岭院函〔2026〕1号
recipients: 各相关单位
signer: 中山大学岭南学院
date: 2026年4月4日
contact: 张三，联系电话：020-84112188
attachments:
  - 附件一名称
---
# 关于开展年度工作总结的通知

正文第一段。

## 一、一级标题

正文内容。

### （一）二级标题

正文内容。
```

**第四步：生成公文**

在第一步创建的临时目录中执行：

```bash
cd "gongwen-lingnan-2-temp/${filename}-${timestamp}"
python .claude/skills/gongwen-lingnan-2-est/scripts/gongwen_doc.py --md input.md -o 输出.docx
```

## 格式规范

- 标题：方正小标宋简体 22pt 居中
- 正文：仿宋_GB2312 16pt，27磅行距，首行缩进2字符
- 一级标题：黑体 16pt
- 公开方式：「公开方式」黑体，「依申请公开」仿宋GB2312，四号字

## 环境要求

Windows + Microsoft Word + Python 3.6+ + pywin32

## 作者
梁佳仪
