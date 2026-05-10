# Full Style Docx Generator - 完整样式Word文档生成器

## 📋 项目概述

**Full Style Docx Generator** 是一个高级的文档处理系统，用于将结构化文本文件自动转换为样式规范的Word文档（DOCX）。该项目特别适用于学术论文、技术报告等需要严格格式控制的文档生成场景。

**核心应用场景**: 哈工大本科综合设计（论文）书写规范示范文档的自动化处理与生成。

---

## ✨ 主要功能

### 1. 文本文件解析 (TXT Parsing)
- **远程文件下载**: 支持从URL直接下载TXT文件进行处理
- **智能文本分块**: 自动识别和分解文本中的各种元素（标题、段落、列表等）
- **Markdown符号移除**: 可选地清理文本中的Markdown格式标记，保留纯内容
- **元素标准化**: 将解析后的文本元素转换为统一的JSON结构

### 2. 文档结构识别
- **短文本块识别**: 自动检测标题、目录项等短文本元素
- **目录生成**: 识别文档中的TOC（Table of Contents）结构
- **层级识别**: 支持多层级标题的自动识别（H1/H2/H3）

### 3. 高级内容处理
- **表格检测与提取**: 自动识别Markdown格式表格、制表符表格、多列结构表格
- **公式转换**: LaTeX公式到Office MathML（OMML）的自动转换
  - 支持行内公式($...$)的自动扫描和转换
  - 使用XSL样式表进行格式转换
- **图像处理**: 图像引用检测、分组和位置分配
  - 自动识别图片引用关键词（"如图"、"见图"等）
  - 智能图像分组和章节关联

### 4. Word文档生成
- **格式化DOCX输出**: 生成包含完整样式的Word文档
- **样式回填**: 将文本分析后的样式信息自动应用到生成的文档中
- **DOCX还原**: 支持从JSON格式恢复原始DOCX文档结构

### 5. RESTful API服务
- **Flask Web服务**: 完整的HTTP API接口
- **异步处理**: 支持大规模文档处理
- **错误处理**: 全局异常处理和详细日志记录

---

## 📁 项目结构

```
full_style_docx_generator/
├── main.py                      # 应用启动入口
├── server.py                    # Flask服务器配置
├── api/
│   ├── routes.py               # API路由定义（核心接口）
│   └── __init__.py
├── core/                        # 核心处理模块
│   ├── document_processor.py    # 文档处理器（短块识别、DOCX生成）
│   ├── file_parser.py           # 文件解析器（统一接口）
│   ├── formula/                 # 公式处理模块
│   │   ├── converter.py         # LaTeX→OMML转换
│   │   ├── detector.py          # 公式检测
│   │   ├── models.py            # 数据模型
│   │   └── assets/              # XSL样式表资源
│   ├── image/                   # 图像处理模块
│   │   ├── generator.py         # 图像生成和位置分配
│   │   ├── detector.py          # 图像引用检测
│   │   ├── models.py            # 数据结构
│   │   └── assets/              # 图像资源
│   ├── table/                   # 表格处理模块
│   │   ├── detector.py          # 表格检测
│   │   ├── extractor.py         # 表格提取
│   │   ├── models.py            # 表格数据结构
│   │   └── docxs/               # 表格模板
│   ├── short_block/             # 短文本块处理
│   │   └── __init__.py
│   └── __init__.py
├── utils/                       # 工具函数库
│   ├── docx_parser.py           # DOCX文档解析器
│   ├── text_extractor.py        # 文本提取工具
│   ├── docx_style_backfill.py   # 样式回填工具
│   ├── docx_restorer.py         # DOCX文档还原
│   └── __init__.py
├── temp/                        # 临时文件存储
└── Readme.md                    # 项目文档
```

---

## 🔌 API 接口

### POST /parse-txt-file
解析TXT文件并提取文本元素。

**请求体 (JSON)**:
```json
{
  "url": "https://example.com/document.txt",
  "remove_markdown": true
}
```

或兼容格式:
```json
{
  "txt": {
    "url": "https://example.com/document.txt"
  },
  "remove_markdown": true
}
```

**请求参数**:
- `url` (string, 必需): TXT文件的远程URL
- `remove_markdown` (boolean, 可选, 默认true): 是否移除Markdown符号

**响应 (200 OK)**:
```json
{
  "status": "success",
  "message": "File parsed and blocks saved",
  "text_count": 156
}
```

**功能流程**:
1. 从URL下载TXT文件
2. 将文件保存到临时目录
3. 解析文件为文本元素列表
4. 保存解析结果到JSON文件
   - `data/parsed_blocks.json` - 完整解析结果
   - `data/full_parsed.json` - 所有提取的元素

---

## 🔧 核心模块详解

### 文件解析器 (file_parser.py)

主要函数:

**`remove_markdown_symbols(text) -> str`**
- 移除Markdown格式标记
- 支持: `#标题`, `**加粗**`, `_斜体_`, `` `代码` ``, `>引用`, 列表, 链接等

**`parse_file(path, remove_md=True) -> Dict`**
- 统一的文件解析接口
- 支持: TXT、DOCX、PDF等格式
- 返回标准化的元素列表

**`parse_text_to_elements(text, remove_md=True) -> List[Dict]`**
- 将纯文本转换为结构化元素
- 输出格式:
  ```json
  {
    "id": "elem_1",
    "type": "body",
    "content": "文本内容",
    "style": {
      "style_name": "Normal",
      "alignment": "left"
    }
  }
  ```

### 文档处理器 (document_processor.py)

**`identify_short_blocks(blocks) -> List[Dict]`**
- 两阶段识别算法:
  1. 第一阶段: 收集内容长度<30的文本块
  2. 第二阶段: 收集有对应TOC项的文本块（标题）
- 优先返回第二阶段结果（包含标题）

**`generate_docx_document(parsed_styles, output_path) -> bool`**
- 从解析的样式信息生成DOCX文档
- 支持多层级标题、目录项、正文段落
- 应用格式化样式（对齐、缩进等）

### 表格处理 (table/detector.py)

支持三种表格格式检测:
1. **Markdown竖线表格** - `| 列1 | 列2 |`
2. **制表符分列** - 使用`\t`分隔列
3. **多列结构** - 多个2+空格分隔的列

**`is_suspected_table(content) -> bool`**
- 检测内容是否为表格

**`detect_table_blocks(elements) -> List[Dict]`**
- 从元素列表中提取所有疑似表格

### 公式处理 (formula/converter.py)

**`latex_to_omath(latex) -> str`**
- 将LaTeX公式转换为Office MathML格式
- 依赖: `latex2mathml`, `lxml`

**`scan_and_convert_dollar_inline(elements) -> List[Dict]`**
- 扫描所有元素中的行内公式 ($...$)
- 自动转换，失败时跳过

### 图像处理 (image/generator.py)

支持智能图像处理流程:
- 图像引用关键词识别（"如图"、"见图"、"下图"等）
- 段落候选项筛选（含引用关键词及上下文）
- 按章节自动分组
- 智能位置分配

---

## 📊 数据格式示例

### 解析后的元素格式

```json
{
  "id": "elem_1",
  "type": "h1",
  "content": "第1章 绪论",
  "style": {
    "style_name": "Heading1",
    "alignment": "center",
    "font_size": 14
  }
}
```

### 完整解析结果

```json
{
  "text_elements": [
    {
      "id": "elem_0",
      "type": "title",
      "content": "哈尔滨工业大学本科综合设计（论文）",
      "style": {...}
    },
    {
      "id": "elem_1",
      "type": "h1",
      "content": "摘要",
      "style": {...}
    },
    {
      "id": "elem_2",
      "type": "body",
      "content": "气体静压轴承由于具有运动精度高、摩擦损耗小...",
      "style": {...}
    }
  ]
}
```

---

## 🚀 快速开始

### 1. 环境配置

**依赖包**:
```
Flask
python-docx
requests
latex2mathml
lxml
```

**安装依赖**:
```bash
pip install -r requirements.txt
```

### 2. 启动服务

```bash
# 方式1: 使用main.py
python main.py

# 方式2: 直接运行Flask
python server.py
```

服务器将在 `http://0.0.0.0:5000` 启动

### 3. 调用API

```bash
curl -X POST http://localhost:5000/parse-txt-file \
  -H "Content-Type: application/json" \
  -d '{
    "url": "https://example.com/document.txt",
    "remove_markdown": true
  }'
```

### 4. 查看结果

解析结果保存到:
- `data/parsed_blocks.json` - 完整元数据
- `data/full_parsed.json` - 所有元素列表

---

## 🔬 工作流程

```
输入TXT文件
    ↓
[下载/读取文件]
    ↓
[可选: 移除Markdown符号]
    ↓
[文本分块 - 逐行处理]
    ↓
[元素识别]
├─→ [短块识别] → 标题、目录项
├─→ [表格检测] → 表格元素
├─→ [公式检测] → LaTeX公式
└─→ [图像检测] → 图片引用
    ↓
[样式分析] - 为每个元素分配样式
    ↓
[JSON结构化] - 输出标准元素列表
    ↓
[DOCX生成] - 根据样式生成Word文档
    ↓
输出DOCX文件
```

---

## 🛠️ 常见使用场景

### 场景1: 论文格式标准化

```python
from core.file_parser import parse_file

# 解析论文TXT版本
result = parse_file("paper.txt", remove_md=True)

# 生成规范格式的DOCX
from core.document_processor import generate_docx_document
generate_docx_document(result['text_elements'], "paper_formatted.docx")
```

### 场景2: 批量文档处理

```bash
# 通过API处理多个文件
for file in *.txt; do
  curl -X POST http://localhost:5000/parse-txt-file \
    -H "Content-Type: application/json" \
    -d "{\"url\": \"file://$file\"}"
done
```

### 场景3: 自定义样式应用

```python
from core.file_parser import parse_file
from utils.docx_style_backfill import backfill_styles

# 解析文件
elements = parse_file("document.txt")['text_elements']

# 应用自定义样式
styled_elements = backfill_styles(elements, style_config)

# 生成文档
from core.document_processor import generate_docx_document
generate_docx_document(styled_elements, "output.docx")
```

---

## 📝 注意事项

### 平台兼容性
- **Windows**: 完全支持，包括Word COM接口
- **Linux/Mac**: 部分功能受限，不支持.doc格式转换

### 依赖要求
- Python 3.7+
- python-docx >= 0.8.11
- 公式转换需要: `latex2mathml`, `lxml`
- 表格处理需要相应的DOCX模板文件

### 性能考虑
- 大文件（>10MB）处理时间较长
- 图像处理依赖LLM调用，需要网络连接
- 建议启用异步处理大批量任务

### 文件格式要求
- TXT文件编码: **UTF-8推荐**
- 表格检测: 支持Markdown、制表符、多列格式
- 公式格式: LaTeX语法($...$)

---

## 🔄 扩展和定制

### 添加新的内容处理器

1. 在`core/`下创建新模块文件夹
2. 实现内容检测器 (detector.py)
3. 实现内容生成器 (generator.py)
4. 在`document_processor.py`中注册处理流程

### 自定义样式方案

编辑`utils/docx_style_backfill.py`中的样式映射规则:

```python
STYLE_MAPPING = {
    'h1': {'font_size': 16, 'bold': True, 'alignment': 'center'},
    'h2': {'font_size': 14, 'bold': True},
    'body': {'font_size': 12, 'alignment': 'justify'}
}
```

### 集成新的数据源

在`api/routes.py`中添加新的路由处理不同的输入源（数据库、API、文件系统等）

---

## 📋 调试和日志

服务器启用详细日志记录，包括:
- 📦 请求入参日志
- 🔍 解析过程日志
- 💾 文件保存日志
- ❌ 错误堆栈跟踪

所有日志输出到标准输出，便于监控和调试。

示例日志:
```
==================================================
🌐 [API IN] 收到请求: POST /parse-txt-file
📦 [PAYLOAD] 原始入参: {"url": "...", "remove_markdown": true}
🔍 [PARSE] 提取到的 URL: https://example.com/doc.txt
⬇️  [ACTION] 开始下载 txt 文件...
✅ [SUCCESS] 文件下载成功，保存至: /tmp/input_12345.txt (大小: 45678 bytes)
📊 [INFO] 解析完成，共提取 156 个文本块
💾 [SAVE] 解析结果已保存至: data/parsed_blocks.json
🏁 [API OUT] 请求处理成功返回 200
```

---

## 📚 相关工具

- **parent**: `docx_manager` - 文档管理系统
- **sibling**: `docx_engine` - 高级文档引擎
- **utility**: `utils/` - 通用工具库

---

## 📄 许可证

本项目属于哈尔滨工业大学论文规范辅助工具集。

---

## 👨‍💻 技术栈

- **后端**: Python Flask
- **文档处理**: python-docx
- **公式处理**: LaTeX → OMML (via latex2mathml)
- **XML处理**: lxml
- **数据存储**: JSON
- **HTTP通信**: requests

---

## 📞 支持与反馈

如遇到问题或有改进建议，请检查:
1. 输入文件格式是否正确 (UTF-8编码)
2. 所有依赖包是否正确安装
3. 服务器日志中的详细错误信息
4. API请求格式是否符合规范
