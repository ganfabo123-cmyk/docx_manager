# Full Style DocX Fixer

一个用于从LLM平台接收POST请求并组装用户数据的HTTP服务器。

## 功能特性

- 通过HTTP POST接收来自LLM平台的各类数据
- 支持所有文档元素类型：章节、目录、标题、正文、表格、公式、图片等
- 实时数据收集和组装
- 支持数据导出为JSON格式
- 提供健康检查和重置功能

## 项目结构

```
full_style_docx_fixer/
├── __init__.py
├── main.py                 # 主入口文件
├── config.py               # 配置文件
├── requirements.txt        # 依赖包
├── models/
│   ├── __init__.py
│   └── models.py          # 数据模型定义
└── utils/
    ├── generate_user_data.py
    ├── parse_full_docx.py
    └── user_config.json
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 启动服务器

```bash
python -m full_style_docx_fixer.main
```

### 自定义参数

```bash
python -m full_style_docx_fixer.main --host 127.0.0.1 --port 8080 --debug
```

### 保存数据到文件

```bash
python -m full_style_docx_fixer.main --output data/collected_data.json
```

## API端点

### 基础端点

| 端点 | 方法 | 描述 |
|------|------|------|
| `/health` | GET | 健康检查 |
| `/get_data` | GET | 获取完整用户数据 |
| `/reset` | POST | 重置所有数据 |

### 数据接收端点

| 端点 | 方法 | 描述 |
|------|------|------|
| `/_doc` | POST | 接收文档描述 |
| `/page_footer_config` | POST | 接收页脚配置 |
| `/toc_mode` | POST | 接收目录模式 |
| `/toc_entries` | POST | 接收目录条目 |
| `/content_section` | POST | 接收章节内容 |
| `/content_toc` | POST | 接收目录内容 |
| `/content_heading1` | POST | 接收一级标题 |
| `/content_heading2` | POST | 接收二级标题 |
| `/content_heading3` | POST | 接收三级标题 |
| `/content_body` | POST | 接收正文 |
| `/content_table` | POST | 接收表格 |
| `/content_formula` | POST | 接收公式 |
| `/content_image` | POST | 接收图片 |
| `/references` | POST | 接收参考文献 |
| `/citations` | POST | 接收引用 |

## 请求格式

所有POST请求都应使用JSON格式，数据应包含在`value`字段中：

```json
{
  "value": {
    "字段名": "字段值"
  }
}
```

### 示例请求

### 1. 接收文档描述

```bash
curl -X POST http://localhost:5000/_doc \
  -H "Content-Type: application/json" \
  -d '{"value": "测试文档：从LLM平台收集的用户数据"}'
```

### 2. 接收页脚配置

```bash
curl -X POST http://localhost:5000/page_footer_config \
  -H "Content-Type: application/json" \
  -d '{
    "value": [
      {
        "section": "frontmatter",
        "style": "roman_lower_center",
        "start": 1
      },
      {
        "section": "mainmatter",
        "style": "arabic_dash",
        "start": 1
      }
    ]
  }'
```

### 3. 接收目录模式

```bash
curl -X POST http://localhost:5000/toc_mode \
  -H "Content-Type: application/json" \
  -d '{"value": "manual"}'
```

### 4. 接收目录条目

```bash
curl -X POST http://localhost:5000/toc_entries \
  -H "Content-Type: application/json" \
  -d '{
    "value": [
      {
        "title": "摘  要",
        "level": 1,
        "page": "i"
      },
      {
        "title": "第一章  引言",
        "level": 1,
        "page": 1
      },
      {
        "title": "1.1  研究背景",
        "level": 2,
        "page": 1
      }
    ]
  }'
```

### 5. 接收章节内容

```bash
curl -X POST http://localhost:5000/content_section \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "section_type": "abstract",
      "toc_exclude": true,
      "value": "本测试文档验证了LLM数据收集器的功能。\n\n涵盖内容：标题、正文、表格、公式、图片等。"
    }
  }'
```

### 6. 接收目录内容

```bash
curl -X POST http://localhost:5000/content_toc \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "title": "目  录",
      "toc_title_exclude": true
    }
  }'
```

### 7. 接收一级标题

```bash
curl -X POST http://localhost:5000/content_heading1 \
  -H "Content-Type: application/json" \
  -d '{"value": {"value": "第一章  引言"}}'
```

### 8. 接收二级标题

```bash
curl -X POST http://localhost:5000/content_heading2 \
  -H "Content-Type: application/json" \
  -d '{"value": {"value": "1.1  研究背景"}}'
```

### 9. 接收三级标题

```bash
curl -X POST http://localhost:5000/content_heading3 \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "value": "1.1.1  研究现状",
      "toc_exclude": false
    }
  }'
```

### 10. 接收正文（用户最常用的功能）

```bash
curl -X POST http://localhost:5000/content_body \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "value": "这是正文内容的第一段落。\n\n这是正文内容的第二段落，可以包含多行文本。\n\n这是正文内容的第三段落。"
    }
  }'
```

### 11. 接收表格

```bash
curl -X POST http://localhost:5000/content_table \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "caption": "表 1  测试表格",
      "data": [
        ["参数名称", "取值", "说明"],
        ["学习率", "0.001", "Adam 初始学习率"],
        ["批次大小", "32", "每次迭代样本数"]
      ]
    }
  }'
```

### 12. 接收公式

```bash
curl -X POST http://localhost:5000/content_formula \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "label": "式(1)",
      "latex": "E=mc^2"
    }
  }'
```

### 13. 接收图片

```bash
curl -X POST http://localhost:5000/content_image \
  -H "Content-Type: application/json" \
  -d '{
    "value": {
      "base64": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      "ext": "png",
      "caption": "图 1  测试图片",
      "width": 3.5,
      "align": "center"
    }
  }'
```

### 14. 接收参考文献

```bash
curl -X POST http://localhost:5000/references \
  -H "Content-Type: application/json" \
  -d '{
    "value": [
      {
        "id": 1,
        "text": "Knuth D E. The TeXbook[M]. Addison-Wesley, 1984."
      },
      {
        "id": 2,
        "text": "Clark A. python-docx documentation[EB/OL]. 2023."
      }
    ]
  }'
```

### 15. 接收引用

```bash
curl -X POST http://localhost:5000/citations \
  -H "Content-Type: application/json" \
  -d '{
    "value": [
      {
        "ref_id": 1,
        "before": "LaTeX 是国外学术界广泛使用的排版系统",
        "after": ""
      }
    ]
  }'
```

### 16. 获取完整数据

```bash
curl http://localhost:5000/get_data
```

### 17. 重置数据

```bash
curl -X POST http://localhost:5000/reset
```

### 18. 健康检查

```bash
curl http://localhost:5000/health
```

## 输出数据格式

完整数据遵循`full_user_data_v6.json`的格式：

```json
{
  "_doc": "文档描述",
  "page_footer_config": [...],
  "toc_mode": "manual",
  "toc_entries": [...],
  "content": [...],
  "references": [...],
  "citations": [...]
}
```

## 配置说明

编辑`config.py`文件可以修改以下配置：

- `SERVER_HOST`: 服务器监听地址
- `SERVER_PORT`: 服务器监听端口
- `DEBUG_MODE`: 调试模式
- `AUTO_SAVE`: 自动保存开关
- `AUTO_SAVE_PATH`: 自动保存路径
- `LOG_LEVEL`: 日志级别
- `MAX_CONTENT_LENGTH`: 最大内容长度

## 与LLM平台集成

LLM平台需要按照以下规则发送POST请求：

1. 所有数据应包含在`value`字段中
2. 对于列表数据（如`toc_entries`），`value`应为数组
3. 对于内容数据（如`content_heading1`），`value`应为对象
4. 服务器会返回`{"status": "success"}`或错误信息

## 错误处理

所有端点都有错误处理机制，错误时会返回：

```json
{
  "status": "error",
  "message": "错误描述"
}
```

## 注意事项

1. 服务器默认监听所有网络接口（0.0.0.0），生产环境请配置防火墙
2. 数据存储在内存中，服务器重启后数据会丢失
3. 建议使用`--output`参数定期保存数据
4. 使用`/reset`端点可以清空所有数据重新开始收集
