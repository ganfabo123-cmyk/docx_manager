import json
import os

# 测试所有JSON文件
json_files = [
    'data/test_headings_body.json',
    'data/test_tables.json',
    'data/test_formulas.json',
    'data/test_images.json',
    'data/test_toc.json'
]

for file_path in json_files:
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            print(f"✓ {file_path}: 读取成功")
            if 'content' in data:
                print(f"  - 内容数量: {len(data['content'])}")
        except Exception as e:
            print(f"✗ {file_path}: 读取失败 - {e}")
    else:
        print(f"✗ {file_path}: 文件不存在")
