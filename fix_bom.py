import json

# 修复test_tables.json的BOM问题
file_path = 'data/test_tables.json'

try:
    # 以utf-8-sig编码读取文件（自动处理BOM）
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
    
    # 以utf-8编码写回文件（不带BOM）
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ 修复了 {file_path} 的BOM问题")
except Exception as e:
    print(f"✗ 修复BOM问题失败: {e}")
