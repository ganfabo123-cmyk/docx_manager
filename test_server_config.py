import requests
import json

# 测试服务器地址
BASE_URL = 'http://localhost:5000'

def test_toc_title():
    """测试 toc_title 配置"""
    url = f'{BASE_URL}/toc_title'
    data = {'value': '目录'}
    response = requests.post(url, json=data)
    print(f"\ntest_toc_title:")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}")

def test_image_defaults():
    """测试 image_defaults 配置"""
    url = f'{BASE_URL}/image_defaults'
    data = {
        'value': {
            'width': 4.0,
            'align': 'left',
            'ext': 'jpg'
        }
    }
    response = requests.post(url, json=data)
    print(f"\ntest_image_defaults:")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}")

def test_formula_defaults():
    """测试 formula_defaults 配置"""
    url = f'{BASE_URL}/formula_defaults'
    data = {
        'value': {
            'label_prefix': '公式'
        }
    }
    response = requests.post(url, json=data)
    print(f"\ntest_formula_defaults:")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}")

def test_page_footer_config_valid():
    """测试有效的 page_footer_config"""
    url = f'{BASE_URL}/page_footer_config'
    data = {
        'value': [
            {
                'section': 'frontmatter',
                'style': 'roman_lower_center',
                'start': 1
            },
            {
                'section': 'mainmatter',
                'style': 'arabic_dash',
                'start': 1
            }
        ]
    }
    response = requests.post(url, json=data)
    print(f"\ntest_page_footer_config_valid:")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}")

def test_page_footer_config_invalid():
    """测试无效的 page_footer_config（样式不在白名单中）"""
    url = f'{BASE_URL}/page_footer_config'
    data = {
        'value': [
            {
                'section': 'frontmatter',
                'style': 'invalid_style',
                'start': 1
            }
        ]
    }
    response = requests.post(url, json=data)
    print(f"\ntest_page_footer_config_invalid:")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}")

def test_get_data():
    """测试获取所有配置数据"""
    url = f'{BASE_URL}/get_data'
    response = requests.get(url)
    print(f"\ntest_get_data:")
    print(f"Status: {response.status_code}")
    data = response.json()
    if data['status'] == 'success':
        print("配置数据:")
        print(f"  toc_title: {data['data'].get('toc_title')}")
        print(f"  image_defaults: {data['data'].get('image_defaults')}")
        print(f"  formula_defaults: {data['data'].get('formula_defaults')}")
        print(f"  page_footer_config: {data['data'].get('page_footer_config')}")

if __name__ == '__main__':
    print("开始测试服务器配置路由...")
    
    # 先重置数据
    reset_url = f'{BASE_URL}/reset'
    requests.post(reset_url)
    print("\n已重置数据")
    
    # 运行所有测试
    test_toc_title()
    test_image_defaults()
    test_formula_defaults()
    test_page_footer_config_valid()
    test_page_footer_config_invalid()
    test_get_data()
    
    print("\n测试完成！")
