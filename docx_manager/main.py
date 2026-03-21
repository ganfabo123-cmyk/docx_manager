#!/usr/bin/env python3
"""
DocX Manager - 主入口
====================
启动 Flask 服务器，提供文档处理 API
"""

import argparse
from server import create_app


def main():
    parser = argparse.ArgumentParser(description="DocX Manager - 文档处理服务")
    parser.add_argument('--host', type=str, default='0.0.0.0', help='服务器监听地址')
    parser.add_argument('--port', type=int, default=5001, help='服务器监听端口')
    parser.add_argument('--debug', action='store_true', help='启用调试模式')
    parser.add_argument('--output', type=str, help='输出JSON文件路径')

    args = parser.parse_args()

    app = create_app(default_output_path=args.output)

    print(f"DocX Manager 服务启动中...")
    print(f"监听地址: {args.host}:{args.port}")
    print(f"调试模式: {args.debug}")
    print(f"\n可用的API端点:")
    print(f"  POST /save - 保存配置到文件")
    print(f"  POST /citations - 接收引用")
    print(f"  POST /recieve_right_style_docx - 接收并解析文档")
    print(f"  POST /generate_user_data - 生成用户数据并格式化文档")
    print(f"  GET  /download/<filename> - 下载生成的文档")
    print(f"  GET  /health - 健康检查")
    print(f"\n服务器正在运行，按 Ctrl+C 停止...")

    try:
        app.run(host=args.host, port=args.port, debug=args.debug)
    except KeyboardInterrupt:
        print("\n服务器已停止")


if __name__ == '__main__':
    main()
