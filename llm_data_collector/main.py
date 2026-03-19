import argparse
import json
from pathlib import Path
from llm_data_collector.core.server import create_app

def main():
    parser = argparse.ArgumentParser(description="LLM数据收集器")
    parser.add_argument('--host', type=str, default='0.0.0.0')
    parser.add_argument('--port', type=int, default=5001)
    parser.add_argument('--debug', action='store_true')
    parser.add_argument('--output', type=str, help='输出JSON文件路径')

    args = parser.parse_args()

    # 2. 将 args.output 传给 create_app
    app = create_app(default_output_path=args.output)

    print(f"LLM数据收集器启动中...")
    print(f"监听地址: {args.host}:{args.port}")
    print(f"调试模式: {args.debug}")
    print(f"\n可用的API端点:")
    print(f"  POST /_doc - 接收文档描述")
    print(f"  POST /page_footer_config - 接收页脚配置")
    print(f"  POST /toc_mode - 接收目录模式")
    print(f"  POST /toc_entries - 接收目录条目")
    print(f"  POST /content_section - 接收章节内容")
    print(f"  POST /content_toc - 接收目录内容")
    print(f"  POST /content_heading1 - 接收一级标题")
    print(f"  POST /content_heading2 - 接收二级标题")
    print(f"  POST /content_heading3 - 接收三级标题")
    print(f"  POST /content_body - 接收正文")
    print(f"  POST /content_table - 接收表格")
    print(f"  POST /content_formula - 接收公式")
    print(f"  POST /content_image - 接收图片")
    print(f"  POST /references - 接收参考文献")
    print(f"  POST /citations - 接收引用")
    print(f"  GET /get_data - 获取完整用户数据")
    print(f"  POST /reset - 重置数据")
    print(f"  GET /health - 健康检查")
    print(f"\n服务器正在运行，按 Ctrl+C 停止...")

    try:
        app.run(host=args.host, port=args.port, debug=args.debug)
    except KeyboardInterrupt:
        print("\n服务器已停止")

        if args.output:
            from llm_data_collector.core.server import collector
            data = collector.get_user_data()
            output_path = Path(args.output)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"数据已保存到: {output_path}")


if __name__ == '__main__':
    main()
