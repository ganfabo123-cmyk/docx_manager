import os
import ast
from pathlib import Path

# === 配置 ===
IGNORE_DIRS = {'.git', '.idea', '__pycache__', '.venv', 'venv', 'env', 'logs', 'node_modules', 'dist', 'build'}
IGNORE_FILES = {'.DS_Store', 'Thumbs.db', '.gitignore', 'LICENSE'}
# 是否只显示公共成员（不显示 _开头的函数/变量）
ONLY_PUBLIC = True 

def get_node_signature(node):
    """获取函数/方法的签名字符串"""
    if hasattr(ast, 'unparse'):
        # Python 3.9+
        args = ast.unparse(node.args)
        returns = f" -> {ast.unparse(node.returns)}" if node.returns else ""
    else:
        args = "..."
        returns = ""
    
    # 移除 self 和 cls 参数以节省 token
    if args.startswith("self, "): args = args[6:]
    elif args.startswith("self"): args = ""
    elif args.startswith("cls, "): args = args[5:]
    elif args.startswith("cls"): args = ""
    
    return f"({args}){returns}"

def get_docstring_summary(node):
    """获取文档注释的第一行"""
    doc = ast.get_docstring(node)
    if doc:
        summary = doc.strip().split('\n')[0]
        # 截断过长的注释
        return f"  # {summary[:50]}..." if len(summary) > 50 else f"  # {summary}"
    return ""

def parse_py_structure(file_path: Path, prefix: str):
    """
    解析 py 文件，生成 类->方法 和 函数 的树状结构
    """
    results = []
    try:
        content = file_path.read_text(encoding='utf-8')
        tree = ast.parse(content)
    except Exception as e:
        return [f"{prefix}└── ⚠️ Parse Error: {str(e)}"]

    # 获取所有顶层节点
    definitions = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            definitions.append(node)

    # 排序：先类后函数，按行号
    definitions.sort(key=lambda x: (isinstance(x, ast.FunctionDef), x.lineno))

    pointers = [("├── ", "│   ")] * (len(definitions) - 1) + [("└── ", "    ")]

    for pointer, node in zip(pointers, definitions):
        is_async = "async " if isinstance(node, ast.AsyncFunctionDef) else ""
        
        # === 处理类 ===
        if isinstance(node, ast.ClassDef):
            # 获取基类信息
            bases = [ast.unparse(b) for b in node.bases] if hasattr(ast, 'unparse') else []
            base_str = f"({', '.join(bases)})" if bases else ""
            doc = get_docstring_summary(node)
            
            line = f"{prefix}{pointer[0]}C class {node.name}{base_str}{doc}"
            results.append(line)
            
            # 解析类内部的方法
            methods = [n for n in node.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
            if ONLY_PUBLIC:
                methods = [m for m in methods if not m.name.startswith('_') or m.name == '__init__']
            
            if methods:
                method_pointers = [("├── ", "│   ")] * (len(methods) - 1) + [("└── ", "    ")]
                class_prefix = prefix + pointer[1]
                for m_ptr, method in zip(method_pointers, methods):
                    m_sig = get_node_signature(method)
                    m_doc = get_docstring_summary(method)
                    m_async = "async " if isinstance(method, ast.AsyncFunctionDef) else ""
                    # m 代表 Method
                    results.append(f"{class_prefix}{m_ptr[0]}m {m_async}{method.name}{m_sig}{m_doc}")

        # === 处理函数 ===
        else:
            if ONLY_PUBLIC and node.name.startswith('_'):
                continue
            sig = get_node_signature(node)
            doc = get_docstring_summary(node)
            # f 代表 Function
            results.append(f"{prefix}{pointer[0]}f {is_async}{node.name}{sig}{doc}")

    return results

def generate_tree(dir_path: Path, prefix: str = ""):
    try:
        contents = list(dir_path.iterdir())
    except PermissionError:
        return

    contents.sort(key=lambda x: (not x.is_dir(), x.name))
    contents = [x for x in contents if x.name not in IGNORE_FILES and x.name not in IGNORE_DIRS]
    
    pointers = [("├── ", "│   ")] * (len(contents) - 1) + [("└── ", "    ")]
    
    for pointer, path in zip(pointers, contents):
        yield prefix + pointer[0] + path.name
        
        if path.is_dir():
            yield from generate_tree(path, prefix=prefix + pointer[1])
        elif path.suffix == '.py':
            # 缩进 deeper 进入文件内部
            file_structure = parse_py_structure(path, prefix=prefix + pointer[1])
            for line in file_structure:
                yield line

def save_optimized_tree(start_path=".", output_file="PROJECT_CONTEXT.txt"):
    root = Path(start_path)
    print(f"🌲 Generating optimized context for: {root.resolve().name}")
    
    header = [
        "Project Context Tree",
        "====================",
        "Legend:",
        "  C = Class",
        "  m = Method (inside Class)",
        "  f = Function (Global)",
        "  # = Docstring summary",
        "====================",
        f"{root.resolve().name}/"
    ]
    
    tree_lines = list(generate_tree(root))
    full_content = "\n".join(header + tree_lines)
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(full_content)
    
    print(full_content)
    print(f"\n✅ Optimized context saved to: {output_file}")

if __name__ == "__main__":
    save_optimized_tree(".")