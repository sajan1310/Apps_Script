import os
import re

def resolve_include(match):
    filename = match.group(1)
    path = f"{filename}.html"
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        # Recursively resolve includes inside the included file
        return re.sub(r'<\?!=?\s*include\([\'"]([^\'"]+)[\'"]\);\s*\?>', resolve_include, content)
    else:
        print(f"Warning: include file {path} not found")
        return f"<!-- Include failed: {filename} -->"

def compile_project():
    if not os.path.exists("Index.html"):
        print("Error: Index.html not found in current directory")
        return
    
    with open("Index.html", "r", encoding="utf-8") as f:
        content = f.read()
    
    compiled = re.sub(r'<\?!=?\s*include\([\'"]([^\'"]+)[\'"]\);\s*\?>', resolve_include, content)
    
    os.makedirs("dist", exist_ok=True)
    with open("dist/index.html", "w", encoding="utf-8") as f:
        f.write(compiled)
    print("Project compiled successfully to dist/index.html")

if __name__ == "__main__":
    compile_project()
