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

def compile_template(template_name, output_name):
    template_path = f"{template_name}.html"
    if not os.path.exists(template_path):
        print(f"Error: {template_path} not found in current directory")
        return

    with open(template_path, "r", encoding="utf-8") as f:
        content = f.read()

    compiled = re.sub(r'<\?!=?\s*include\([\'"]([^\'"]+)[\'"]\);\s*\?>', resolve_include, content)

    os.makedirs("dist", exist_ok=True)
    out_path = f"dist/{output_name}"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(compiled)
    print(f"{template_path} compiled successfully to {out_path}")

def compile_project():
    # Desktop shell (main.js#doGet default branch)
    compile_template("Index", "index.html")
    # Mobile shell (main.js#doGet's ?ui=mobile branch) -- same include()
    # resolver, just a different entry template, so Playwright verify
    # scripts have a real compiled file to load for mobile-only screens.
    compile_template("Mobile_Index", "mobile.html")

if __name__ == "__main__":
    compile_project()
