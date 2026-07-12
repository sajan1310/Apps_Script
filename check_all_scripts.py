import os
import subprocess

def check_all_scripts():
    if not os.path.exists("dist/index.html"):
        print("dist/index.html not found.")
        return
        
    with open("dist/index.html", "r", encoding="utf-8") as f:
        html = f.read()
        
    pos = 0
    idx = 1
    while True:
        # Search for any <script tag
        start_idx = html.find("<script", pos)
        if start_idx == -1:
            break
            
        tag_end = html.find(">", start_idx)
        if tag_end == -1:
            break
            
        tag_content = html[start_idx:tag_end]
        if "src=" in tag_content:
            pos = tag_end + 1
            continue
            
        content_start = tag_end + 1
        end_idx = html.find("</script>", content_start)
        if end_idx == -1:
            break
            
        script_content = html[content_start:end_idx].strip()
        
        temp_file = f"temp_check_{idx}.js"
        with open(temp_file, "w", encoding="utf-8") as f_temp:
            f_temp.write(script_content)
            
        res = subprocess.run(["node", "--check", temp_file], capture_output=True, text=True)
        if res.returncode == 0:
            print(f"Script tag {idx}: SUCCESS")
            if os.path.exists(temp_file):
                os.remove(temp_file)
        else:
            print(f"Script tag {idx}: FAILED")
            print(res.stderr)
            
        pos = end_idx + len("</script>")
        idx += 1

    # Cleanup remaining temp files
    for f in os.listdir("."):
        if f.startswith("temp_check_") and f.endswith(".js"):
            try:
                os.remove(f)
            except:
                pass

if __name__ == "__main__":
    check_all_scripts()
