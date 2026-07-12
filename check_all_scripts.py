import os
import subprocess

def check_scripts_in_file(dist_path, label):
    if not os.path.exists(dist_path):
        print(f"{dist_path} not found.")
        return

    with open(dist_path, "r", encoding="utf-8") as f:
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

        temp_file = f"temp_check_{label}_{idx}.js"
        with open(temp_file, "w", encoding="utf-8") as f_temp:
            f_temp.write(script_content)

        res = subprocess.run(["node", "--check", temp_file], capture_output=True, text=True)
        if res.returncode == 0:
            print(f"[{label}] Script tag {idx}: SUCCESS")
            if os.path.exists(temp_file):
                os.remove(temp_file)
        else:
            print(f"[{label}] Script tag {idx}: FAILED")
            print(res.stderr)

        pos = end_idx + len("</script>")
        idx += 1

def check_all_scripts():
    check_scripts_in_file("dist/index.html", "desktop")
    check_scripts_in_file("dist/mobile.html", "mobile")

    # Cleanup remaining temp files
    for f in os.listdir("."):
        if f.startswith("temp_check_") and f.endswith(".js"):
            try:
                os.remove(f)
            except:
                pass

if __name__ == "__main__":
    check_all_scripts()
