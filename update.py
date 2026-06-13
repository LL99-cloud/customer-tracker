# 一鍵更新業績儀表板
# 用法：把業績 Excel 檔「拖曳」到「更新業績.bat」圖示上放開即可。
# 流程：讀取 Excel → 產生 data.js → （若已設定好 GitHub）自動上傳。
import sys, os, subprocess, datetime

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

def run(cmd):
    return subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")

def line():
    print("=" * 52)

line()
print(" 客戶業績儀表板　一鍵更新")
line()

# --- 找出要轉換的 Excel 檔（拖曳進來的第一個 xlsx/xlsm）---
xlsx = next((a for a in sys.argv[1:] if a.lower().endswith((".xlsx", ".xlsm"))), None)
if not xlsx:
    print()
    print("  找不到 Excel 檔案。")
    print("  請把業績 Excel 檔『拖曳』到「更新業績.bat」圖示上後放開，")
    print("  不要用滑鼠雙擊。")
    print()
    sys.exit(1)
if not os.path.exists(xlsx):
    print(f"\n  檔案不存在：{xlsx}\n")
    sys.exit(1)

print(f"\n  來源檔案：{os.path.basename(xlsx)}")

# --- 步驟 1：轉換 ---
print("\n[1/3] 讀取 Excel、產生 data.js …")
# -X utf8 確保子程式輸出為 UTF-8，避免中文亂碼
r = run([sys.executable, "-X", "utf8", "convert_xlsx.py", xlsx])
if r.returncode != 0:
    print("  轉換失敗：")
    print("  " + (r.stderr or r.stdout).strip().replace("\n", "\n  "))
    print()
    sys.exit(1)
print("  完成　" + r.stdout.strip())

# --- 檢查 GitHub 是否已設定好 ---
has_identity = bool(run(["git", "config", "user.name"]).stdout.strip())
has_remote = bool(run(["git", "remote"]).stdout.strip())

if not (has_identity and has_remote):
    print()
    line()
    print("  data.js 已更新完成，但尚未上傳到 GitHub：")
    if not has_identity:
        print("   ・還沒設定 git 身分（user.name / user.email）")
    if not has_remote:
        print("   ・還沒連結到 GitHub repository")
    print("  完成一次性的 GitHub 設定後，以後拖曳就會自動上傳。")
    line()
    sys.exit(0)

# --- 步驟 2：記錄變更 ---
print("\n[2/3] 記錄變更 …")
run(["git", "add", "data.js"])
today = datetime.date.today().isoformat()
commit = run(["git", "commit", "-m", f"更新業績資料 {today}"])
if "nothing to commit" in (commit.stdout + commit.stderr):
    print("  資料沒有變化，不需要上傳。")
    sys.exit(0)
print("  完成")

# --- 步驟 3：上傳 ---
print("\n[3/3] 上傳到 GitHub …")
push = run(["git", "push"])
if push.returncode != 0:
    print("  上傳失敗：")
    print("  " + (push.stderr or push.stdout).strip().replace("\n", "\n  "))
    print()
    sys.exit(1)

print()
line()
print("  全部完成！約一分鐘後，網站就會顯示新資料。")
line()
