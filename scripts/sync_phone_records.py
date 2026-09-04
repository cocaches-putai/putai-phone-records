#!/usr/bin/env python3
"""
一鍵全自動電話紀錄同步與發布腳本 (One-Click Phone Records Sync & Publish)
- 自動抓取 Gmail 最新電話紀錄附件
- 自動同步 Word 檔至網頁庫
- 自動重新編譯手機專用儀表板 (index.html)
- 自動推送到 GitHub Pages 上線
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

# 路徑設定
BASE_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = BASE_DIR / "scripts"
CLIPPINGS_DIR = BASE_DIR / "Clippings"
WEB_DIR = BASE_DIR / "web"
WEB_CLIPPINGS_DIR = WEB_DIR / "Clippings"

sys.path.insert(0, str(SCRIPTS_DIR))

def run_cmd(cmd, cwd=None, check=True):
    """執行 shell 命令並安全輸出"""
    result = subprocess.run(
        cmd,
        cwd=str(cwd or BASE_DIR),
        capture_output=True,
        text=True,
        shell=isinstance(cmd, str)
    )
    if check and result.returncode != 0:
        print(f"[!] 執行命令失敗: {cmd}")
        print(f"錯誤訊息: {result.stderr.strip()}")
    return result

def main():
    print("=" * 60)
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 🚀 啟動普台電話紀錄全自動同步發布程式")
    print("=" * 60)

    # 1. 抓取 Gmail 最新附件
    print("\n[步驟 1/5] 📬 檢查並下載 Gmail 最新「電話紀錄」附件...")
    try:
        from fetch_gmail import authenticate, search_and_fetch_emails
        service = authenticate()
        if service:
            CLIPPINGS_DIR.mkdir(parents=True, exist_ok=True)
            emails = search_and_fetch_emails(
                service=service,
                query="電話紀錄",
                max_results=10,
                output_dir=str(CLIPPINGS_DIR),
                download=True
            )
            print(f"[✓] 成功掃描 Gmail 郵件，附件已安全儲存至 Clippings/。")
        else:
            print("[!] Gmail 未授權，跳過郵件抓取步驟。")
    except Exception as e:
        print(f"[!] 抓取 Gmail 附件時發生例外: {e}")

    # 2. 同步 Word 附件至 web/Clippings/
    print("\n[步驟 2/5] 📂 同步 Word 檔案至網頁子目錄...")
    WEB_CLIPPINGS_DIR.mkdir(parents=True, exist_ok=True)
    copied_count = 0
    for docx_file in CLIPPINGS_DIR.glob("115.*.docx"):
        target_file = WEB_CLIPPINGS_DIR / docx_file.name
        if not target_file.exists() or docx_file.stat().st_size != target_file.stat().st_size:
            shutil.copy2(docx_file, target_file)
            copied_count += 1
    print(f"[✓] 已同步更新 {copied_count} 個 Word 檔案至 web/Clippings/。")

    # 3. 確保 Git 遠端狀態最新
    print("\n[步驟 3/5] 🔄 檢查 GitHub 遠端版本狀態...")
    run_cmd("git fetch origin", cwd=WEB_DIR, check=False)
    run_cmd("git reset --hard origin/main", cwd=WEB_DIR, check=False)

    # 確保剛剛拷貝的檔案仍在
    for docx_file in CLIPPINGS_DIR.glob("115.*.docx"):
        target_file = WEB_CLIPPINGS_DIR / docx_file.name
        if not target_file.exists():
            shutil.copy2(docx_file, target_file)

    # 4. 重新編譯儀表板 HTML
    print("\n[步驟 4/5] ⚙️  解析 Word 紀錄並重新編譯手機專用儀表板...")
    try:
        from generate_dashboard import parse_all_docx, generate_html
        data = parse_all_docx()
        generate_html(data)
        
        # 統計最新月份資訊
        latest_month = sorted(data.keys(), reverse=True)[0] if data else "無資料"
        month_records = len(data[latest_month]['records']) if latest_month in data else 0
        dates = [r['date'] for r in data[latest_month]['records']] if latest_month in data else []
        latest_date = max(dates) if dates else "未知"
        
        print(f"[✓] 編譯完成！最新月份：{latest_month}（共 {month_records} 筆通話，資料收錄至 {latest_date}）")
    except Exception as e:
        print(f"[!] 編譯儀表板時發生錯誤: {e}")
        sys.exit(1)

    # 5. 推送更新至 GitHub Pages
    print("\n[步驟 5/5] 🌐 自動發布更新至 GitHub Pages...")
    run_cmd("git add .", cwd=WEB_DIR)
    diff_res = run_cmd("git diff --staged --quiet", cwd=WEB_DIR, check=False)
    
    if diff_res.returncode == 0:
        print("[ℹ️] 網頁已是最新狀態，無需重複推送。")
    else:
        commit_msg = f"chore: auto-sync phone records to {latest_date} (total {month_records} records) [skip ci]"
        run_cmd(["git", "commit", "-m", commit_msg], cwd=WEB_DIR)
        push_res = run_cmd("git push origin main", cwd=WEB_DIR)
        if push_res.returncode == 0:
            print("[✓] 成功推送到 GitHub Pages 遠端！")
        else:
            print("[!] 推送失敗，請檢查網路連線。")

    print("\n" + "=" * 60)
    print(f"🎉 全部同步完成！最新網站：https://cocaches-putai.github.io/putai-phone-records/")
    print(f"📅 目前資料收錄至：{latest_date}（{latest_month} 月份已收錄 {month_records} 筆）")
    print("=" * 60)

if __name__ == "__main__":
    main()
