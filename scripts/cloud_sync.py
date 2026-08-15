#!/usr/bin/env python3
"""
Cloud Automated Daily Sync Runner for GitHub Actions
1. Authenticates with Gmail API via GitHub Secrets (GMAIL_TOKEN_JSON)
2. Fetches latest "電話紀錄" emails and downloads docx to Clippings/
3. Re-generates mobile dashboard index.html
"""

import os
import sys
from pathlib import Path
from datetime import datetime

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

def main():
    print("=" * 50)
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 啟動 GitHub Actions 雲端電話紀錄同步機器人")
    print("=" * 50)

    # 1. Fetch latest email attachments
    try:
        from fetch_gmail import authenticate, search_and_fetch_emails
        service = authenticate()
        if service:
            clippings_dir = REPO_ROOT / "Clippings"
            clippings_dir.mkdir(parents=True, exist_ok=True)
            print(f"[*] 正在從 Gmail 搜尋最新電話紀錄...")
            emails = search_and_fetch_emails(
                service=service,
                query="電話紀錄",
                max_results=5,
                output_dir=str(clippings_dir),
                download=True
            )
            print(f"[✓] 成功檢查/下載 {len(emails)} 封郵件。")
        else:
            print("[!] Gmail 未授權，跳過郵件抓取步驟。")
    except Exception as e:
        print(f"[!] 抓取 Gmail 附件時發生錯誤: {e}")

    # 2. Re-compile dashboard index.html
    try:
        from generate_dashboard import main as update_dashboard
        print("[*] 正在重新編譯手機版網頁 (index.html)...")
        update_dashboard()
        print("[✓] 網頁編譯完成！")
    except Exception as e:
        print(f"[!] 編譯儀表板時發生錯誤: {e}")
        sys.exit(1)

    print("\n[✓] 雲端同步任務圓滿結束！")

if __name__ == "__main__":
    main()
