#!/usr/bin/env python3
"""
Gmail Fetcher Tool for Second Brain
Fetches emails and attachments using Gmail API (Read-only OAuth 2.0).
"""

import os
import sys
import base64
import argparse
import json
from pathlib import Path
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Readonly scope for safety
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

BASE_DIR = Path(__file__).resolve().parent.parent
CREDENTIALS_PATH = BASE_DIR / 'credentials.json'
TOKEN_PATH = BASE_DIR / 'token.json'


def authenticate(auth_only=False):
    """Authenticates the user with Gmail API using OAuth 2.0 (supports File and Cloud Env var)."""
    creds = None

    # 1. Check Cloud Environment Variable (GitHub Actions)
    if os.environ.get('GMAIL_TOKEN_JSON'):
        try:
            token_info = json.loads(os.environ['GMAIL_TOKEN_JSON'])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
            print("[✓] 成功從雲端安全變數載入 Gmail 憑證。")
        except Exception as e:
            print(f"[!] 從環境變數讀取 GMAIL_TOKEN_JSON 失敗: {e}")
            creds = None

    # 2. Check Local File (Mac execution)
    if not creds and TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except Exception as e:
            print(f"[!] 讀取現有 token.json 失敗: {e}，將重新進行授權。")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                if TOKEN_PATH.parent.exists():
                    with open(TOKEN_PATH, 'w', encoding='utf-8') as token_file:
                        token_file.write(creds.to_json())
                print("[✓] 憑證已自動刷新成功。")
            except Exception as e:
                print(f"[!] 憑證刷新失敗: {e}")
                if os.environ.get('CI'):
                    return None

        if not creds:
            if not CREDENTIALS_PATH.exists():
                print(f"[X] 找不到憑證檔案：{CREDENTIALS_PATH}")
                print("請確認 credentials.json 已放置在專案目錄下。")
                if os.environ.get('CI'):
                    return None
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES
            )
            print("[*] 正在開啟瀏覽器進行 Google 帳號授權...")
            print("請在跳出的瀏覽器視窗中點擊登入並允許「唯讀」存取權限。")
            creds = flow.run_local_server(port=0)
            
            with open(TOKEN_PATH, 'w', encoding='utf-8') as token_file:
                token_file.write(creds.to_json())
            print(f"[✓] 授權成功！認證權杖已安全儲存至 {TOKEN_PATH.name}")

    if auth_only:
        print("[✓] 授權驗證已完成！")
        return None

    return build('gmail', 'v1', credentials=creds)


def get_header(headers, name, default=''):
    """Helper to extract header value by name."""
    for header in headers:
        if header.get('name', '').lower() == name.lower():
            return header.get('value', default)
    return default


def parse_parts(service, user_id, msg_id, parts, output_dir=None, downloaded_files=None):
    """Recursively parses email parts to extract body text and download attachments."""
    body_text = []
    if downloaded_files is None:
        downloaded_files = []

    for part in parts:
        mime_type = part.get('mimeType', '')
        filename = part.get('filename', '')
        body = part.get('body', {})

        if filename and body.get('attachmentId') and output_dir:
            attachment_id = body.get('attachmentId')
            try:
                attachment = service.users().messages().attachments().get(
                    userId=user_id, messageId=msg_id, id=attachment_id
                ).execute()
                file_data = base64.urlsafe_b64decode(attachment.get('data', ''))
                
                os.makedirs(output_dir, exist_ok=True)
                save_path = Path(output_dir) / filename
                
                # If file exists, append timestamp
                if save_path.exists():
                    stem = save_path.stem
                    suffix = save_path.suffix
                    ts = datetime.now().strftime("%H%M%S")
                    save_path = Path(output_dir) / f"{stem}_{ts}{suffix}"

                with open(save_path, 'wb') as f:
                    f.write(file_data)
                
                downloaded_files.append({
                    'filename': save_path.name,
                    'path': str(save_path),
                    'size_bytes': len(file_data)
                })
            except Exception as e:
                print(f"[!] 下載附件 {filename} 失敗: {e}")

        elif mime_type == 'text/plain' and body.get('data'):
            try:
                decoded = base64.urlsafe_b64decode(body.get('data')).decode('utf-8', errors='replace')
                body_text.append(decoded)
            except Exception:
                pass

        if 'parts' in part:
            nested_text, _ = parse_parts(service, user_id, msg_id, part['parts'], output_dir, downloaded_files)
            body_text.extend(nested_text)

    return body_text, downloaded_files


def search_and_fetch_emails(service, query="電話紀錄", max_results=5, output_dir=None, download=True):
    """Searches messages and retrieves details & attachments."""
    try:
        results = service.users().messages().list(
            userId='me', q=query, maxResults=max_results
        ).execute()
        messages = results.get('messages', [])

        if not messages:
            print(f"[*] 搜尋關鍵字「{query}」未找到任何相符的郵件。")
            return []

        print(f"[✓] 找到 {len(messages)} 封符合條件「{query}」的郵件：\n")
        emails_data = []

        for idx, msg_summary in enumerate(messages, 1):
            msg = service.users().messages().get(
                userId='me', id=msg_summary['id'], format='full'
            ).execute()

            payload = msg.get('payload', {})
            headers = payload.get('headers', [])

            subject = get_header(headers, 'Subject', '(無主旨)')
            sender = get_header(headers, 'From', '(未知寄件者)')
            recipient = get_header(headers, 'To', '')
            date_str = get_header(headers, 'Date', '')
            snippet = msg.get('snippet', '')

            # Parse attachments and body
            download_dest = output_dir if download else None
            body_text_parts = []
            downloaded = []

            if 'parts' in payload:
                body_text_parts, downloaded = parse_parts(
                    service, 'me', msg['id'], payload['parts'], download_dest, []
                )
            else:
                body_data = payload.get('body', {}).get('data', '')
                if body_data:
                    try:
                        decoded = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='replace')
                        body_text_parts.append(decoded)
                    except Exception:
                        pass

            email_entry = {
                'id': msg['id'],
                'thread_id': msg.get('threadId'),
                'subject': subject,
                'from': sender,
                'to': recipient,
                'date': date_str,
                'snippet': snippet,
                'body': "\n".join(body_text_parts).strip(),
                'attachments': downloaded
            }
            emails_data.append(email_entry)

            print(f"--- 郵件 {idx} ---")
            print(f"📌 主旨：{subject}")
            print(f"👤 寄件者：{sender}")
            print(f"📅 日期：{date_str}")
            if downloaded:
                print("📎 下載附件：")
                for att in downloaded:
                    print(f"   - {att['filename']} ({att['size_bytes']} bytes) -> {att['path']}")
            else:
                print("📎 附件：(無附件下載)")
            print(f"💬 摘要：{snippet[:120]}...\n")

        return emails_data

    except HttpError as error:
        print(f"[X] 呼叫 Gmail API 發生錯誤: {error}")
        return []


def main():
    parser = argparse.ArgumentParser(description="Gmail Fetcher for Second Brain")
    parser.add_argument("-q", "--query", default="電話紀錄", help="Gmail 搜尋語法 (預設: 電話紀錄)")
    parser.add_argument("-n", "--max-results", type=int, default=3, help="抓取郵件上限數量 (預設: 3)")
    parser.add_argument("-o", "--output-dir", default=str(BASE_DIR / "Clippings"), help="附件儲存資料夾 (預設: Clippings/)")
    parser.add_argument("--no-download", action="store_true", help="不自動下載附件")
    parser.add_argument("--auth-only", action="store_true", help="僅進行授權驗證")
    parser.add_argument("--json", action="store_true", help="輸出 JSON 格式")

    args = parser.parse_args()

    service = authenticate(auth_only=args.auth_only)
    if args.auth_only:
        return

    results = search_and_fetch_emails(
        service=service,
        query=args.query,
        max_results=args.max_results,
        output_dir=args.output_dir,
        download=not args.no_download
    )

    # Auto-regenerate mobile dashboard
    try:
        try:
            from scripts.generate_dashboard import main as update_dashboard
        except ImportError:
            from generate_dashboard import main as update_dashboard
        update_dashboard()
    except Exception as e:
        print(f"[!] 自動更新儀表板時發生錯誤: {e}")

    if args.json:
        print("\n--- JSON OUTPUT ---")
        print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
