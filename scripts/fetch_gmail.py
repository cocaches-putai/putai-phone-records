#!/usr/bin/env python3
"""
Gmail Fetcher Tool for Second Brain & GitHub Actions
Fetches emails and attachments using Gmail API (Read-only OAuth 2.0).
Supports both local files (token.json/credentials.json) and Cloud Environment Variables.
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

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

BASE_DIR = Path(__file__).resolve().parent.parent
CREDENTIALS_PATH = BASE_DIR / 'credentials.json'
TOKEN_PATH = BASE_DIR / 'token.json'


def authenticate(auth_only=False):
    """Authenticates the user with Gmail API using OAuth 2.0 (File or Env var)."""
    creds = None

    # 1. Try Environment Variable (for GitHub Actions Cloud)
    if os.environ.get('GMAIL_TOKEN_JSON'):
        try:
            token_info = json.loads(os.environ['GMAIL_TOKEN_JSON'])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
            print("[✓] 成功從雲端安全變數載入 Gmail 憑證。")
        except Exception as e:
            print(f"[!] 從環境變數讀取 GMAIL_TOKEN_JSON 失敗: {e}")
            creds = None

    # 2. Try Local File (for Mac local execution)
    if not creds and TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except Exception as e:
            print(f"[!] 讀取現有 token.json 失敗: {e}，將嘗試重新驗證。")
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
                creds = None

        if not creds:
            if os.environ.get('GMAIL_CREDENTIALS_JSON'):
                try:
                    client_config = json.loads(os.environ['GMAIL_CREDENTIALS_JSON'])
                    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                except Exception as e:
                    print(f"[X] 解析 GMAIL_CREDENTIALS_JSON 失敗: {e}")
                    sys.exit(1)
            elif CREDENTIALS_PATH.exists():
                flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            else:
                print(f"[X] 找不到憑證檔案：{CREDENTIALS_PATH}，亦無雲端環境變數。")
                sys.exit(1)

            if not os.environ.get('CI'):
                print("[*] 正在開啟瀏覽器進行 Google 帳號授權...")
                creds = flow.run_local_server(port=0)
                with open(TOKEN_PATH, 'w', encoding='utf-8') as token_file:
                    token_file.write(creds.to_json())
                print(f"[✓] 授權成功！認證權杖已安全儲存至 {TOKEN_PATH.name}")
            else:
                print("[X] CI 環境中憑證無效或缺少 refresh_token，請檢查 GitHub Secrets。")
                sys.exit(1)

    if auth_only:
        print("[✓] 授權驗證已完成！")
        return None

    return build('gmail', 'v1', credentials=creds)


def get_header(headers, name, default=''):
    for header in headers:
        if header.get('name', '').lower() == name.lower():
            return header.get('value', default)
    return default


def parse_parts(service, user_id, msg_id, parts, output_dir=None, downloaded_files=None):
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
                file_path = Path(output_dir) / filename
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                downloaded_files.append(str(file_path))
                print(f"  [下載附件] {filename} -> {file_path}")
            except HttpError as error:
                print(f"  [錯誤] 下載附件 {filename} 失敗: {error}")

        if mime_type == 'text/plain' and 'data' in body:
            text = base64.urlsafe_b64decode(body['data']).decode('utf-8', errors='replace')
            body_text.append(text)
        elif mime_type == 'text/html' and not body_text and 'data' in body:
            html = base64.urlsafe_b64decode(body['data']).decode('utf-8', errors='replace')
            body_text.append(html)

        if 'parts' in part:
            sub_text, sub_files = parse_parts(
                service, user_id, msg_id, part['parts'], output_dir, downloaded_files
            )
            body_text.extend(sub_text)

    return body_text, downloaded_files


def search_and_fetch_emails(service, query="電話紀錄", max_results=5, output_dir=None, download=True):
    try:
        results = service.users().messages().list(
            userId='me', q=query, maxResults=max_results
        ).execute()
        messages = results.get('messages', [])

        if not messages:
            print(f"[!] 找不到符合查詢 '{query}' 的郵件。")
            return []

        print(f"[*] 找到 {len(messages)} 封符合條件的郵件：")
        email_records = []

        for msg_summary in messages:
            msg_id = msg_summary['id']
            msg = service.users().messages().get(
                userId='me', id=msg_id, format='full'
            ).execute()

            payload = msg.get('payload', {})
            headers = payload.get('headers', [])

            subject = get_header(headers, 'Subject', '(無主旨)')
            sender = get_header(headers, 'From', '(未知寄件者)')
            date_str = get_header(headers, 'Date', '')
            internal_date_ms = int(msg.get('internalDate', 0))
            recv_time = datetime.fromtimestamp(internal_date_ms / 1000.0).strftime('%Y-%m-%d %H:%M:%S')

            print(f"\n- 郵件 ID: {msg_id}")
            print(f"  主旨: {subject}")
            print(f"  寄件者: {sender}")
            print(f"  接收時間: {recv_time}")

            downloaded_files = []
            body_texts = []

            if 'parts' in payload:
                body_texts, downloaded_files = parse_parts(
                    service, 'me', msg_id, payload['parts'],
                    output_dir=output_dir if download else None,
                    downloaded_files=downloaded_files
                )
            else:
                body_data = payload.get('body', {}).get('data', '')
                if body_data:
                    text = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='replace')
                    body_texts.append(text)

            email_records.append({
                'id': msg_id,
                'subject': subject,
                'sender': sender,
                'date': recv_time,
                'body': '\n'.join(body_texts).strip(),
                'attachments': downloaded_files
            })

        return email_records

    except HttpError as error:
        print(f"[X] 呼叫 Gmail API 時發生錯誤: {error}")
        return []

if __name__ == '__main__':
    service = authenticate()
    if service:
        search_and_fetch_emails(service, output_dir=str(BASE_DIR / 'Clippings'))
