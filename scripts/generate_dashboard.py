#!/usr/bin/env python3
"""
Mobile Web Dashboard Generator for Putai Second Brain
Features:
- Real-Time Cloud Synchronization (Google Apps Script / Google Sheets API)
- Real-time cross-device horizontal communication: Any director's edit is visible to all within seconds
- Mobile-Optimized Teacher Chip Manager: 1-Tap Delete (✕) + Presets
- Department / Division Customization
- Time Order: Strictly Newest to Oldest (倒序排列，最新優先)
- Full User/Supervisor Label Tagging Control (🔴, 🟡, 🟢, ⚪)
- Dynamically loads from Obsidian Markdown Note: `知識庫/主管追蹤處置備註表.md`
"""

import os
import glob
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path
import docx

SCRIPT_DIR = Path(__file__).resolve().parent
if (SCRIPT_DIR.parent / "Clippings").exists():
    BASE_DIR = SCRIPT_DIR.parent
elif (SCRIPT_DIR.parent.parent / "Clippings").exists():
    BASE_DIR = SCRIPT_DIR.parent.parent
else:
    BASE_DIR = Path("/Users/lianjie/Desktop/普台第二大腦")

CLIPPINGS_DIR = BASE_DIR / "Clippings"
DAILY_DIR = BASE_DIR / "每日筆記" / "電話追蹤"
KNOWLEDGE_DIR = BASE_DIR / "知識庫"
TRACKING_MD = KNOWLEDGE_DIR / "主管追蹤處置備註表.md"
WEB_DIR = BASE_DIR / "web" if (BASE_DIR / "web").exists() else BASE_DIR
OUTPUT_HTML = WEB_DIR / "index.html"

GAS_SYNC_URL = "https://script.google.com/macros/s/AKfycby6runYRKHbWUK2e7WHKhfDx7oLV4YxjmRhU0V3lsqIVI5MKpqskGHS0tlZj8OgSQr8/exec"

CODE_MAP = {
    "1": "教務處",
    "2": "學務處",
    "3": "總務處",
    "4": "輔導室",
    "5": "國際部",
    "6": "音樂中心",
    "7": "人事室",
    "8": "住宿處",
    "9": "國中部導師",
    "10": "高中部導師",
    "11": "招生",
    "12": "其他"
}

def parse_tracking_markdown():
    """Dynamically parses `知識庫/主管追蹤處置備註表.md` into structured month data with strict teacher deduplication."""
    if not TRACKING_MD.exists():
        return {}

    content = TRACKING_MD.read_text(encoding="utf-8")
    months_cases = {}

    month_sections = re.split(r'##\s+📅\s+115\s*年\s*(\d+)\s*月[（\(]115\.(\d{2})[）\)]', content)
    
    i = 1
    while i < len(month_sections):
        month_key = f"115.{month_sections[i+1]}"
        sec_text = month_sections[i+2]
        i += 3

        cases = []
        case_blocks = re.split(r'###\s+([🔴🟡🟢⚪])\s+([A-Z0-9\-]+):\s+(.+)', sec_text)
        j = 1
        while j < len(case_blocks):
            icon = case_blocks[j]
            case_id = case_blocks[j+1]
            caller = case_blocks[j+2].strip()
            block_body = case_blocks[j+3]
            j += 4

            level = "red" if icon == "🔴" else "yellow" if icon == "🟡" else "green" if icon == "🟢" else "normal"
            level_text = "重點關懷" if level == "red" else "處室追蹤" if level == "yellow" else "已結案" if level == "green" else "常規業務"

            def get_field(pattern, default=""):
                m = re.search(pattern, block_body)
                return m.group(1).strip() if m else default

            date_val = get_field(r'-\s*\*\*日期\*\*[：:]\s*(.+)', f"{month_key}.01")
            raw_dept = get_field(r'-\s*\*\*業務處室\*\*[：:]\s*(.+)', "[12] 其他")
            content_val = get_field(r'-\s*\*\*事由內容\*\*[：:]\s*(.+)', "")
            action_val = get_field(r'-\s*\*\*處理說明\*\*[：:]\s*(.+)', "")
            teachers_raw = get_field(r'-\s*\*\*主責師長\*\*[：:]\s*(.+)', "")
            follow_up_val = get_field(r'-\s*\*\*處置追蹤\*\*[：:]\s*(.*)', "")

            code_match = re.search(r'\[(\d+)\]\s*(.+)', raw_dept)
            if code_match:
                code = code_match.group(1)
                dept_name = code_match.group(2)
            else:
                code = "12"
                dept_name = raw_dept

            raw_t_list = re.split(r'[、,，\s]+', teachers_raw)
            seen_teachers = set()
            teachers = []
            for t in raw_t_list:
                cleaned = t.strip().replace("[[", "").replace("]]", "")
                if cleaned and cleaned not in seen_teachers:
                    seen_teachers.add(cleaned)
                    teachers.append(cleaned)

            cases.append({
                "id": case_id,
                "date": date_val,
                "code": code,
                "dept_name": dept_name,
                "caller": caller,
                "level": level,
                "level_text": level_text,
                "content": content_val,
                "action": action_val,
                "teachers": teachers,
                "follow_up": follow_up_val
            })

        months_cases[month_key] = cases

    return months_cases

def parse_all_docx():
    files = sorted([f for f in glob.glob(str(CLIPPINGS_DIR / "115.*.docx")) if "_" not in f])
    months_data = {}
    structured_cases = parse_tracking_markdown()
    
    for f in files:
        doc = docx.Document(f)
        stem = Path(f).stem
        m = re.match(r'115\.(\d{2})', stem)
        if not m:
            continue
        month_num = m.group(1)
        month_key = f"115.{month_num}"
        month_label = f"115年{int(month_num)}月"
        
        if month_key not in months_data:
            months_data[month_key] = {
                "key": month_key,
                "label": month_label,
                "records": [],
                "key_cases": structured_cases.get(month_key, [])
            }
            
        for table in doc.tables:
            for row_idx, row in enumerate(table.rows[1:]):
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                dedup = []
                for c in cells:
                    if not dedup or c != dedup[-1]:
                        dedup.append(c)
                
                if len(dedup) >= 5 and dedup[0] not in ["業務編號", "業務\n編號", "業務 電話 統計", "業務電話統計", "記錄 人員", "記錄人員"]:
                    raw_code = dedup[0].replace("★1","").replace("★2","").replace("★3","").replace("★","").strip()
                    code = raw_code if raw_code in CODE_MAP else "12"
                    dept_name = CODE_MAP.get(code, "其他")
                    
                    time_str = dedup[1]
                    caller = dedup[2]
                    content = dedup[3]
                    action = dedup[4]
                    handler = dedup[5] if len(dedup) > 5 else ""
                    
                    raw_id = f"RAW-{month_num}-{len(months_data[month_key]['records'])+1:03d}"
                    
                    rec = {
                        "id": raw_id,
                        "date": stem,
                        "code": code,
                        "dept_name": dept_name,
                        "time": time_str,
                        "caller": caller,
                        "content": content,
                        "action": action,
                        "handler": handler,
                        "level": "normal",
                        "level_text": "常規",
                        "teachers": [],
                        "follow_up": ""
                    }
                    months_data[month_key]["records"].append(rec)

    for m_k in months_data.keys():
        if not months_data[m_k]["key_cases"] and m_k in structured_cases:
            months_data[m_k]["key_cases"] = structured_cases[m_k]

    return months_data

def generate_html(months_data):
    months_keys = sorted(months_data.keys(), reverse=True)
    months_json = json.dumps(months_data, ensure_ascii=False)
    months_keys_json = json.dumps(months_keys, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <title>普台高級中學 — 護學會電話紀錄與事項追蹤</title>
  <meta name="robots" content="noindex, nofollow, noarchive">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <meta name="theme-color" content="#1e3a8a">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --primary: #1e3a8a;
      --primary-light: #2563eb;
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --text: #0f172a;
      --text-muted: #64748b;
      --border: #e2e8f0;
      
      --red-bg: #fef2f2;
      --red-border: #fecaca;
      --red-text: #b91c1c;

      --yellow-bg: #fffbeb;
      --yellow-border: #fde68a;
      --yellow-text: #b45309;

      --green-bg: #f0fdf4;
      --green-border: #bbf7d0;
      --green-text: #15803d;

      --radius: 12px;
      --shadow-sm: 0 1px 3px rgba(0,0,0,0.05), 0 1px 2px rgba(0,0,0,0.03);
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      -webkit-tap-highlight-color: transparent;
    }}

    body {{
      font-family: 'Noto Sans TC', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding-bottom: 70px;
    }}

    /* Passcode Lock Screen Overlay */
    #lockOverlay {{
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(15, 23, 42, 0.96);
      backdrop-filter: blur(12px);
      z-index: 9999;
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 20px;
    }}

    .lock-card {{
      background: #ffffff;
      width: 100%;
      max-width: 360px;
      padding: 28px 24px;
      border-radius: 18px;
      text-align: center;
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3);
      animation: fadeIn 0.3s ease-out;
    }}

    .lock-icon {{
      font-size: 2.8rem;
      margin-bottom: 12px;
    }}

    .lock-title {{
      font-size: 1.15rem;
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 4px;
    }}

    .lock-desc {{
      font-size: 0.82rem;
      color: #64748b;
      margin-bottom: 20px;
    }}

    .lock-input {{
      width: 100%;
      padding: 12px 14px;
      border: 2px solid #cbd5e1;
      border-radius: 10px;
      font-size: 1.05rem;
      text-align: center;
      letter-spacing: 2px;
      margin-bottom: 14px;
      outline: none;
      transition: border-color 0.2s;
    }}

    .lock-input:focus {{
      border-color: #2563eb;
    }}

    .lock-btn {{
      width: 100%;
      background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
      color: white;
      border: none;
      padding: 12px;
      border-radius: 10px;
      font-size: 0.95rem;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.3);
    }}

    .lock-btn:active {{
      transform: scale(0.98);
    }}

    .lock-error {{
      font-size: 0.8rem;
      color: #dc2626;
      margin-top: 10px;
      display: none;
    }}

    @keyframes shake {{
      0%, 100% {{ transform: translateX(0); }}
      20%, 60% {{ transform: translateX(-6px); }}
      40%, 80% {{ transform: translateX(6px); }}
    }}

    /* Header */
    header {{
      background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
      color: white;
      padding: 16px;
      position: sticky;
      top: 0;
      z-index: 100;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }}

    .header-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }}

    .school-title {{
      font-size: 1.05rem;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .cloud-badge {{
      font-size: 0.68rem;
      background: rgba(34, 197, 94, 0.2);
      border: 1px solid rgba(34, 197, 94, 0.4);
      color: #86efac;
      padding: 2px 6px;
      border-radius: 12px;
      display: inline-flex;
      align-items: center;
      gap: 3px;
      font-weight: 500;
    }}

    .month-select {{
      background: rgba(255,255,255,0.18);
      border: 1px solid rgba(255,255,255,0.3);
      color: white;
      padding: 4px 10px;
      border-radius: 20px;
      font-size: 0.82rem;
      font-weight: 600;
      outline: none;
      cursor: pointer;
    }}

    .month-select option {{
      background: #0f172a;
      color: white;
    }}

    .search-wrapper {{
      position: relative;
    }}

    .search-input {{
      width: 100%;
      background: rgba(255, 255, 255, 0.95);
      border: none;
      padding: 8px 14px 8px 34px;
      border-radius: 8px;
      font-size: 0.88rem;
      color: #1e293b;
      outline: none;
    }}

    .search-input::placeholder {{
      color: #94a3b8;
    }}

    .search-icon {{
      position: absolute;
      left: 10px;
      top: 50%;
      transform: translateY(-50%);
      color: #64748b;
      font-size: 0.85rem;
      pointer-events: none;
    }}

    /* View Switcher Bar (Key Highlights vs All) */
    .view-switcher-bar {{
      background: #ffffff;
      padding: 8px 16px;
      border-bottom: 1px solid var(--border);
      display: flex;
      gap: 8px;
      align-items: center;
      justify-content: space-between;
    }}

    .switch-group {{
      display: flex;
      background: #f1f5f9;
      padding: 2px;
      border-radius: 8px;
      width: 100%;
    }}

    .switch-btn {{
      flex: 1;
      border: none;
      background: transparent;
      padding: 6px 10px;
      font-size: 0.8rem;
      font-weight: 600;
      color: #64748b;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.15s;
    }}

    .switch-btn.active {{
      background: #ffffff;
      color: var(--primary);
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }}

    /* Stats bar */
    .quick-stats-row {{
      background: #ffffff;
      padding: 8px 12px;
      display: flex;
      gap: 6px;
      justify-content: space-around;
      border-bottom: 1px solid var(--border);
      text-align: center;
    }}

    .q-stat {{
      flex: 1;
      font-size: 0.72rem;
      color: var(--text-muted);
      cursor: pointer;
      padding: 6px 4px;
      border-radius: 8px;
      border: 1px solid transparent;
      transition: all 0.15s;
    }}

    .q-stat:active {{
      transform: scale(0.96);
    }}

    .q-stat strong {{
      display: block;
      font-size: 0.98rem;
      color: #0f172a;
    }}

    .q-stat.active-total {{
      background: #eff6ff;
      border-color: #93c5fd;
    }}

    .q-stat.active-red {{
      background: #fef2f2;
      border-color: #fca5a5;
    }}

    .q-stat.active-yellow {{
      background: #fffbeb;
      border-color: #fcd34d;
    }}

    .q-stat.active-green {{
      background: #f0fdf4;
      border-color: #86efac;
    }}

    /* 12 Codes Statistics Panel */
    .stat-collapse-btn {{
      width: 100%;
      background: #ffffff;
      border: none;
      border-bottom: 1px solid var(--border);
      padding: 8px 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 0.78rem;
      font-weight: 700;
      color: var(--primary);
      cursor: pointer;
    }}

    .official-stats-panel {{
      background: #ffffff;
      border-bottom: 1px solid var(--border);
      padding: 8px 16px 12px 16px;
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 6px;
    }}

    .stat-chip {{
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 6px;
      padding: 5px 3px;
      text-align: center;
      cursor: pointer;
      transition: all 0.15s;
    }}

    .stat-chip:active, .stat-chip.active {{
      background: #eff6ff;
      border-color: #3b82f6;
    }}

    .chip-code {{
      font-size: 0.65rem;
      color: var(--text-muted);
    }}

    .chip-name {{
      font-size: 0.74rem;
      font-weight: 600;
      color: #1e293b;
    }}

    .chip-count {{
      font-size: 0.82rem;
      font-weight: 700;
      color: var(--primary-light);
    }}

    /* Main Container */
    main {{
      padding: 12px 16px;
      max-width: 760px;
      margin: 0 auto;
    }}

    .section-title {{
      font-size: 0.8rem;
      font-weight: 700;
      color: var(--text-muted);
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}

    .sort-indicator {{
      font-size: 0.7rem;
      color: #2563eb;
      background: #eff6ff;
      padding: 2px 6px;
      border-radius: 4px;
      font-weight: 600;
    }}

    /* Card Item */
    .card-list {{
      display: flex;
      flex-direction: column;
      gap: 10px;
    }}

    .case-card {{
      background: var(--card-bg);
      border-radius: var(--radius);
      padding: 14px;
      box-shadow: var(--shadow-sm);
      border: 1px solid var(--border);
      position: relative;
    }}

    .case-card.highlight-red {{
      border-left: 4px solid #dc2626;
      background: #fffafa;
    }}

    .case-card.highlight-yellow {{
      border-left: 4px solid #d97706;
      background: #fffdfa;
    }}

    .case-card.highlight-green {{
      border-left: 4px solid #16a34a;
      background: #fcfdfc;
    }}

    .card-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 6px;
    }}

    .code-badge {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      background: #e2e8f0;
      color: #334155;
      font-size: 0.7rem;
      font-weight: 700;
      padding: 2px 7px;
      border-radius: 4px;
    }}

    .code-badge.c1 {{ background: #dbeafe; color: #1e40af; }}
    .code-badge.c2 {{ background: #fce7f3; color: #9d174d; }}
    .code-badge.c3 {{ background: #fef3c7; color: #92400e; }}
    .code-badge.c4 {{ background: #e0e7ff; color: #3730a3; }}
    .code-badge.c8 {{ background: #ede9fe; color: #5b21b6; }}
    .code-badge.c9 {{ background: #ccfbf1; color: #115e59; }}
    .code-badge.c11 {{ background: #fae8ff; color: #86198f; }}

    .card-time {{
      font-size: 0.75rem;
      color: var(--text-muted);
      font-weight: 500;
    }}

    .caller-name {{
      font-size: 0.95rem;
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 4px;
    }}

    .content-box {{
      font-size: 0.85rem;
      color: #334155;
      margin-bottom: 8px;
    }}

    .action-box {{
      font-size: 0.8rem;
      color: #475569;
      background: #f8fafc;
      padding: 6px 10px;
      border-radius: 6px;
      margin-bottom: 6px;
      border-left: 2px solid #cbd5e1;
    }}

    .follow-up-box {{
      font-size: 0.82rem;
      color: #0c4a6e;
      background: #f0f9ff;
      padding: 8px 10px;
      border-radius: 6px;
      border: 1px dashed #7dd3fc;
      margin-top: 6px;
    }}

    .follow-up-box.pending {{
      color: #92400e;
      background: #fffbeb;
      border-color: #fde68a;
      font-style: italic;
    }}

    .teachers-tags {{
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
      margin-top: 6px;
    }}

    .teacher-pill {{
      font-size: 0.72rem;
      background: #f1f5f9;
      color: #1e293b;
      padding: 1px 6px;
      border-radius: 10px;
      border: 1px solid #cbd5e1;
    }}

    .card-footer {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 8px;
      padding-top: 6px;
      border-top: 1px solid #f1f5f9;
      font-size: 0.72rem;
      color: var(--text-muted);
    }}

    .btn-reply {{
      background: #eff6ff;
      border: 1px solid #93c5fd;
      color: #1e40af;
      padding: 4px 10px;
      border-radius: 6px;
      font-size: 0.74rem;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 3px;
    }}

    .btn-reply:active {{
      background: #dbeafe;
    }}

    /* Supervisor Tag & Edit Modal */
    #replyModal {{
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(15, 23, 42, 0.7);
      backdrop-filter: blur(4px);
      z-index: 10000;
      display: flex;
      justify-content: center;
      align-items: flex-end;
    }}

    .reply-sheet {{
      background: #ffffff;
      width: 100%;
      max-width: 500px;
      border-radius: 20px 20px 0 0;
      padding: 18px 16px 26px 16px;
      box-shadow: 0 -10px 25px rgba(0,0,0,0.2);
      animation: slideUp 0.25s ease-out;
      max-height: 90vh;
      overflow-y: auto;
    }}

    @keyframes slideUp {{
      from {{ transform: translateY(100%); }}
      to {{ transform: translateY(0); }}
    }}

    .sheet-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }}

    .sheet-title {{
      font-size: 1.05rem;
      font-weight: 700;
      color: #0f172a;
    }}

    .sheet-close {{
      background: #f1f5f9;
      border: none;
      width: 28px;
      height: 28px;
      border-radius: 50%;
      font-size: 1rem;
      color: #64748b;
      cursor: pointer;
    }}

    .sheet-case-name {{
      background: #f8fafc;
      padding: 6px 10px;
      border-radius: 8px;
      font-size: 0.82rem;
      font-weight: 600;
      color: #1e293b;
      margin-bottom: 10px;
      border: 1px solid var(--border);
    }}

    .form-group {{
      margin-bottom: 10px;
      text-align: left;
    }}

    .form-label {{
      font-size: 0.76rem;
      font-weight: 700;
      color: #475569;
      margin-bottom: 4px;
      display: block;
    }}

    /* Interactive Teacher Tag Chip Container */
    .teacher-chip-container {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      min-height: 36px;
      background: #f8fafc;
      border: 1px dashed #cbd5e1;
      border-radius: 8px;
      padding: 6px 8px;
      align-items: center;
    }}

    .edit-teacher-chip {{
      background: #e0e7ff;
      color: #1e40af;
      border: 1px solid #bfdbfe;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 0.76rem;
      font-weight: 600;
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }}

    .chip-del-btn {{
      color: #dc2626;
      font-weight: 800;
      font-size: 0.82rem;
      cursor: pointer;
      padding: 0 2px;
      border-radius: 50%;
    }}

    .chip-del-btn:active {{
      background: #fee2e2;
    }}

    .btn-add-teacher {{
      background: var(--primary-light);
      color: white;
      border: none;
      padding: 0 14px;
      border-radius: 8px;
      font-size: 0.8rem;
      font-weight: 700;
      cursor: pointer;
      white-space: nowrap;
    }}

    .btn-add-teacher:active {{
      transform: scale(0.96);
    }}

    .quick-preset-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 4px;
      align-items: center;
      margin-top: 6px;
    }}

    .preset-label {{
      font-size: 0.68rem;
      color: #64748b;
      font-weight: 600;
    }}

    .preset-pill {{
      background: #f1f5f9;
      border: 1px solid #cbd5e1;
      font-size: 0.7rem;
      padding: 2px 7px;
      border-radius: 10px;
      cursor: pointer;
      color: #334155;
    }}

    .preset-pill:active {{
      background: #dbeafe;
      border-color: #93c5fd;
      color: #1e40af;
    }}

    /* 4-Way Label Tag Selector */
    .tag-selector-grid {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 6px;
    }}

    .tag-btn-option {{
      background: #f8fafc;
      border: 1.5px solid #cbd5e1;
      padding: 8px 6px;
      border-radius: 8px;
      text-align: center;
      font-size: 0.78rem;
      font-weight: 600;
      color: #334155;
      cursor: pointer;
      transition: all 0.15s;
    }}

    .tag-btn-option.selected-red {{
      background: #fef2f2;
      border-color: #dc2626;
      color: #b91c1c;
      font-weight: 700;
    }}

    .tag-btn-option.selected-yellow {{
      background: #fffbeb;
      border-color: #d97706;
      color: #b45309;
      font-weight: 700;
    }}

    .tag-btn-option.selected-green {{
      background: #f0fdf4;
      border-color: #16a34a;
      color: #15803d;
      font-weight: 700;
    }}

    .tag-btn-option.selected-normal {{
      background: #f1f5f9;
      border-color: #64748b;
      color: #475569;
      font-weight: 700;
    }}

    .form-input, .form-textarea {{
      width: 100%;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 0.88rem;
      outline: none;
    }}

    .form-input:focus, .form-textarea:focus {{
      border-color: var(--primary-light);
    }}

    .form-textarea {{
      height: 60px;
      resize: none;
    }}

    .sheet-submit-btn {{
      width: 100%;
      background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
      color: white;
      border: none;
      padding: 12px;
      border-radius: 10px;
      font-size: 0.92rem;
      font-weight: 700;
      cursor: pointer;
      margin-top: 6px;
    }}

    .sheet-submit-btn:active {{
      transform: scale(0.98);
    }}

    .empty-state {{
      text-align: center;
      padding: 40px 20px;
      color: var(--text-muted);
    }}
  </style>
</head>
<body>

  <!-- Passcode Lock Modal -->
  <div id="lockOverlay">
    <div class="lock-card" id="lockCard">
      <div class="lock-icon">🏫</div>
      <div class="lock-title">普台高級中學</div>
      <div class="lock-desc">護學會電話紀錄追蹤系統</div>
      <input type="password" id="passcodeInput" class="lock-input" placeholder="請輸入通關密碼" autofocus onkeydown="if(event.key==='Enter') verifyPasscode()">
      <button class="lock-btn" onclick="verifyPasscode()">登入查看</button>
      <div class="lock-error" id="lockError">密碼錯誤，請重新輸入</div>
    </div>
  </div>

  <!-- Header -->
  <header id="mainApp" style="display:none;">
    <div class="header-top">
      <div class="school-title">
        <span>📞</span> 普台高中 護學會電話紀錄
        <span class="cloud-badge" id="cloudStatusBadge">🟢 雲端同步中</span>
      </div>
      <select id="monthSelect" class="month-select" onchange="changeMonth(this.value)">
        <!-- Options injected via JS -->
      </select>
    </div>
    <div class="search-wrapper">
      <span class="search-icon">🔍</span>
      <input type="text" id="searchInput" class="search-input" placeholder="搜尋學生、來電事由、師長或分機..." oninput="handleSearch()">
    </div>
  </header>

  <!-- App Body -->
  <div id="appBody" style="display:none;">
    
    <!-- View Switcher -->
    <div class="view-switcher-bar">
      <div class="switch-group">
        <button class="switch-btn active" id="btnViewKey" onclick="switchView('key')">🎯 關鍵追蹤事項 (<span id="countKey">0</span>)</button>
        <button class="switch-btn" id="btnViewAll" onclick="switchView('all')">📋 全部通話明細 (<span id="countAll">0</span>)</button>
      </div>
    </div>

    <!-- Quick Stats Row -->
    <div class="quick-stats-row">
      <div class="q-stat" id="qStatTotal" onclick="filterByStatus('all')">
        <strong id="statMonthTotal">0</strong>
        <span>本月總通話</span>
      </div>
      <div class="q-stat" id="qStatRed" onclick="filterByStatus('red')">
        <strong style="color:#dc2626;" id="statMonthRed">0</strong>
        <span>重點關懷</span>
      </div>
      <div class="q-stat" id="qStatYellow" onclick="filterByStatus('yellow')">
        <strong style="color:#d97706;" id="statMonthYellow">0</strong>
        <span>處室追蹤</span>
      </div>
      <div class="q-stat" id="qStatGreen" onclick="filterByStatus('green')">
        <strong style="color:#16a34a;" id="statMonthGreen">0</strong>
        <span>已結案</span>
      </div>
    </div>

    <!-- Official 12 Codes Statistics Panel -->
    <button class="stat-collapse-btn" onclick="toggleStats()">
      <span>📊 護學會 12 業務電話統計表（點擊可快速過濾）</span>
      <span id="statArrow">▶</span>
    </button>

    <div class="official-stats-panel" id="statsPanel" style="display:none;">
      <!-- Chips injected via JS -->
    </div>

    <!-- Content List -->
    <main>
      <div class="section-title">
        <div>
          <span id="listTitle">關鍵追蹤事項清單</span>
          <span class="sort-indicator">⏱️ 最新優先</span>
        </div>
        <span style="font-size:0.75rem; color:#94a3b8;" id="matchCount">共 0 筆</span>
      </div>

      <div class="card-list" id="cardsContainer">
        <!-- Injected via JS -->
      </div>
    </main>
  </div>

  <!-- Supervisor Interactive Tag & Edit Sheet -->
  <div id="replyModal" style="display:none;">
    <div class="reply-sheet">
      <div class="sheet-header">
        <div class="sheet-title">🏷️ 編輯處室單位、主責師長與標籤</div>
        <button class="sheet-close" onclick="closeReplyModal()">✕</button>
      </div>

      <div class="sheet-case-name" id="modalCaseName">案件名稱</div>

      <!-- Department Selector -->
      <div class="form-group">
        <label class="form-label">🏢 負責業務處室／細部組別：</label>
        <div style="display: flex; gap: 8px;">
          <select id="modalDeptCodeSelect" class="form-input" style="flex: 1.2; font-weight: 600;">
            <option value="1">[1] 教務處</option>
            <option value="2">[2] 學務處</option>
            <option value="3">[3] 總務處</option>
            <option value="4">[4] 輔導室</option>
            <option value="5">[5] 國際部</option>
            <option value="6">[6] 音樂中心</option>
            <option value="7">[7] 人事室</option>
            <option value="8">[8] 住宿處</option>
            <option value="9">[9] 國中部導師</option>
            <option value="10">[10] 高中部導師</option>
            <option value="11">[11] 招生組</option>
            <option value="12">[12] 其他</option>
          </select>
          <input type="text" id="modalSubDeptInput" class="form-input" style="flex: 1;" placeholder="細部組別（如：出納組）">
        </div>
      </div>

      <!-- 4-Way Label Selector -->
      <div class="form-group">
        <label class="form-label">📌 自選案件列管標籤：</label>
        <div class="tag-selector-grid">
          <div class="tag-btn-option" id="tagOptRed" onclick="selectTagOption('red')">
            🔴 重點關懷（校級）
          </div>
          <div class="tag-btn-option" id="tagOptYellow" onclick="selectTagOption('yellow')">
            🟡 處室追蹤（待辦）
          </div>
          <div class="tag-btn-option" id="tagOptGreen" onclick="selectTagOption('green')">
            🟢 已結案（完成）
          </div>
          <div class="tag-btn-option" id="tagOptNormal" onclick="selectTagOption('normal')">
            ⚪ 常規業務（不列管）
          </div>
        </div>
      </div>

      <!-- Mobile-Friendly Teacher Chips Manager -->
      <div class="form-group">
        <label class="form-label">👤 主責師長名單（點 ✕ 可直接刪除）：</label>
        <div id="modalTeacherChips" class="teacher-chip-container">
          <!-- Injected via JS chips -->
        </div>
        <div style="display: flex; gap: 6px; margin-top: 6px;">
          <input type="text" id="newTeacherInput" class="form-input" placeholder="輸入師長姓名或職稱..." style="flex: 1;" onkeydown="if(event.key==='Enter'){{event.preventDefault();addModalTeacher();}}">
          <button type="button" class="btn-add-teacher" onclick="addModalTeacher()">＋加入</button>
        </div>
        <!-- Quick Preset Pills -->
        <div class="quick-preset-row">
          <span class="preset-label">快捷新增：</span>
          <span class="preset-pill" onclick="quickAddTeacher('洪誌陽（副校長）')">副校長</span>
          <span class="preset-pill" onclick="quickAddTeacher('連傑（教務主任）')">教務主任</span>
          <span class="preset-pill" onclick="quickAddTeacher('辜漢宗（學務主任）')">學務主任</span>
          <span class="preset-pill" onclick="quickAddTeacher('張月馨（教務副主任）')">月馨副主任</span>
          <span class="preset-pill" onclick="quickAddTeacher('陳怡婷（住宿副主任）')">住宿副主任</span>
          <span class="preset-pill" onclick="quickAddTeacher('羅可強（生輔組長）')">生輔組長</span>
          <span class="preset-pill" onclick="quickAddTeacher('出納組')">出納組</span>
          <span class="preset-pill" onclick="quickAddTeacher('國中註冊組')">註冊組</span>
          <span class="preset-pill" onclick="quickAddTeacher('輔導室')">輔導室</span>
        </div>
      </div>

      <!-- Follow up note -->
      <div class="form-group">
        <label class="form-label">💬 處置追蹤文字（可自由輸入與修改）：</label>
        <textarea id="modalActionInput" class="form-textarea" placeholder="請填寫最新處置、回電情形或追蹤備註..."></textarea>
      </div>

      <button class="sheet-submit-btn" onclick="submitReply()">確認儲存處置與師長</button>
    </div>
  </div>

  <script>
    const CORRECT_PASS = "putai";
    const GAS_SYNC_URL = "{GAS_SYNC_URL}";
    const monthsData = {months_json};
    const monthKeys = {months_keys_json};
    const CODE_NAMES = {json.dumps(CODE_MAP, ensure_ascii=False)};

    let currentMonth = monthKeys[0] || '115.08';
    let currentView = 'key'; // 'key' or 'all'
    let currentStatusFilter = null; // null, 'red', 'yellow', 'green'
    let currentCodeFilter = null; // null or '1'..'12'
    let searchQuery = '';
    let activeCaseId = null;
    let selectedModalTag = 'yellow';
    let currentModalTeachers = [];

    function checkAuth() {{
      const saved = localStorage.getItem("putai_auth_v2");
      if (saved === "ok") {{
        unlockApp();
      }}
    }}

    function verifyPasscode() {{
      const val = document.getElementById("passcodeInput").value.trim().toLowerCase();
      if (val === CORRECT_PASS) {{
        localStorage.setItem("putai_auth_v2", "ok");
        unlockApp();
      }} else {{
        const err = document.getElementById("lockError");
        const card = document.getElementById("lockCard");
        err.style.display = "block";
        card.style.animation = "shake 0.3s";
        setTimeout(() => {{ card.style.animation = ""; }}, 300);
      }}
    }}

    function unlockApp() {{
      document.getElementById("lockOverlay").style.display = "none";
      document.getElementById("mainApp").style.display = "block";
      document.getElementById("appBody").style.display = "block";
      initMonthsDropdown();
      applySavedSupervisorFeedbacks();
      render();
      
      // Fetch live cloud database from Google Apps Script
      fetchCloudLiveSync();
      // Auto-poll live sync every 20 seconds
      setInterval(fetchCloudLiveSync, 20000);
    }}

    function initMonthsDropdown() {{
      const select = document.getElementById("monthSelect");
      select.innerHTML = monthKeys.map(k => `
        <option value="${{k}}" ${{k === currentMonth ? 'selected' : ''}}>${{monthsData[k].label}}</option>
      `).join('');
    }}

    function changeMonth(m) {{
      currentMonth = m;
      currentStatusFilter = null;
      currentCodeFilter = null;
      currentView = 'key';
      resetStatStyles();
      document.getElementById("btnViewKey").classList.add("active");
      document.getElementById("btnViewAll").classList.remove("active");
      document.getElementById("listTitle").innerText = '關鍵追蹤事項清單';
      render();
    }}

    function switchView(view) {{
      currentView = view;
      currentStatusFilter = null;
      currentCodeFilter = null;
      resetStatStyles();
      document.getElementById("btnViewKey").classList.toggle("active", view === 'key');
      document.getElementById("btnViewAll").classList.toggle("active", view === 'all');
      document.getElementById("listTitle").innerText = (view === 'key') ? '關鍵追蹤事項清單' : '全部通話紀錄明細';
      render();
    }}

    function filterByStatus(status) {{
      if (status === 'all') {{
        switchView('all');
        document.getElementById('qStatTotal').classList.add('active-total');
        return;
      }}

      currentView = 'key';
      currentCodeFilter = null;
      document.getElementById("btnViewKey").classList.add("active");
      document.getElementById("btnViewAll").classList.remove("active");

      if (currentStatusFilter === status) {{
        currentStatusFilter = null;
        resetStatStyles();
        document.getElementById("listTitle").innerText = '關鍵追蹤事項清單';
      }} else {{
        currentStatusFilter = status;
        resetStatStyles();

        if (status === 'red') {{
          document.getElementById('qStatRed').classList.add('active-red');
          document.getElementById('listTitle').innerText = '🔴 重點關懷事項清單';
        }} else if (status === 'yellow') {{
          document.getElementById('qStatYellow').classList.add('active-yellow');
          document.getElementById('listTitle').innerText = '🟡 處室追蹤事項清單';
        }} else if (status === 'green') {{
          document.getElementById('qStatGreen').classList.add('active-green');
          document.getElementById('listTitle').innerText = '🟢 已結案事項清單';
        }}
      }}
      render();
    }}

    function resetStatStyles() {{
      ['qStatTotal', 'qStatRed', 'qStatYellow', 'qStatGreen'].forEach(id => {{
        const el = document.getElementById(id);
        if (el) el.className = 'q-stat';
      }});
    }}

    function filterByCode(code) {{
      if (currentCodeFilter === code) {{
        currentCodeFilter = null;
        currentView = 'all';
        document.getElementById("listTitle").innerText = '全部通話紀錄明細';
      }} else {{
        currentCodeFilter = code;
        currentStatusFilter = null;
        currentView = 'all';
        resetStatStyles();
        document.getElementById("btnViewKey").classList.remove("active");
        document.getElementById("btnViewAll").classList.add("active");
        document.getElementById("listTitle").innerText = `[${{code}}] ${{CODE_NAMES[code]}} 通話紀錄`;
      }}
      render();
    }}

    function handleSearch() {{
      searchQuery = document.getElementById("searchInput").value;
      render();
    }}

    function toggleStats() {{
      const panel = document.getElementById("statsPanel");
      const arrow = document.getElementById("statArrow");
      if (panel.style.display === "none") {{
        panel.style.display = "grid";
        arrow.innerText = "▼";
      }} else {{
        panel.style.display = "none";
        arrow.innerText = "▶";
      }}
    }}

    /* 4-Way Tag Selection UI Handler */
    function selectTagOption(tag) {{
      selectedModalTag = tag;
      ['red', 'yellow', 'green', 'normal'].forEach(t => {{
        const cap = t.charAt(0).toUpperCase() + t.slice(1);
        const el = document.getElementById('tagOpt' + cap);
        if (el) {{
          el.className = (t === tag) ? `tag-btn-option selected-${{t}}` : 'tag-btn-option';
        }}
      }});
    }}

    /* Interactive Mobile Teacher Chip Functions */
    function renderModalTeacherChips() {{
      const container = document.getElementById('modalTeacherChips');
      if (currentModalTeachers.length === 0) {{
        container.innerHTML = `<span style="font-size:0.75rem; color:#94a3b8; font-style:italic;">尚未指定師長（請由下方輸入或點選快捷）</span>`;
        return;
      }}
      container.innerHTML = currentModalTeachers.map((t, idx) => `
        <span class="edit-teacher-chip">
          ${{t}}
          <span class="chip-del-btn" onclick="removeModalTeacher('${{t}}')">✕</span>
        </span>
      `).join('');
    }}

    function removeModalTeacher(name) {{
      currentModalTeachers = currentModalTeachers.filter(t => t !== name);
      renderModalTeacherChips();
    }}

    function addModalTeacher() {{
      const input = document.getElementById('newTeacherInput');
      const val = input.value.trim().replace(/^\[\[/, '').replace(/\]\]$/, '');
      if (val && !currentModalTeachers.includes(val)) {{
        currentModalTeachers.push(val);
        renderModalTeacherChips();
        input.value = '';
      }}
      input.focus();
    }}

    function quickAddTeacher(name) {{
      if (!currentModalTeachers.includes(name)) {{
        currentModalTeachers.push(name);
        renderModalTeacherChips();
      }}
    }}

    function openReplyModal(caseId, caseName, teachersArray, currentFollowUp, currentLevel, currentCode, currentDeptName) {{
      activeCaseId = caseId;
      document.getElementById('modalCaseName').innerText = `📋 案件：${{caseName}}`;
      
      currentModalTeachers = Array.from(new Set(teachersArray || []));
      renderModalTeacherChips();
      document.getElementById('newTeacherInput').value = '';

      document.getElementById('modalActionInput').value = currentFollowUp || '';
      
      const codeSelect = document.getElementById('modalDeptCodeSelect');
      codeSelect.value = currentCode || '12';
      
      let subDept = '';
      if (currentDeptName) {{
        const parts = currentDeptName.split('/');
        if (parts.length > 1) {{
          subDept = parts.slice(1).join('/').trim();
        }}
      }}
      document.getElementById('modalSubDeptInput').value = subDept;

      selectTagOption(currentLevel || 'yellow');
      document.getElementById('replyModal').style.display = 'flex';
    }}

    function closeReplyModal() {{
      document.getElementById('replyModal').style.display = 'none';
    }}

    /* Submit with Real-Time Google Apps Script Background Transmission */
    function submitReply() {{
      const action = document.getElementById('modalActionInput').value.trim();
      const code = document.getElementById('modalDeptCodeSelect').value;
      const subDept = document.getElementById('modalSubDeptInput').value.trim();
      
      const mainDeptName = CODE_NAMES[code] || '其他';
      const fullDeptName = subDept ? `${{mainDeptName}} / ${{subDept}}` : mainDeptName;

      const feedback = {{
        caseId: activeCaseId,
        month: currentMonth,
        code: code,
        dept_name: fullDeptName,
        teachers: currentModalTeachers,
        action: action,
        status: selectedModalTag,
        timestamp: new Date().toISOString()
      }};

      // 1. Instant local memory update (0.01s visual feedback)
      let allFeedbacks = JSON.parse(localStorage.getItem('putai_supervisor_feedbacks') || '[]');
      allFeedbacks = allFeedbacks.filter(f => f.caseId !== activeCaseId);
      allFeedbacks.push(feedback);
      localStorage.setItem('putai_supervisor_feedbacks', JSON.stringify(allFeedbacks));

      applyFeedbackToMemory(feedback);
      closeReplyModal();
      render();

      // 2. Real-time background cloud transmission to Google Sheet
      const badge = document.getElementById('cloudStatusBadge');
      if (badge) badge.innerText = '⏳ 雲端同步中...';

      fetch(GAS_SYNC_URL, {{
        method: 'POST',
        headers: {{ 'Content-Type': 'text/plain;charset=utf-8' }},
        body: JSON.stringify(feedback)
      }})
      .then(resp => resp.json())
      .then(data => {{
        if (badge) badge.innerText = '🟢 雲端同步完成';
        setTimeout(() => {{ if (badge) badge.innerText = '🟢 雲端即時連線'; }}, 3000);
      }})
      .catch(err => {{
        console.warn('GAS sync offline fallback:', err);
        if (badge) badge.innerText = '🟡 本地已保存（待連線）';
      }});

      alert('✓ 處置已儲存並即時同步至全校主管雲端！');
    }}

    function applyFeedbackToMemory(fb) {{
      const mKey = fb.month ? (fb.month.toString().includes('.') ? fb.month.toString() : '115.' + fb.month.toString().padStart(2,'0')) : currentMonth;
      const mData = monthsData[mKey] || monthsData[currentMonth];
      if (!mData) return;

      let foundKey = mData.key_cases.find(c => c.id === fb.caseId);
      if (foundKey) {{
        if (fb.status === 'normal') {{
          mData.key_cases = mData.key_cases.filter(c => c.id !== fb.caseId);
        }} else {{
          foundKey.level = fb.status;
          foundKey.level_text = (fb.status === 'red') ? '重點關懷' : (fb.status === 'yellow') ? '處室追蹤' : '已結案';
          if (fb.code) foundKey.code = fb.code.toString();
          if (fb.dept_name) foundKey.dept_name = fb.dept_name;
          if (fb.teachers) foundKey.teachers = Array.from(new Set(fb.teachers));
          if (fb.action !== undefined) {{
            foundKey.follow_up = fb.action;
          }}
        }}
      }} else if (fb.status !== 'normal') {{
        let rawRec = mData.records.find(r => r.id === fb.caseId);
        if (rawRec) {{
          const newCase = {{
            id: rawRec.id,
            date: rawRec.date,
            code: fb.code ? fb.code.toString() : rawRec.code,
            dept_name: fb.dept_name || rawRec.dept_name,
            caller: rawRec.caller,
            level: fb.status,
            level_text: (fb.status === 'red') ? '重點關懷' : (fb.status === 'yellow') ? '處室追蹤' : '已結案',
            content: rawRec.content,
            action: rawRec.action,
            teachers: fb.teachers ? Array.from(new Set(fb.teachers)) : [],
            follow_up: fb.action || ''
          }};
          mData.key_cases.push(newCase);
        }}
      }}
    }}

    function applySavedSupervisorFeedbacks() {{
      const saved = JSON.parse(localStorage.getItem('putai_supervisor_feedbacks') || '[]');
      saved.forEach(fb => applyFeedbackToMemory(fb));
    }}

    /* Real-Time Cloud Sync Pull */
    function fetchCloudLiveSync() {{
      fetch(GAS_SYNC_URL)
        .then(resp => resp.json())
        .then(cloudList => {{
          if (Array.isArray(cloudList) && cloudList.length > 0) {{
            cloudList.forEach(fb => applyFeedbackToMemory(fb));
            render();
            const badge = document.getElementById('cloudStatusBadge');
            if (badge) badge.innerText = '🟢 雲端即時連線';
          }}
        }})
        .catch(err => {{
          console.log('Cloud sync standby');
        }});
    }}

    function render() {{
      const mData = monthsData[currentMonth] || {{ records: [], key_cases: [] }};
      const records = mData.records;
      const keyCases = mData.key_cases;

      const redCount = keyCases.filter(r => r.level === 'red').length;
      const yellowCount = keyCases.filter(r => r.level === 'yellow').length;
      const greenCount = keyCases.filter(r => r.level === 'green').length;

      document.getElementById("countKey").innerText = keyCases.length;
      document.getElementById("countAll").innerText = records.length;
      document.getElementById("statMonthTotal").innerText = records.length;
      document.getElementById("statMonthRed").innerText = redCount;
      document.getElementById("statMonthYellow").innerText = yellowCount;
      document.getElementById("statMonthGreen").innerText = greenCount;

      const codeCounts = {{}};
      for (let c in CODE_NAMES) codeCounts[c] = 0;
      records.forEach(r => {{
        codeCounts[r.code] = (codeCounts[r.code] || 0) + 1;
      }});

      const chipsPanel = document.getElementById("statsPanel");
      chipsPanel.innerHTML = Object.keys(CODE_NAMES).map(code => `
        <div class="stat-chip ${{currentCodeFilter === code ? 'active' : ''}}" onclick="filterByCode('${{code}}')">
          <div class="chip-code">${{code}}</div>
          <div class="chip-name">${{CODE_NAMES[code]}}</div>
          <div class="chip-count">${{codeCounts[code] || 0}}</div>
        </div>
      `).join('');

      let sourceList = (currentView === 'key') ? keyCases : records;
      
      const filtered = sourceList.filter(item => {{
        if (currentStatusFilter && item.level !== currentStatusFilter) return false;
        if (currentCodeFilter && item.code !== currentCodeFilter) return false;

        if (searchQuery.trim()) {{
          const q = searchQuery.toLowerCase();
          const fullText = (item.date + (item.time||'') + item.caller + item.content + item.action + item.dept_name + (item.teachers||[]).join(' ') + (item.follow_up||'')).toLowerCase();
          return fullText.includes(q);
        }}
        return true;
      }});

      // Sort Newest to Oldest (倒序排列，最新優先)
      filtered.sort((a, b) => {{
        const dA = (a.date || '').replace(/\\D/g, '');
        const dB = (b.date || '').replace(/\\D/g, '');
        if (dA !== dB) return dB.localeCompare(dA);
        
        const tA = (a.time || '').replace(/\\D/g, '').padEnd(4, '0');
        const tB = (b.time || '').replace(/\\D/g, '').padEnd(4, '0');
        if (tA !== tB) return tB.localeCompare(tA);
        
        return (b.id || '').localeCompare(a.id || '');
      }});

      document.getElementById("matchCount").innerText = `共 ${{filtered.length}} 筆`;

      const container = document.getElementById("cardsContainer");
      if (filtered.length === 0) {{
        container.innerHTML = `
          <div class="empty-state">
            <div style="font-size:2rem;">📭</div>
            <p>沒有符合條件的紀錄</p>
          </div>
        `;
        return;
      }}

      container.innerHTML = filtered.map(item => {{
        const isRed = item.level === 'red';
        const isYellow = item.level === 'yellow';
        const isGreen = item.level === 'green';
        const highlightClass = isRed ? 'highlight-red' : isYellow ? 'highlight-yellow' : isGreen ? 'highlight-green' : '';

        const uniqueTeachers = Array.from(new Set(item.teachers || []));
        const teachersHtml = (uniqueTeachers.length > 0) ? `
          <div class="teachers-tags">
            <span style="font-size:0.72rem; color:#64748b; font-weight:600;">👤 主責師長：</span>
            ${{uniqueTeachers.map(t => `<span class="teacher-pill">${{t}}</span>`).join('')}}
          </div>
        ` : '';

        let followUpHtml = '';
        if (item.follow_up) {{
          followUpHtml = `
            <div class="follow-up-box">
              <strong>💬 處置追蹤：</strong> ${{item.follow_up}}
            </div>
          `;
        }} else if (isRed || isYellow) {{
          followUpHtml = `
            <div class="follow-up-box pending">
              <strong>⏳ 處置追蹤：</strong> 尚未有處室主管／導師回報（請點右下方設定標籤回報）
            </div>
          `;
        }}

        const safeFollowUp = (item.follow_up || '').replace(/'/g, "\\'");
        const safeDeptName = (item.dept_name || '').replace(/'/g, "\\'");
        const teachersJsonStr = JSON.stringify(uniqueTeachers).replace(/"/g, '&quot;');

        const replyButtonHtml = `
          <button class="btn-reply" onclick="openReplyModal('${{item.id}}', '${{item.caller}}', ${{teachersJsonStr}}, '${{safeFollowUp}}', '${{item.level}}', '${{item.code}}', '${{safeDeptName}}')">
            🏷️ 編輯處室與師長
          </button>
        `;

        return `
          <div class="case-card ${{highlightClass}}">
            <div class="card-top">
              <span class="code-badge c${{item.code}}">
                [${{item.code}}] ${{item.dept_name}}
              </span>
              <span class="card-time">📅 ${{item.date}} ${{item.time || ''}}</span>
            </div>

            <div class="caller-name">
              ${{isRed ? '🔴 ' : isYellow ? '🟡 ' : isGreen ? '🟢 ' : ''}}${{item.caller}}
            </div>

            <div class="content-box">
              ${{item.content}}
            </div>

            <div class="action-box">
              <strong>處理說明：</strong>${{item.action}}
            </div>

            ${{teachersHtml}}
            ${{followUpHtml}}

            <div class="card-footer">
              <span>處置代號：${{item.code}} ${{item.dept_name}}</span>
              ${{replyButtonHtml}}
            </div>
          </div>
        `;
      }}).join('');
    }}

    // Check authentication on load
    checkAuth();
  </script>
</body>
</html>
"""
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[✓] 成功產出支援 Google 雲端即時連線的主管追蹤儀表板：{OUTPUT_HTML}")

def main():
    WEB_DIR.mkdir(parents=True, exist_ok=True)
    months_data = parse_all_docx()
    generate_html(months_data)

if __name__ == "__main__":
    main()
