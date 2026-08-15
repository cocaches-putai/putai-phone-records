#!/usr/bin/env python3
"""
Mobile Web Dashboard Generator for Putai Second Brain & GitHub Actions
"""

import os
import glob
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path
import docx

# Detect if running in root repo or subfolder
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent

# Check potential directories
if (REPO_ROOT / "Clippings").exists():
    BASE_DIR = REPO_ROOT
elif (REPO_ROOT.parent / "Clippings").exists():
    BASE_DIR = REPO_ROOT.parent
else:
    BASE_DIR = Path("/Users/lianjie/Desktop/普台第二大腦")

CLIPPINGS_DIR = BASE_DIR / "Clippings"
KNOWLEDGE_DIR = BASE_DIR / "知識庫"
TRACKING_MD = KNOWLEDGE_DIR / "主管追蹤處置備註表.md"
OUTPUT_HTML = REPO_ROOT / "index.html" if (REPO_ROOT / "index.html").exists() or not (BASE_DIR / "web").exists() else BASE_DIR / "web" / "index.html"

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
    # (Read the master html template)
    from scripts.generate_dashboard import generate_html as gen_master
    # Run master generator to OUTPUT_HTML
    from generate_dashboard import generate_html as local_gen
