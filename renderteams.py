# teams_sender_ui_simplified_poc.py
# Simplified UI as requested:
# - One message box only
# - One checkbox for @Everyone
# - One checkbox for POC mention from Excel
# - If POC mention checkbox is selected:
#     * If POC text/email from Excel is already in message, convert it to real Teams mention
#     * Otherwise append the POC mention at bottom automatically
# - Optional App/CTASK table
# - Group-to-group delay
#
# Excel supported format:
# App ID | App Name | Group Name | CTASK ID | CTASK Group | Chat ID | POC Name
# Optional: POC Email
#
# Install once:
#   pip install flask pandas openpyxl requests
# Run:
#   python teams_sender_ui_simplified_poc.py
# Open:
#   http://127.0.0.1:5000

import html
import os
import re
import shutil
import time
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import requests
from flask import Flask, request, render_template_string, send_file

GRAPH = "https://graph.microsoft.com/v1.0"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

REQUIRED_COLUMNS = ["App ID", "App Name", "CTASK ID", "CTASK Group"]
PREFERRED_GROUP_COLUMNS = ["Teams Group Name", "Group Name", "CTASK Group"]
POC_NAME_COLUMNS = ["POC Name", "Poc Name", "POC", "POC Names"]
POC_EMAIL_COLUMNS = ["POC Email", "POC Mail ID", "POC Mail Id", "POC Email ID", "POC Mail", "POC UPN", "POC UserPrincipalName"]
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
INLINE_MENTION_PATTERN = re.compile(r"@\{([^|{}]+)\|([^{}|]+)\}")

app = Flask(__name__)
LAST_EXCEL_FILE = None
LAST_DISCOVERED_FILE = None
GRAPH_TOKEN = None
USER_CACHE = {}

DEFAULT_MESSAGE_HTML = (
    'Hi @Everyone,<br><br>'
    'Please proceed with pre-piv, attach artefacts to CTASK and move to '
    '&ldquo;In progress&rdquo; once done.<br>- CHG0690343 &ndash; Sat 04 July'
)
LAST_FORM_STATE = {
    "message_html": DEFAULT_MESSAGE_HTML,
    "everyone": True,
    "mention_poc": False,
    "include_table": True,
    "delay_seconds": "5",
}

HTML = """
<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Teams Message Sender</title>
    <style>
        :root {
            --brand: #5B5FC7;
            --brand-dark: #464775;
            --bg: #F5F6FB;
            --card: #FFFFFF;
            --border: #E1E4EA;
            --text: #242424;
            --muted: #6B6F76;
            --danger: #C4314B;
            --success: #237B4B;
        }
        * { box-sizing: border-box; }
        body {
            font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif;
            margin: 0; padding: 32px 16px 64px;
            background: linear-gradient(180deg, #EEF0FA 0%, var(--bg) 220px);
            color: var(--text);
        }
        .page { max-width: 880px; margin: 0 auto; }
        .app-header { display: flex; align-items: center; gap: 14px; margin-bottom: 28px; }
        .app-header .logo {
            width: 44px; height: 44px; border-radius: 10px;
            background: var(--brand); color: #fff; display: flex; align-items: center;
            justify-content: center; font-size: 20px; font-weight: 700; flex-shrink: 0;
        }
        .app-header h1 { font-size: 22px; margin: 0; }
        .app-header p { margin: 2px 0 0; color: var(--muted); font-size: 13px; }

        .box {
            background: var(--card); padding: 22px 24px; border-radius: 12px;
            margin-bottom: 20px; border: 1px solid var(--border);
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }
        .box-title { display: flex; align-items: center; gap: 10px; margin-bottom: 4px; }
        .step-badge {
            width: 26px; height: 26px; border-radius: 50%; background: var(--brand); color: #fff;
            font-size: 13px; font-weight: 700; display: flex; align-items: center; justify-content: center;
            flex-shrink: 0;
        }
        .box-title h2 { font-size: 16px; margin: 0; }
        .box-desc { color: var(--muted); font-size: 13px; margin: 4px 0 16px 36px; line-height: 1.5; }

        label.field-label { font-weight: 600; font-size: 13px; display: block; margin-top: 14px; margin-bottom: 6px; color: var(--text); }
        input[type=file], select, input[type=password], input[type=text] {
            width: 100%; padding: 9px 12px; border: 1px solid var(--border); border-radius: 8px;
            font-size: 14px; font-family: inherit; background: #fff;
        }
        input[type=password]:focus, input[type=text]:focus, select:focus { outline: 2px solid var(--brand); outline-offset: 1px; }

        .token-row { display: flex; gap: 8px; align-items: stretch; margin-top: 6px; }
        .token-row input { flex: 1; font-family: Consolas, monospace; letter-spacing: 1px; }
        .icon-btn {
            background: #F0F1F8; border: 1px solid var(--border); border-radius: 8px; padding: 0 14px;
            cursor: pointer; font-size: 13px; color: var(--text); white-space: nowrap;
        }
        .icon-btn:hover { background: #E4E6F5; }

        button, .btn {
            background: var(--brand); color: #fff; border: none; padding: 10px 18px;
            border-radius: 8px; margin-top: 16px; cursor: pointer; margin-right: 8px;
            font-size: 14px; font-weight: 600;
        }
        button:hover { background: var(--brand-dark); }
        button.secondary { background: #fff; color: var(--brand); border: 1px solid var(--brand); }
        button.secondary:hover { background: #F0F1F8; }
        button.danger-outline { background: #fff; color: var(--danger); border: 1px solid var(--danger); }
        button.danger-outline:hover { background: #FCECEC; }

        .status-pill {
            display: inline-flex; align-items: center; gap: 6px; font-size: 13px;
            padding: 6px 12px; border-radius: 999px; margin-top: 12px; font-weight: 600;
        }
        .status-ok { background: #E6F4EA; color: var(--success); }
        .status-warn { background: #FCE8EA; color: var(--danger); }
        .dot { width: 8px; height: 8px; border-radius: 50%; background: currentColor; display:inline-block; }

        .checkbox-row { display: flex; align-items: center; gap: 8px; margin-top: 12px; font-size: 14px; }
        .checkbox-row input { width: auto; margin: 0; }

        /* Rich text editor */
        .editor-toolbar {
            display: flex; gap: 4px; margin-top: 6px; padding: 6px; background: #F7F8FC;
            border: 1px solid var(--border); border-bottom: none; border-radius: 8px 8px 0 0; flex-wrap: wrap;
        }
        .editor-toolbar button {
            margin: 0; padding: 6px 11px; font-size: 13px; font-weight: 700; background: #fff;
            border: 1px solid var(--border); color: var(--text); border-radius: 6px;
        }
        .editor-toolbar button:hover { background: #ECEEF9; }
        .editor-toolbar button.italic { font-style: italic; font-weight: 400; }
        .editor-toolbar button.underline { text-decoration: underline; font-weight: 400; }
        .editor-toolbar .sep { width: 1px; background: var(--border); margin: 2px 4px; }

        #message_editor {
            min-height: 200px; max-height: 480px; overflow-y: auto;
            border: 1px solid var(--border); border-radius: 0 0 8px 8px; padding: 14px;
            font-size: 14px; line-height: 1.55; background: #fff;
        }
        #message_editor:focus { outline: 2px solid var(--brand); outline-offset: -1px; }
        #message_editor table { border-collapse: collapse; margin: 10px 0; max-width: 100%; }
        #message_editor td, #message_editor th { border: 1px solid #ccc; padding: 5px 9px; font-size: 13px; }
        #message_editor th { background: #F5F6FB; font-weight: 700; }

        .hint { color: var(--muted); font-size: 12.5px; line-height: 1.5; margin-top: 8px; }
        pre.log {
            background: #1E1E1E; color: #D4D4D4; padding: 16px; white-space: pre-wrap;
            border-radius: 8px; font-size: 13px; line-height: 1.5; max-height: 480px; overflow-y: auto;
        }
        .preview { border: 1px solid var(--border); padding: 14px; border-radius: 8px; background: #FAFAFD; margin-top: 8px; }
        .preview table { border-collapse: collapse; }
        .preview td, .preview th { border: 1px solid #ccc; padding: 4px 8px; }
        .mention { background: #FFF2B2; color: #000; padding: 1px 4px; border-radius: 4px; font-weight: 600; }
        .chat-tag {
            display: inline-block; background: #ECEEF9; color: var(--brand-dark); font-size: 12px;
            font-weight: 700; padding: 3px 9px; border-radius: 6px; margin-bottom: 8px;
        }
        a { color: var(--brand); }
        .links a { margin-right: 16px; font-weight: 600; font-size: 13px; }
    </style>
    <script>
        function toggleTokenVisibility() {
            var input = document.getElementById('token_input');
            var btn = document.getElementById('token_toggle_btn');
            if (input.type === 'password') {
                input.type = 'text';
                btn.textContent = 'Hide';
            } else {
                input.type = 'password';
                btn.textContent = 'Show';
            }
        }

        function applyStyle(cmd) {
            document.getElementById('message_editor').focus();
            document.execCommand(cmd, false, null);
        }

        function cleanPastedHtml(dirty) {
            var out = dirty;
            out = out.replace(/<!--[\\s\\S]*?-->/g, '');
            out = out.replace(/<style[\\s\\S]*?<\\/style>/gi, '');
            out = out.replace(/<script[\\s\\S]*?<\\/script>/gi, '');
            out = out.replace(/<o:p[^>]*>[\\s\\S]*?<\\/o:p>/gi, '');
            out = out.replace(/ class="[^"]*"/gi, '');
            out = out.replace(/ style="[^"]*"/gi, '');
            out = out.replace(/<xml>[\\s\\S]*?<\\/xml>/gi, '');
            return out;
        }

        function setupPasteHandler() {
            var editor = document.getElementById('message_editor');
            editor.addEventListener('paste', function (e) {
                e.preventDefault();
                var htmlData = (e.clipboardData || window.clipboardData).getData('text/html');
                var textData = (e.clipboardData || window.clipboardData).getData('text/plain');
                if (htmlData) {
                    document.execCommand('insertHTML', false, cleanPastedHtml(htmlData));
                } else {
                    document.execCommand('insertText', false, textData);
                }
            });
        }

        function syncMessageBeforeSubmit(btn) {
            document.getElementById('message_html_hidden').value = document.getElementById('message_editor').innerHTML;
            if (btn.value === 'send') {
                return confirm('Are you sure you want to send actual Teams messages?');
            }
            return true;
        }

        window.addEventListener('DOMContentLoaded', setupPasteHandler);
    </script>
</head>
<body>
<div class="page">

    <div class="app-header">
        <div class="logo">T</div>
        <div>
            <h1>Teams Message Sender</h1>
            <p>Send formatted group messages, mentions and App/CTASK tables to Microsoft Teams chats</p>
        </div>
    </div>

    <div class="box">
        <div class="box-title"><span class="step-badge">0</span><h2>Save Graph Token</h2></div>
        <div class="box-desc">Paste a valid Microsoft Graph bearer token. It is kept in memory only for this session and is masked below.</div>
        <form method="post" action="/set-token">
            <div class="token-row">
                <input type="password" id="token_input" name="token" placeholder="Paste fresh token here" autocomplete="off">
                <button type="button" class="icon-btn" id="token_toggle_btn" onclick="toggleTokenVisibility()">Show</button>
            </div>
            <button type="submit">Save Token</button>
        </form>
        <form method="post" action="/clear-token" style="display:inline;">
            <button type="submit" class="secondary">Clear Token</button>
        </form>
        {% if token_set %}
            <div class="status-pill status-ok"><span class="dot"></span> Token saved &middot; ending ****{{ token_tail }}</div>
        {% else %}
            <div class="status-pill status-warn"><span class="dot"></span> No token saved yet</div>
        {% endif %}
    </div>

    <div class="box">
        <div class="box-title"><span class="step-badge">1</span><h2>Upload Excel and Update Chat IDs</h2></div>
        <div class="box-desc">
            Required columns: <b>App ID, App Name, Group Name, CTASK ID, CTASK Group, Chat ID, POC Name</b>.
            If <b>POC Name</b> contains an email/UPN, it will be resolved to a real Teams mention. Optional column: <b>POC Email</b>.
        </div>
        <form method="post" enctype="multipart/form-data" action="/update-chat-ids">
            <label class="field-label">Select Excel file (.xlsx)</label>
            <input type="file" name="excel" accept=".xlsx" required>
            <button type="submit">Update Chat IDs in Excel</button>
        </form>
    </div>

    <div class="box">
        <div class="box-title"><span class="step-badge">2</span><h2>Write Message</h2></div>
        <div class="box-desc">
            Use the toolbar for <b>Bold</b>/<i>Italic</i>/<u>Underline</u>. You can also paste a table directly from Excel or Word and it will keep its table formatting.<br>
            Select <b>Mention POC from Excel</b> to auto-convert or append the POC mention. One-off mentions can be typed as <b>@{Display Name|email@company.com}</b>.
        </div>
        <form method="post" action="/send">
            <label class="field-label">Message to send</label>
            <div class="editor-toolbar">
                <button type="button" onclick="applyStyle('bold')" title="Bold"><b>B</b></button>
                <button type="button" class="italic" onclick="applyStyle('italic')" title="Italic">I</button>
                <button type="button" class="underline" onclick="applyStyle('underline')" title="Underline">U</button>
                <span class="sep"></span>
                <button type="button" onclick="applyStyle('insertUnorderedList')" title="Bullet list">&#8226; List</button>
                <button type="button" onclick="applyStyle('insertOrderedList')" title="Numbered list">1. List</button>
                <span class="sep"></span>
                <button type="button" onclick="applyStyle('removeFormat')" title="Clear formatting">Clear</button>
            </div>
            <div id="message_editor" contenteditable="true">{{ form_state.message_html|safe }}</div>
            <textarea id="message_html_hidden" name="message_html" style="display:none;"></textarea>

            <div class="checkbox-row"><input type="checkbox" name="everyone" {% if form_state.everyone %}checked{% endif %}> Convert @Everyone to actual Teams mention</div>
            <div class="checkbox-row"><input type="checkbox" name="mention_poc" {% if form_state.mention_poc %}checked{% endif %}> Mention POC from Excel</div>
            <div class="checkbox-row"><input type="checkbox" name="include_table" {% if form_state.include_table %}checked{% endif %}> Include App / CTASK table below message</div>

            <label class="field-label">Delay between each group message</label>
            <select name="delay_seconds">
                <option value="0" {% if form_state.delay_seconds == "0" %}selected{% endif %}>0 Seconds</option>
                <option value="3" {% if form_state.delay_seconds == "3" %}selected{% endif %}>3 Seconds</option>
                <option value="5" {% if form_state.delay_seconds == "5" %}selected{% endif %}>5 Seconds</option>
                <option value="10" {% if form_state.delay_seconds == "10" %}selected{% endif %}>10 Seconds</option>
                <option value="15" {% if form_state.delay_seconds == "15" %}selected{% endif %}>15 Seconds</option>
                <option value="30" {% if form_state.delay_seconds == "30" %}selected{% endif %}>30 Seconds</option>
                <option value="60" {% if form_state.delay_seconds == "60" %}selected{% endif %}>60 Seconds</option>
            </select>

            <button type="submit" name="mode" value="dry_run" class="secondary" onclick="return syncMessageBeforeSubmit(this)">Dry Run Only</button>
            <button type="submit" name="mode" value="send" onclick="return syncMessageBeforeSubmit(this)">Send Actual Messages</button>
        </form>
    </div>

    {% if previews %}
    <div class="box">
        <div class="box-title"><h2>Teams Preview</h2></div>
        {% for p in previews %}
            <div class="chat-tag">Chat ID: {{ p.chat_id }}</div>
            <div class="preview">{{ p.html|safe }}</div>
        {% endfor %}
    </div>
    {% endif %}

    {% if message %}
    <div class="box">
        <div class="box-title"><h2>Result / Log</h2></div>
        {% if download_excel or download_discovered %}
        <div class="links">
            {% if download_excel %}<a href="/download/excel">Download Updated Excel</a>{% endif %}
            {% if download_discovered %}<a href="/download/discovered">Download Discovered Chats</a>{% endif %}
        </div>
        {% endif %}
        <pre class="log">{{ message }}</pre>
    </div>
    {% endif %}

</div>
</body>
</html>
"""


def render_page(message=None, download_excel=False, download_discovered=False, previews=None):
    return render_template_string(
        HTML,
        message=message,
        download_excel=download_excel,
        download_discovered=download_discovered,
        previews=previews or [],
        token_set=bool(GRAPH_TOKEN),
        token_tail=GRAPH_TOKEN[-4:] if GRAPH_TOKEN else "",
        form_state=LAST_FORM_STATE,
    )


def find_column(df, candidates):
    lower = {c.lower(): c for c in df.columns}
    for candidate in candidates:
        if candidate.lower() in lower:
            return lower[candidate.lower()]
    return None


def sanitize_editor_html(raw):
    """Light server-side cleanup of HTML coming from the contenteditable editor.
    Removes script/style tags and inline event handlers, keeps formatting tags
    (b/i/u/table/etc.) and pasted tables intact."""
    if not raw:
        return ""
    cleaned = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", raw, flags=re.IGNORECASE | re.DOTALL)
    cleaned = re.sub(r"<!--.*?-->", "", cleaned, flags=re.DOTALL)
    cleaned = re.sub(r'\son\w+="[^"]*"', "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\son\w+='[^']*'", "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


def strip_tags_for_check(html_str):
    return re.sub(r"<[^>]+>", "", html_str or "").strip()


def normalise(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    for ch in ["_", "-", "/", "\\", ",", ".", "(", ")", "[", "]", "&"]:
        text = text.replace(ch, " ")
    return " ".join(text.split())


def graph_get_all(url, token):
    headers = {"Authorization": f"Bearer {token}"}
    results = []
    while url:
        r = requests.get(url, headers=headers, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph GET failed {r.status_code}: {r.text}")
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_my_group_chats(token):
    rows = []
    for chat in graph_get_all(f"{GRAPH}/me/chats?$top=50", token):
        if chat.get("chatType") == "group":
            rows.append({
                "Teams Group Name": chat.get("topic") or "",
                "Chat ID": chat.get("id") or "",
                "Chat Type": chat.get("chatType") or "",
                "Created DateTime": chat.get("createdDateTime") or "",
            })
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["Teams Group Name", "Chat ID", "Chat Type", "Created DateTime"])
    df["Match Key"] = df["Teams Group Name"].apply(normalise)
    return df


def read_all_valid_sheets(excel_file):
    workbook = pd.read_excel(excel_file, sheet_name=None, engine="openpyxl")
    valid = {}
    for sheet, df in workbook.items():
        df.columns = df.columns.astype(str).str.strip()
        if all(c in df.columns for c in REQUIRED_COLUMNS):
            valid[sheet] = df
    if not valid:
        raise ValueError(f"No valid sheet found. Required columns: {REQUIRED_COLUMNS}")
    return workbook, valid


def choose_group_column(df):
    for c in PREFERRED_GROUP_COLUMNS:
        if c in df.columns:
            return c
    raise ValueError("No group name column found. Add Group Name or Teams Group Name.")


def find_chat_id(group_name, chats_df):
    key = normalise(group_name)
    if not key:
        return "", "Blank group name"
    exact = chats_df[chats_df["Match Key"] == key]
    if len(exact) == 1:
        return exact.iloc[0]["Chat ID"], "Exact match"
    candidates = chats_df[chats_df["Match Key"].apply(lambda x: bool(x) and (key in x or x in key))]
    if len(candidates) == 1:
        return candidates.iloc[0]["Chat ID"], f"Partial match: {candidates.iloc[0]['Teams Group Name']}"
    if len(candidates) > 1:
        return "", "Multiple matches - fill manually"
    return "", "No match found"


def update_chat_ids(excel_file, token):
    workbook, valid = read_all_valid_sheets(excel_file)
    chats_df = get_my_group_chats(token)
    discovered = excel_file.replace(".xlsx", "_discovered_chats.xlsx")
    chats_df.drop(columns=["Match Key"], errors="ignore").to_excel(discovered, index=False, engine="openpyxl")

    total_existing = total_updated = total_missing = 0
    for sheet, df in valid.items():
        group_col = choose_group_column(df)
        for col in ["Chat ID", "Chat ID Match Status"]:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].astype("object").where(pd.notna(df[col]), "").astype(str).replace({"nan": "", "NaN": ""})
        for idx, row in df.iterrows():
            existing = str(row.get("Chat ID", "")).strip()
            if existing:
                total_existing += 1
                df.loc[idx, "Chat ID Match Status"] = "Already available"
                continue
            chat_id, status = find_chat_id(row.get(group_col, ""), chats_df)
            df.loc[idx, "Chat ID"] = chat_id
            df.loc[idx, "Chat ID Match Status"] = status
            if chat_id:
                total_updated += 1
            else:
                total_missing += 1
        workbook[sheet] = df

    backup = excel_file.replace(".xlsx", f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(excel_file, backup)
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for sheet, df in workbook.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return discovered, f"Backup created: {backup}\nSame Excel updated: {excel_file}\nAlready had Chat ID: {total_existing}\nNew Chat IDs updated: {total_updated}\nRows still missing Chat ID: {total_missing}"


def collect_rows(excel_file):
    _, valid = read_all_valid_sheets(excel_file)
    frames = []
    for sheet, df in valid.items():
        if "Chat ID" not in df.columns:
            continue
        df = df.copy()
        df["Chat ID"] = df["Chat ID"].astype("object").where(pd.notna(df["Chat ID"]), "").astype(str).str.strip()
        df = df.dropna(subset=["App ID", "App Name", "CTASK ID", "CTASK Group"])
        df = df[df["Chat ID"].ne("")]
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def resolve_user(identifier, token):
    if not identifier:
        return None, "No email/UPN provided"
    identifier = str(identifier).strip()
    key = identifier.lower()
    if key in USER_CACHE:
        return USER_CACHE[key]

    headers = {"Authorization": f"Bearer {token}"}
    select = "id,displayName,mail,userPrincipalName"
    try:
        r = requests.get(f"{GRAPH}/users/{quote(identifier, safe='')}", headers=headers, params={"$select": select}, timeout=30)
        if r.status_code == 200:
            data = r.json()
            result = ({"id": data.get("id"), "displayName": data.get("displayName") or identifier, "upn": data.get("userPrincipalName") or identifier}, None)
            USER_CACHE[key] = result
            return result
        first = f"{r.status_code}: {r.text[:180]}"
        safe_id = identifier.replace("'", "''")
        r2 = requests.get(
            f"{GRAPH}/users",
            headers=headers,
            params={"$filter": f"mail eq '{safe_id}' or userPrincipalName eq '{safe_id}'", "$select": select},
            timeout=30,
        )
        if r2.status_code == 200 and r2.json().get("value"):
            data = r2.json()["value"][0]
            result = ({"id": data.get("id"), "displayName": data.get("displayName") or identifier, "upn": data.get("userPrincipalName") or identifier}, None)
            USER_CACHE[key] = result
            return result
        reason = f"Cannot resolve '{identifier}'. Direct lookup: {first}. Search: {r2.status_code}: {r2.text[:180]}"
    except requests.RequestException as e:
        reason = f"Graph/network error while resolving '{identifier}': {e}"
    result = (None, reason)
    USER_CACHE[key] = result
    return result


def split_multi_values(value):
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return []
    parts = re.split(r"[\n;,/]+", text)
    return [p.strip() for p in parts if p.strip()]


def get_group_pocs(rows_full, token):
    name_col = find_column(rows_full, POC_NAME_COLUMNS)
    email_col = find_column(rows_full, POC_EMAIL_COLUMNS)
    if not name_col and not email_col:
        return []

    pocs = []
    seen = set()
    for _, row in rows_full.iterrows():
        name_values = split_multi_values(row.get(name_col, "")) if name_col else []
        email_values = split_multi_values(row.get(email_col, "")) if email_col else []
        entries = []

        # POC Name can be email or display name.
        for val in name_values:
            if EMAIL_PATTERN.fullmatch(val):
                entries.append({"raw": val, "email": val, "name": ""})
            else:
                entries.append({"raw": val, "email": "", "name": val})

        # If separate POC Email column exists, pair it with POC Name values.
        if email_values:
            if entries:
                for i, entry in enumerate(entries):
                    if i < len(email_values) and not entry.get("email"):
                        entry["email"] = email_values[i]
            else:
                for email in email_values:
                    entries.append({"raw": email, "email": email, "name": ""})

        for entry in entries:
            email = entry.get("email", "").strip()
            raw = entry.get("raw", "").strip()
            name = entry.get("name", "").strip()
            display = name
            if email:
                resolved, _ = resolve_user(email, token)
                if resolved:
                    display = resolved.get("displayName") or name or email
            if not display:
                display = raw or email
            key = (display.lower(), email.lower())
            if key in seen:
                continue
            seen.add(key)
            pocs.append({"display": display, "email": email, "raw": raw})
    return pocs


def add_mention(mention_map, mtype, text, email=""):
    mid = len(mention_map)
    mention_map.append({"id": mid, "type": mtype, "text": text, "email": email})
    return mid


def replace_once(source, target, replacement):
    idx = source.find(target)
    if idx == -1:
        return source, False
    return source[:idx] + replacement + source[idx + len(target):], True


def build_table(rows):
    table_rows = ""
    for _, row in rows.iterrows():
        table_rows += (
            "<tr>"
            f"<td>{html.escape(str(row['App ID']))}</td>"
            f"<td>{html.escape(str(row['App Name']))}</td>"
            f"<td>{html.escape(str(row['CTASK ID']))}</td>"
            f"<td>{html.escape(str(row['CTASK Group']))}</td>"
            "</tr>"
        )
    return (
        "<br><br><table border='1' style='border-collapse:collapse'>"
        "<tr><th>App ID</th><th>App Name</th><th>CTASK ID</th><th>Assignment Group</th></tr>"
        + table_rows + "</table>"
    )


def build_message_html(message_html, rows, pocs, token, everyone_checked, poc_checked, include_table):
    mention_map = []
    # message_html already comes as HTML (bold/italic/underline/table markup preserved)
    # from the rich-text editor, so it is used as-is instead of escaping plain text.
    safe = message_html

    # @Everyone behavior requested:
    # - If checkbox selected: convert @Everyone into actual mention.
    # - If not selected: leave @Everyone exactly as plain text.
    if everyone_checked and "@Everyone" in safe:
        mid = add_mention(mention_map, "everyone", "Everyone")
        safe = safe.replace("@Everyone", f'<at id="{mid}">Everyone</at>', 1)

    # Inline one-off syntax: @{Display Name|email}
    def inline_repl(match):
        display = match.group(1).strip()
        email = match.group(2).strip()
        mid = add_mention(mention_map, "user", display, email)
        return f'<at id="{mid}">{html.escape(display)}</at>'
    safe = INLINE_MENTION_PATTERN.sub(inline_repl, safe)

    if poc_checked:
        already_emails = {m["email"].lower() for m in mention_map if m.get("email")}
        to_append = []

        for poc in pocs:
            email = poc.get("email", "")
            display = poc.get("display", "")
            raw = poc.get("raw", "")
            if not email or email.lower() in already_emails:
                continue

            # First try to convert POC if existing in message.
            replaced_ok = False
            for candidate in [raw, email, display]:
                if not candidate:
                    continue
                escaped = html.escape(candidate)
                if escaped in safe:
                    mid = add_mention(mention_map, "user", display, email)
                    safe, replaced_ok = replace_once(safe, escaped, f'<at id="{mid}">{html.escape(display)}</at>')
                    if replaced_ok:
                        already_emails.add(email.lower())
                    break

            # If POC is not in message, append at bottom automatically.
            if not replaced_ok and email.lower() not in already_emails:
                mid = add_mention(mention_map, "user", display, email)
                to_append.append(f'<at id="{mid}">{html.escape(display)}</at>')
                already_emails.add(email.lower())

        if to_append:
            safe += "<br><br>POC: " + ", ".join(to_append)

    if include_table:
        safe += build_table(rows)
    return safe, mention_map


def build_mentions_payload(chat_id, mention_map, token):
    payload_mentions = []
    diagnostics = []
    for mention in mention_map:
        if mention["type"] == "everyone":
            payload_mentions.append({
                "id": mention["id"],
                "mentionText": "Everyone",
                "mentioned": {"conversation": {"id": chat_id, "displayName": "Everyone", "conversationIdentityType": "chat"}}
            })
        else:
            resolved, reason = resolve_user(mention["email"], token)
            if resolved and resolved.get("id"):
                payload_mentions.append({
                    "id": mention["id"],
                    "mentionText": mention["text"],
                    "mentioned": {"user": {"id": resolved["id"], "displayName": resolved["displayName"], "userIdentityType": "aadUser"}}
                })
                diagnostics.append(f"POC '{mention['text']}' resolved OK and will notify as Teams mention.")
            else:
                diagnostics.append(f"POC '{mention['text']}' could not resolve, so it will not notify. Reason: {reason}")
    return payload_mentions, diagnostics


def preview_html(message_html, mention_map, token):
    preview = message_html
    for mention in mention_map:
        tag = f'<at id="{mention["id"]}">{html.escape(mention["text"])}</at>'
        if mention["type"] == "everyone":
            repl = '<span class="mention">@Everyone</span>'
        else:
            resolved, _ = resolve_user(mention["email"], token)
            repl = f'<span class="mention">@{html.escape(mention["text"])}</span>' if resolved and resolved.get("id") else html.escape(mention["text"])
        preview = preview.replace(tag, repl)
    return preview


def check_mentions(mention_map, token):
    output = []
    for mention in mention_map:
        if mention["type"] == "everyone":
            output.append("@Everyone mention will be attempted because checkbox is selected.")
        else:
            resolved, reason = resolve_user(mention["email"], token)
            if resolved and resolved.get("id"):
                output.append(f"POC '{mention['text']}' ({mention['email']}) WOULD resolve to real Teams mention.")
            else:
                output.append(f"POC '{mention['text']}' ({mention['email']}) would NOT resolve. Reason: {reason}")
    return output


def post_message(chat_id, message_html, mention_map, token):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    mentions, diagnostics = build_mentions_payload(chat_id, mention_map, token)
    payload = {"body": {"contentType": "html", "content": message_html}}
    if mentions:
        payload["mentions"] = mentions

    r = requests.post(f"{GRAPH}/chats/{chat_id}/messages", headers=headers, json=payload, timeout=60)
    if r.status_code >= 400 and mentions:
        diagnostics.append(f"Mention payload rejected ({r.status_code}: {r.text[:200]}). Retrying without mentions.")
        plain_html = re.sub(r'<at id="\d+">(.*?)</at>', r'\1', message_html)
        r = requests.post(f"{GRAPH}/chats/{chat_id}/messages", headers=headers, json={"body": {"contentType": "html", "content": plain_html}}, timeout=60)
    if r.status_code not in [200, 201, 202]:
        raise RuntimeError(f"{r.status_code}: {r.text}")
    return diagnostics


@app.route("/")
def home():
    return render_page()


@app.route("/set-token", methods=["POST"])
def set_token_route():
    global GRAPH_TOKEN
    token = request.form.get("token", "").strip()
    if not token:
        return render_page(message="ERROR:\nToken box was empty.")
    GRAPH_TOKEN = token
    USER_CACHE.clear()
    return render_page(message="Token saved. Now upload Excel and update Chat IDs.")


@app.route("/clear-token", methods=["POST"])
def clear_token_route():
    global GRAPH_TOKEN
    GRAPH_TOKEN = None
    USER_CACHE.clear()
    return render_page(message="Token cleared.")


@app.route("/update-chat-ids", methods=["POST"])
def update_route():
    global LAST_EXCEL_FILE, LAST_DISCOVERED_FILE
    try:
        if not GRAPH_TOKEN:
            raise ValueError("No token set. Save token first.")
        file = request.files["excel"]
        excel_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(excel_path)
        LAST_EXCEL_FILE = excel_path
        discovered, msg = update_chat_ids(excel_path, GRAPH_TOKEN)
        LAST_DISCOVERED_FILE = discovered
        msg += "\n\nNext: write message, choose POC/table options, then Dry Run."
        return render_page(message=msg, download_excel=True, download_discovered=True)
    except Exception as e:
        return render_page(message=f"ERROR:\n{e}")


@app.route("/send", methods=["POST"])
def send_route():
    global LAST_FORM_STATE
    try:
        if not GRAPH_TOKEN:
            raise ValueError("No token set. Save token first.")
        if not LAST_EXCEL_FILE or not os.path.exists(LAST_EXCEL_FILE):
            raise ValueError("Upload Excel and update Chat IDs first.")

        message_html = sanitize_editor_html(request.form.get("message_html", "").strip())
        if not strip_tags_for_check(message_html):
            raise ValueError("Message box is empty.")

        include_table = "include_table" in request.form
        everyone_checked = "everyone" in request.form
        poc_checked = "mention_poc" in request.form
        dry_run = request.form.get("mode") == "dry_run"
        try:
            delay_seconds = int(request.form.get("delay_seconds", "5"))
        except ValueError:
            delay_seconds = 5
        delay_seconds = max(delay_seconds, 0)

        # Remember exactly what was submitted so the page re-renders with the
        # SAME content after Dry Run, instead of resetting to the old default text.
        LAST_FORM_STATE = {
            "message_html": message_html,
            "everyone": everyone_checked,
            "mention_poc": poc_checked,
            "include_table": include_table,
            "delay_seconds": str(delay_seconds),
        }

        df = collect_rows(LAST_EXCEL_FILE)
        if df.empty:
            return render_page(message="No rows with Chat ID found. Check Chat ID column.", download_excel=True, download_discovered=True)

        previews = []
        output = [
            f"Mode: {'Dry Run' if dry_run else 'Send Actual Messages'}",
            f"@Everyone mention checkbox: {'Selected - convert to actual mention' if everyone_checked else 'Not selected - leave @Everyone as plain text'}",
            f"Mention POC from Excel: {'Yes' if poc_checked else 'No'}",
            f"Include table: {'Yes' if include_table else 'No'}",
            f"Delay: {delay_seconds} seconds",
        ]
        success = failed = 0
        groups = list(df.groupby("Chat ID"))
        total = len(groups)

        for i, (chat_id, group_rows_full) in enumerate(groups, start=1):
            table_rows = group_rows_full[["App ID", "App Name", "CTASK ID", "CTASK Group"]].drop_duplicates()
            pocs = get_group_pocs(group_rows_full, GRAPH_TOKEN)
            msg_html, mention_map = build_message_html(
                message_html=message_html,
                rows=table_rows,
                pocs=pocs,
                token=GRAPH_TOKEN,
                everyone_checked=everyone_checked,
                poc_checked=poc_checked,
                include_table=include_table,
            )
            previews.append({"chat_id": chat_id, "html": preview_html(msg_html, mention_map, GRAPH_TOKEN)})

            output.append(f"\nGroup {i} of {total}")
            output.append(f"Chat ID: {chat_id}")
            output.append("POC(s): " + (", ".join([p["display"] for p in pocs]) if pocs else "None"))
            if include_table:
                output.append(table_rows.to_string(index=False))
            output.append("Mention check:")
            checks = check_mentions(mention_map, GRAPH_TOKEN)
            if checks:
                for line in checks:
                    output.append(" - " + line)
            else:
                output.append(" - No actual mention will be attempted.")

            if dry_run:
                output.append("DRY RUN ONLY - not sent")
            else:
                try:
                    for line in post_message(chat_id, msg_html, mention_map, GRAPH_TOKEN):
                        output.append(" - " + line)
                    output.append("SUCCESS: Message sent")
                    success += 1
                except Exception as e:
                    output.append(f"FAILED: {e}")
                    failed += 1
                if i < total and delay_seconds > 0:
                    output.append(f"Waiting {delay_seconds} seconds before next group...")
                    time.sleep(delay_seconds)

        output.append(f"\nCompleted. Success: {success}, Failed: {failed}")
        return render_page(message="\n".join(output), download_excel=True, download_discovered=True, previews=previews)
    except Exception as e:
        return render_page(message=f"ERROR:\n{e}", download_excel=True, download_discovered=True)


@app.route("/download/excel")
def download_excel():
    return send_file(LAST_EXCEL_FILE, as_attachment=True)


@app.route("/download/discovered")
def download_discovered():
    return send_file(LAST_DISCOVERED_FILE, as_attachment=True)


if __name__ == "__main__":
    print("Starting UI... open http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
