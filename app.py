import json
import os
import re
import subprocess
import tempfile
import threading
import time
import unicodedata
import urllib.parse
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, jsonify, render_template, request
from rapidfuzz import fuzz

try:
    import pythoncom
    import win32com.client
except Exception:
    pythoncom = None
    win32com = None

try:
    import pyautogui
except Exception:
    pyautogui = None

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
POWERBI_HELPER = BASE_DIR / "powerbi_local.ps1"

DEFAULT_CONFIG = {
    "assistant_name": "WorkMate AI",
    "user_name": "User",
    "mail_hours_back": 16,
    "max_emails": 40,
    "priority_keywords": [
        "urgent", "deadline", "important", "priority",
        "error", "issue", "delay", "today"
    ],
    "teams": {"auto_send": False, "send_delay_seconds": 3.0},
    "powerbi_desktop": {"report_name": "YourReport"}
}

app = Flask(__name__)
state = {
    "last_mail_refresh": None,
    "last_mail_summary": None,
    "alerts": [],
    "assistant_log": [],
    "powerbi": {"connected": False, "database": None, "metadata": None}
}


def save_config(value):
    CONFIG_PATH.write_text(
        json.dumps(value, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )


def load_config():
    if not CONFIG_PATH.exists():
        save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))

    saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    value = json.loads(json.dumps(DEFAULT_CONFIG))
    value.update(saved)
    value["teams"].update(saved.get("teams", {}))
    value["powerbi_desktop"].update(saved.get("powerbi_desktop", {}))
    return value


config = load_config()


def clean_text(value, limit=260):
    return re.sub(r"\s+", " ", str(value or "")).strip()[:limit]


def add_alert(title, message, level="info"):
    state["alerts"].insert(0, {
        "title": title,
        "message": message,
        "level": level,
        "time": datetime.now().strftime("%H:%M")
    })
    state["alerts"] = state["alerts"][:20]


def outlook():
    if win32com is None:
        raise RuntimeError("pywin32 is not available.")
    pythoncom.CoInitialize()
    application = win32com.client.Dispatch("Outlook.Application")
    return application, application.GetNamespace("MAPI")


def recent_emails(hours=None):
    hours = int(hours or config["mail_hours_back"])
    _, namespace = outlook()
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    cutoff = datetime.now() - timedelta(hours=hours)
    result = []

    for item in items:
        if len(result) >= int(config["max_emails"]):
            break
        try:
            received = item.ReceivedTime.replace(tzinfo=None)
            if received < cutoff:
                break

            subject = clean_text(item.Subject or "(no subject)", 180)
            body = clean_text(item.Body, 320)
            joined = f"{subject} {body}".lower()
            keywords = [
                word for word in config["priority_keywords"]
                if word.lower() in joined
            ]
            unread = bool(item.UnRead)
            importance = int(item.Importance)
            score = (2 if unread else 0) + (3 if importance == 2 else 0)
            score += min(3, len(keywords))

            result.append({
                "subject": subject,
                "sender": clean_text(
                    item.SenderName or item.SenderEmailAddress, 100
                ),
                "received": received.strftime("%d/%m/%Y %H:%M"),
                "unread": unread,
                "priority_score": score,
                "priority_keywords": keywords,
                "snippet": body
            })
        except Exception:
            continue

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
    return result


def summarize_emails(emails):
    priority = sorted(
        [mail for mail in emails if mail["priority_score"] >= 4],
        key=lambda x: x["priority_score"],
        reverse=True
    )
    summary = {
        "total": len(emails),
        "unread": sum(1 for x in emails if x["unread"]),
        "priority": len(priority),
        "top_priority": priority[:8],
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    state["last_mail_refresh"] = datetime.now().isoformat()
    state["last_mail_summary"] = summary
    return summary


def send_email(to, subject, body, cc="", send_now=False):
    if not str(to).strip():
        raise ValueError("Recipient is required.")

    application, _ = outlook()
    mail = application.CreateItem(0)
    mail.To = to
    mail.CC = cc or ""
    mail.Subject = subject
    mail.Body = body

    if send_now:
        mail.Send()
        result = "sent"
    else:
        mail.Display()
        result = "opened as draft"

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
    return result


PROGRAMS = {
    "outlook": ["cmd", "/c", "start", "", "outlook"],
    "excel": ["cmd", "/c", "start", "", "excel"],
    "word": ["cmd", "/c", "start", "", "winword"],
    "powerpoint": ["cmd", "/c", "start", "", "powerpnt"],
    "teams": ["cmd", "/c", "start", "", "msteams:"],
    "power bi": ["cmd", "/c", "start", "", "pbidesktop"],
    "powerbi": ["cmd", "/c", "start", "", "pbidesktop"]
}


def open_program(name):
    key = str(name).strip().lower()
    if key not in PROGRAMS:
        raise ValueError(f"Unknown program: {name}")
    subprocess.Popen(PROGRAMS[key], shell=False)


def send_teams(recipient, message, auto_send=False):
    if not recipient or not message:
        raise ValueError("Recipient and message are required.")

    url = (
        "https://teams.microsoft.com/l/chat/0/0"
        f"?users={urllib.parse.quote(recipient)}"
        f"&message={urllib.parse.quote(message)}"
    )
    webbrowser.open(url)

    if auto_send:
        if pyautogui is None:
            raise RuntimeError("PyAutoGUI is not available.")

        def press_enter():
            time.sleep(float(config["teams"]["send_delay_seconds"]))
            pyautogui.press("enter")

        threading.Thread(target=press_enter, daemon=True).start()


def run_powerbi_helper(mode, dax=None):
    if not POWERBI_HELPER.exists():
        raise RuntimeError("powerbi_local.ps1 was not found.")

    command = [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(POWERBI_HELPER),
        "-Mode", mode,
        "-ReportName", config["powerbi_desktop"]["report_name"]
    ]

    query_file = None
    try:
        if dax is not None:
            handle, query_file = tempfile.mkstemp(
                prefix="workmate_", suffix=".dax"
            )
            os.close(handle)
            Path(query_file).write_text(dax, encoding="utf-8")
            command += ["-QueryFile", query_file]

        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )

        if result.returncode != 0:
            raise RuntimeError(
                (result.stderr or result.stdout or "Power BI error").strip()
            )

        output = result.stdout.strip()
        if not output:
            raise RuntimeError("Power BI returned no data.")
        return json.loads(output.splitlines()[-1])
    finally:
        if query_file:
            Path(query_file).unlink(missing_ok=True)


def normalize(value):
    value = "".join(
        c for c in unicodedata.normalize("NFD", str(value).lower())
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9%]+", " ", value)).strip()


def connect_powerbi():
    info = run_powerbi_helper("status")
    state["powerbi"]["connected"] = True
    state["powerbi"]["database"] = info.get("database")
    return info


def scan_model():
    raw = run_powerbi_helper("metadata")

    def truthy(value):
        return str(value).strip().lower() in {"true", "1", "yes"}

    tables = [
        {
            "id": f"T{i}",
            "name": str(row.get("Name", "")).strip(),
            "hidden": truthy(row.get("IsHidden", False))
        }
        for i, row in enumerate(raw.get("tables", []), 1)
        if str(row.get("Name", "")).strip()
    ]

    columns = [
        {
            "table": str(row.get("TableName", "")).strip(),
            "name": str(row.get("Name", "")).strip(),
            "data_type": str(row.get("DataType", "")).strip(),
            "hidden": truthy(row.get("IsHidden", False))
        }
        for row in raw.get("columns", [])
        if str(row.get("Name", "")).strip()
    ]

    measures = [
        {
            "table": str(row.get("TableName", "")).strip(),
            "name": str(row.get("Name", "")).strip(),
            "expression": str(row.get("Expression", "") or ""),
            "hidden": truthy(row.get("IsHidden", False))
        }
        for row in raw.get("measures", [])
        if str(row.get("Name", "")).strip()
    ]

    if not tables:
        raise RuntimeError("No Power BI tables were returned.")

    model = {"tables": tables, "columns": columns, "measures": measures}
    state["powerbi"]["connected"] = True
    state["powerbi"]["database"] = raw.get("database")
    state["powerbi"]["metadata"] = model
    return model


def metadata():
    return state["powerbi"]["metadata"] or scan_model()


def best_item(question, items, formatter, threshold=68):
    q = normalize(question)
    scored = []

    for item in items:
        target = normalize(formatter(item))
        score = max(
            fuzz.ratio(q, target),
            fuzz.partial_ratio(q, target),
            fuzz.token_set_ratio(q, target)
        )
        if target and target in q:
            score = min(100, score + 18)
        scored.append((score, item))

    scored.sort(key=lambda x: x[0], reverse=True)
    if not scored or scored[0][0] < threshold:
        return None, 0
    return scored[0][1], scored[0][0]


def dax_name(name):
    return str(name).replace("'", "''").replace("]", "]]")


def execute_dax(query):
    return run_powerbi_helper("query", query)


def answer_powerbi(question):
    model = metadata()
    q = normalize(question)

    if "quali tabelle" in q or "show tables" in q:
        rows = [{"Table": x["name"]} for x in model["tables"] if not x["hidden"]]
        return {"reply": "Available model tables.", "rows": rows}

    if "quali misure" in q or "show measures" in q:
        rows = [{"Measure": x["name"]} for x in model["measures"] if not x["hidden"]]
        return {"reply": "Available model measures.", "rows": rows}

    measure, score = best_item(
        question,
        [m for m in model["measures"] if not m["hidden"]],
        lambda x: x["name"],
        70
    )
    if measure and score >= 78:
        name = dax_name(measure["name"])
        query = f'EVALUATE ROW("Measure", "{name}", "Value", [{name}])'
        result = execute_dax(query)
        rows = result.get("rows", [])
        value = rows[0].get("Value") if rows else None
        return {
            "reply": f'{measure["name"]}: {value}',
            "rows": rows,
            "dax": query
        }

    column, _ = best_item(
        question,
        [c for c in model["columns"] if not c["hidden"]],
        lambda x: f'{x["table"]} {x["name"]}',
        68
    )
    group_terms = ["per ", "by ", "raggruppa", "group", "distribution"]
    if column and any(term in q for term in group_terms):
        table = dax_name(column["table"])
        col = dax_name(column["name"])
        query = (
            "EVALUATE\nTOPN(30,\n"
            "SUMMARIZECOLUMNS(\n"
            f"    '{table}'[{col}],\n"
            f'    "Count", COUNTROWS(\\'{table}\\')\n'
            "),\n[Count], DESC\n)"
        )
        result = execute_dax(query)
        return {
            "reply": f'Grouped by {column["table"]}[{column["name"]}].',
            "rows": result.get("rows", []),
            "dax": query
        }

    suggestions = []
    for item in model["measures"]:
        score = fuzz.token_set_ratio(q, normalize(item["name"]))
        suggestions.append((score, "measure", item["name"]))

    for item in model["columns"]:
        score = fuzz.token_set_ratio(q, normalize(item["name"]))
        suggestions.append(
            (score, "column", f'{item["table"]}[{item["name"]}]')
        )

    suggestions.sort(reverse=True)
    rows = [
        {"type": item_type, "name": name, "score": score}
        for score, item_type, name in suggestions[:8]
        if score >= 45
    ]
    return {
        "reply": "I could not map the request safely. Closest model objects are shown below.",
        "rows": rows
    }


def handle_assistant(text):
    original = str(text or "").strip()
    if not original:
        return {"reply": "Write a request."}

    normalized = normalize(original)

    if any(term in normalized for term in [
        "power bi", "powerbi", "dax", "measure", "misura",
        "kpi", "pending", "eseguire", "odl"
    ]):
        answer = answer_powerbi(original)
        return {
            "reply": answer["reply"],
            "action": "powerbi_question",
            "powerbi_answer": answer
        }

    commands = {
        "refresh_mail": ["check mail", "controlla mail", "aggiorna mail"],
        "open_excel": ["open excel", "apri excel"],
        "open_word": ["open word", "apri word"],
        "open_outlook": ["open outlook", "apri outlook"],
        "open_teams": ["open teams", "apri teams"],
        "open_powerbi": ["open power bi", "apri power bi"]
    }

    best_name, best_score = None, 0
    for name, examples in commands.items():
        for example in examples:
            score = fuzz.token_set_ratio(normalized, normalize(example))
            if score > best_score:
                best_name, best_score = name, score

    if best_score < 62:
        return {
            "reply": "I can check Outlook, open Office apps, prepare Teams messages and query Power BI."
        }

    if best_name == "refresh_mail":
        emails = recent_emails()
        summary = summarize_emails(emails)
        return {
            "reply": (
                f'Found {summary["total"]} messages, '
                f'{summary["unread"]} unread and '
                f'{summary["priority"]} priority.'
            ),
            "summary": summary
        }

    app_name = best_name.replace("open_", "").replace("powerbi", "power bi")
    open_program(app_name)
    return {"reply": f"Opened {app_name.title()}."}


@app.route("/")
def home():
    return render_template("index.html", config=config)


@app.route("/api/status")
def api_status():
    return jsonify({
        "assistant_name": config["assistant_name"],
        "user_name": config["user_name"],
        "last_mail_refresh": state["last_mail_refresh"],
        "mail_summary": state["last_mail_summary"],
        "alerts": state["alerts"],
        "assistant_log": state["assistant_log"][-20:],
        "powerbi_desktop": {
            "connected": state["powerbi"]["connected"],
            "database": state["powerbi"]["database"],
            "has_metadata": bool(state["powerbi"]["metadata"])
        }
    })


@app.route("/api/mails/refresh", methods=["POST"])
def api_mails():
    try:
        emails = recent_emails()
        return jsonify(
            ok=True,
            emails=emails,
            summary=summarize_emails(emails)
        )
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/email/send", methods=["POST"])
def api_email_send():
    try:
        data = request.get_json(force=True)
        result = send_email(
            data.get("to", ""),
            data.get("subject", ""),
            data.get("body", ""),
            data.get("cc", ""),
            bool(data.get("send_now"))
        )
        return jsonify(ok=True, message=f"Email {result}.")
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/program/open", methods=["POST"])
def api_open():
    try:
        open_program(request.get_json(force=True).get("program", ""))
        return jsonify(ok=True)
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 400


@app.route("/api/teams/send", methods=["POST"])
def api_teams():
    try:
        data = request.get_json(force=True)
        send_teams(
            data.get("recipient", ""),
            data.get("message", ""),
            bool(data.get("auto_send"))
        )
        return jsonify(ok=True, message="Teams chat opened.")
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/powerbi-desktop/connect", methods=["POST"])
def api_pbi_connect():
    try:
        info = connect_powerbi()
        add_alert(
            "Power BI connected",
            f'Connected to {info.get("database") or config["powerbi_desktop"]["report_name"]}.',
            "success"
        )
        return jsonify(ok=True, data={
            "database": info.get("database"),
            "report_name": config["powerbi_desktop"]["report_name"],
            "table_count": 0,
            "measure_count": 0
        })
    except Exception as exc:
        state["powerbi"]["connected"] = False
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/powerbi-desktop/metadata")
def api_pbi_metadata():
    try:
        model = metadata()
        return jsonify(
            ok=True,
            connected=True,
            database=state["powerbi"]["database"],
            metadata=model
        )
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/powerbi-desktop/ask", methods=["POST"])
def api_pbi_ask():
    try:
        question = request.get_json(force=True).get("question", "")
        return jsonify(ok=True, answer=answer_powerbi(question))
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/powerbi-desktop/query", methods=["POST"])
def api_pbi_query():
    try:
        query = request.get_json(force=True).get("dax", "").strip()
        if not query:
            raise ValueError("DAX query is required.")
        return jsonify(ok=True, data=execute_dax(query))
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/assistant", methods=["POST"])
def api_assistant():
    try:
        text = request.get_json(force=True).get("text", "")
        return jsonify(ok=True, **handle_assistant(text))
    except Exception as exc:
        return jsonify(ok=False, error=str(exc)), 500


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    global config

    if request.method == "GET":
        return jsonify(config)

    incoming = request.get_json(force=True)
    current = load_config()

    for key in ["assistant_name", "user_name", "mail_hours_back", "max_emails"]:
        if key in incoming:
            current[key] = incoming[key]

    current["teams"].update(incoming.get("teams", {}))
    current["powerbi_desktop"].update(
        incoming.get("powerbi_desktop", {})
    )

    save_config(current)
    config = load_config()
    return jsonify(ok=True, config=config)


if __name__ == "__main__":
    url = "http://127.0.0.1:8765"
    threading.Timer(1, lambda: webbrowser.open(url)).start()
    print(f"WorkMate AI: {url}")
    app.run(host="127.0.0.1", port=8765, debug=False, threaded=True)
