
# Auto_Reminder_List.py — Templates-enabled full version
# 2025-10-13
# - 3 templates in config (T1/T2/T3), user-editable in Settings
# - Per-mail template selection via dropdown (Treeview cell editor)
# - Uses selected template on send, falls back to remind_message or T1
# - Keeps prior features (cancel key, reply detection, icons, tray, etc.)

import os, re, sys, json, time, uuid, argparse, urllib.parse, threading, ctypes
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pythoncom
import win32com.client as win32
import win32event, win32api, winerror
import winreg

import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image
from pystray import Icon as icon, MenuItem as item

# ---------------- Paths / Files ----------------
APPDATA_DIR = os.path.join(os.environ.get("APPDATA", os.getcwd()), "AutoRemindCS")
os.makedirs(APPDATA_DIR, exist_ok=True)

STATE_FILE  = os.path.join(APPDATA_DIR, "state.json")
CONFIG_FILE = os.path.join(APPDATA_DIR, "config.json")

# ---------------- Outlook constants ----------------
OL_MAILITEM = 43
OL_FOLDER_SENT = 5
OL_DEFAULT_ITEM_MAIL = 0
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_DELETED = 3

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "AutoRemindCS"

# ---------------- Icon helper ----------------
def _res_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

_APP_ICO = _res_path("icon.ico")
_APP_PNG = _res_path("icon.png")
_APP_ICONIMG = None

def set_window_icon(win: tk.Misc):
    global _APP_ICONIMG
    try:
        if os.path.exists(_APP_ICO):
            try: win.iconbitmap(_APP_ICO)
            except Exception: pass
        if os.path.exists(_APP_PNG):
            if _APP_ICONIMG is None:
                _APP_ICONIMG = tk.PhotoImage(file=_APP_PNG)
            win.wm_iconphoto(True, _APP_ICONIMG)
    except Exception:
        pass

# ---------------- Logging ----------------
VERBOSE = False
def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if level == "DEBUG" and not VERBOSE:
        return
    line = f"{ts} [{level}] {msg}"
    print(line)
    try:
        week_str = datetime.now().strftime("%Y-W%U")
        log_file = os.path.join(APPDATA_DIR, f"app_{week_str}.log")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def now_naive():
    return datetime.now()

def to_local_naive(dt):
    if dt is None or not isinstance(dt, datetime): return None
    if dt.tzinfo is None: return dt
    try:
        return dt.astimezone(ZoneInfo("Asia/Seoul")).replace(tzinfo=None)
    except Exception:
        return dt.replace(tzinfo=None)

def _pretty_ts(ts: str) -> str:
    if not ts or ts == "-":
        return "-"
    try:
        return datetime.fromisoformat(ts.replace("Z","")).strftime("%Y-%m-%d %H:%M")
    except Exception:
        s = ts.replace("T", " ")
        if "." in s: s = s.split(".", 1)[0]
        return s[:16]

# ---------------- Config / Templates ----------------
DEFAULT_TEMPLATES = [
    {"code": "T1", "label": "Korean",  "text": "안녕하세요,\n이전 메일 관련하여 회신 부탁드립니다.\n확인 후 답장 부탁드립니다. 감사합니다."},
    {"code": "T2", "label": "English", "text": "Hello,\nKind reminder regarding my previous email.\nPlease check and reply when you have a moment. Thanks."},
    {"code": "T3", "label": "Short",   "text": "Reminder: Please reply to my previous email. Thank you."},
]

def _ensure_templates_in_config(cfg: dict) -> dict:
    t = cfg.get("templates")
    if not isinstance(t, list) or len(t) != 3:
        cfg["templates"] = DEFAULT_TEMPLATES.copy()
    else:
        for i, slot in enumerate(t):
            if not isinstance(slot, dict):
                t[i] = DEFAULT_TEMPLATES[i]
                continue
            slot.setdefault("code",  DEFAULT_TEMPLATES[i]["code"])
            slot.setdefault("label", DEFAULT_TEMPLATES[i]["label"])
            slot.setdefault("text",  DEFAULT_TEMPLATES[i]["text"])
    return cfg

def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg = _ensure_templates_in_config(cfg)
    except FileNotFoundError:
        cfg = {
            "remind_message": "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다.",
            "auto_start": False,
            "verbose": False,
            "templates": DEFAULT_TEMPLATES.copy(),
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    # backfill
    cfg.setdefault("remind_message", "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다.")
    cfg.setdefault("auto_start", False)
    cfg.setdefault("verbose", False)
    return cfg

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ---------------- State ----------------
def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception:
        return {}

def save_state(st):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_FILE)

def make_state_key(entry_id, recipient_addr):
    return f"{entry_id}|{(recipient_addr or '').lower()}"

# ---------------- Registry (Auto-start) ----------------
def _exe_path():
    try:
        return sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
    except Exception:
        return os.path.abspath(sys.argv[0])

def register_startup_reg(app_name=APP_RUN_NAME, extra_args=""):
    path = _exe_path()
    cmd = f'"{path}" {extra_args}'.strip()
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, app_name, 0, winreg.REG_SZ, cmd)
        log(f"[INFO] Startup 등록(HKCU\\Run): {cmd}")
    except Exception as e:
        log(f"[ERR] Startup 등록 실패(HKCU\\Run): {e}", level="ERR")

def unregister_startup_reg(app_name=APP_RUN_NAME):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, app_name)
        log("[INFO] Startup 해제(HKCU\\Run) 완료")
    except FileNotFoundError:
        log("[INFO] Startup 해제(HKCU\\Run): 기존 등록 없음")
    except Exception as e:
        log(f"[ERR] Startup 해제 실패(HKCU\\Run): {e}", level="ERR")

# ---------------- Outlook helpers (subset needed here) ----------------
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_TRANSPORT_HEADERS   = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_ATTACH_CONTENT_ID   = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"

BAD_CID_DENYLIST = ("filelist.html","filelist.xml","themedata.thmx","colorschememapping.xml","editdata.mso")

PREFIXES = ["re:", "fw:", "fwd:", "답장:", "회신:", "전달:", "참조:", "回覆:", "転送:"]

def format_body_text(text):
    if not text: return ""
    processed = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    processed = processed.replace('\n', '<br>')
    return f"<p>{processed}</p>"

def canonicalize_subject(subj: str) -> str:
    if not subj: return ""
    s = subj.strip()
    changed = True
    while changed:
        changed = False
        ss = s.lstrip()
        for p in PREFIXES:
            if ss.lower().startswith(p):
                s = ss[len(p):].lstrip(); changed = True; break
    s = re.sub(r"^\s*\[remind\]\s*","", s, flags=re.I)
    s = re.sub(r"\s*\[(?:D[ANH]|F[UIP])\s*\d+\s*(?:MIN|H|D|W|M)\]\s*"," ", s, flags=re.I)
    return re.sub(r"\s+"," ", s).strip().lower()

def parse_yard_tag(subject):
    if not subject: return None, None
    s = subject.upper().replace("［","[").replace("］","]")
    m = re.search(r"\[(SHI|HMD|HHI|HSHI|HO|HJSC)(\d+)(MIN|H|D|W|M)\]", s)
    if not m: return None, None
    yard, num, unit = m.group(1), int(m.group(2)), m.group(3)
    if unit == "MIN": interval_days = num/1440.0
    elif unit == "H": interval_days = num/24.0
    elif unit == "D": interval_days = float(num)
    elif unit == "W": interval_days = float(num)*7.0
    elif unit == "M": interval_days = float(num)*30.0
    else: interval_days = None
    return yard, interval_days

def get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32.GetActiveObject("Outlook.Application")
    except Exception:
        pass
    for proc in os.popen('tasklist').read().splitlines():
        if "OUTLOOK.EXE" in proc.upper():
            break
    else:
        try: os.startfile("outlook.exe")
        except Exception: pass
    t0 = time.time()
    while time.time()-t0 < 30:
        try: return win32.Dispatch("Outlook.Application")
        except Exception: time.sleep(2)
    raise RuntimeError("Outlook COM attach failed")

def my_addresses(ns):
    addrs = set()
    try:
        me = ns.CurrentUser
        ae = me.AddressEntry
        if ae and ae.Type=="EX":
            ex = ae.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                addrs.add(ex.PrimarySmtpAddress.lower())
        elif ae:
            addr = (ae.Address or "").lower()
            if addr: addrs.add(addr)
    except Exception:
        pass
    return addrs

def is_from_me(m, me_set):
    try:
        addr = (getattr(m,"SenderEmailAddress","") or "").lower()
        return addr in me_set
    except Exception:
        return False

def conv_key(mail):
    try:
        msgid = mail.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    except Exception:
        msgid = None
    if msgid:
        return f"MSGID:{msgid}"
    try:
        if getattr(mail,"EntryID",None):
            return f"EID:{mail.EntryID}"
    except Exception:
        pass
    cid = getattr(mail,"ConversationID",None)
    if cid: return f"CID:{cid}"
    topic = getattr(mail,"ConversationTopic","") or ""
    sent_on = to_local_naive(getattr(mail,"SentOn",None))
    sent_key = sent_on.strftime("%Y-%m-%d %H:%M:%S") if sent_on else "NA"
    return f"TOPIC:{topic}|SENT:{sent_key}"

def _walk_folders(folder):
    yield folder
    for i in range(1, folder.Folders.Count+1):
        sub = folder.Folders.Item(i)
        for f in _walk_folders(sub):
            yield f

def _get_deleted_roots(ns):
    roots=[]
    try:
        for store in ns.Stores:
            try:
                di = store.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
            except Exception:
                di = None
            if di:
                try: roots.append(di.FolderPath)
                except Exception: pass
    except Exception:
        pass
    return roots

def _is_under_deleted(folder, deleted_roots):
    try:
        fpath = folder.FolderPath
        if not fpath: return False
        for root in deleted_roots:
            if root and fpath.startswith(root):
                return True
    except Exception:
        return False
    return False

def _all_mail_folders(ns, include_deleted=False):
    deleted_roots = [] if include_deleted else _get_deleted_roots(ns)
    for store in ns.Stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        for f in _walk_folders(root):
            try:
                if (not include_deleted) and _is_under_deleted(f, deleted_roots):
                    continue
                if f.DefaultItemType == OL_DEFAULT_ITEM_MAIL:
                    yield f
            except Exception:
                continue

def check_and_update_replies(app, orig_mail, state, verbose=False):
    ns = app.GetNamespace("MAPI")
    me_set = my_addresses(ns)
    orig_subject = orig_mail.Subject or ""
    orig_sent = to_local_naive(getattr(orig_mail, "SentOn", None))
    recipients = [(r.Address, r.Type) for r in orig_mail.Recipients if r.Type in (1, 3)]

    for addr, rtype in recipients:
        state_key = make_state_key(orig_mail.EntryID, addr)
        cancelled_keys = set(state.get("__cancelled_keys__", []))
        if state_key in cancelled_keys:
            if verbose: log(f"[CANCELLED-SKIP] {state_key} cancelled; skip")
            continue
        if state.get(state_key, {}).get("reply_received", False):
            continue

        for folder in _all_mail_folders(ns, include_deleted=True):
            try:
                items = folder.Items
                items.Sort("ReceivedTime", True)
            except Exception:
                continue

            for m in items:
                try:
                    if m.Class != OL_MAILITEM: continue
                    if is_from_me(m, me_set): continue
                    rt = to_local_naive(getattr(m, "ReceivedTime", None))
                    if not rt or rt <= orig_sent: continue

                    can = canonicalize_subject(getattr(m, "Subject", "") or "")
                    base = canonicalize_subject(orig_subject)
                    if not base: continue
                    ok = (can == base) if len(base)<8 else ((can==base) or (base in can) or (can in base))
                    if not ok: continue

                    sender_addr = (m.SenderEmailAddress or "").lower()
                    if addr.lower() in sender_addr:
                        if verbose:
                            log(f"[REPLY*:{rtype}] {folder.FolderPath} / {rt:%Y-%m-%d %H:%M} / {m.SenderName} / {m.Subject}")
                        state[state_key] = {
                            "reply_received": True,
                            "last_sent": state.get(state_key, {}).get("last_sent"),
                            "detected_at": rt.isoformat()
                        }
                        break
                except Exception:
                    continue

# ---------------- HTML helpers (signature images) ----------------
def _guess_mime_from_ext(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".png"]: return "image/png"
    if ext in [".jpg",".jpeg"]: return "image/jpeg"
    if ext in [".gif"]: return "image/gif"
    if ext in [".bmp"]: return "image/bmp"
    if ext in [".svg"]: return "image/svg+xml"
    return "application/octet-stream"

def _mark_attachment_inline(att, cid: str, path: str, verbose=False):
    pa = att.PropertyAccessor
    mime = _guess_mime_from_ext(path)
    try: pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
    except Exception as e:
        if verbose: log(f"[INLINE-MARK ERR] CID {e}")
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001E", mime)
    except Exception: pass
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3713001E", f"cid:{cid}")
    except Exception: pass
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)
    except Exception: pass
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003", 1)
    except Exception: pass
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370B0003", -1)
    except Exception: pass

IMG_SRC_REGEXES = [
    r'<img\b[^>]*?\bsrc=["\']([^"\']+)["\']',
    r'<v:imagedata\b[^>]*?\bsrc=["\']([^"\']+)["\']',
    r'background(?:-image)?:\s*url\((["\']?)([^)\s"\']+)\1\)',
]

def _unquote_file_url(src):
    if isinstance(src,str) and src.lower().startswith("file:///"):
        p = urllib.parse.urlparse(src)
        path = urllib.parse.unquote(p.path)
        if re.match(r"^/[A-Za-z]:/", path): path = path[1:]
        return path.replace("/","\\")
    return src

def _resolve_signature_path(src):
    src = _unquote_file_url(src)
    if isinstance(src,str) and os.path.isabs(src) and os.path.exists(src): return src
    base = os.path.expanduser(r"~\AppData\Roaming\Microsoft\Signatures")
    cand = os.path.join(base, src) if isinstance(src,str) else None
    if cand and os.path.exists(cand): return cand
    if isinstance(src,str) and os.path.isdir(base):
        name = os.path.basename(src)
        for root,_,files in os.walk(base):
            if name in files: return os.path.join(root,name)
    return None

def _attach_images_and_rewrite_html(mail, html):
    if not html: return html
    replaced = html
    used = {}
    for pattern in IMG_SRC_REGEXES:
        if 'background' in pattern:
            for m in list(re.finditer(pattern, replaced, flags=re.I)):
                src = m.group(2)
                if not src or src.lower().startswith("cid:") or re.match(r'^(https?:)?//', src, flags=re.I): continue
                norm = src.strip()
                cid = used.get(norm)
                if not cid:
                    path = _resolve_signature_path(norm)
                    if not path: continue
                    att = mail.Attachments.Add(path)
                    cid = f"{uuid.uuid4().hex}@sig"
                    _mark_attachment_inline(att, cid, path, verbose=True)
                    used[norm]=cid
                replaced = re.sub(r'(background(?:-image)?:\s*url\((["\']?))'+re.escape(src)+r'(\2\))', r'\1cid:'+cid+r'\3', replaced, flags=re.I)
        elif 'imagedata' in pattern:
            for m in list(re.finditer(pattern, replaced, flags=re.I)):
                src = m.group(1)
                if not src or src.lower().startswith("cid:") or re.match(r'^(https?:)?//', src, flags=re.I): continue
                norm = src.strip()
                cid = used.get(norm)
                if not cid:
                    path = _resolve_signature_path(norm)
                    if not path: continue
                    att = mail.Attachments.Add(path)
                    cid = f"{uuid.uuid4().hex}@sig"
                    _mark_attachment_inline(att, cid, path, verbose=True)
                    used[norm]=cid
                replaced = re.sub(r'(<v:imagedata\b[^>]*?\bsrc=["\'])'+re.escape(src)+r'(["\'])', r'\1cid:'+cid+r'\2', replaced, flags=re.I)
        else:
            for m in list(re.finditer(pattern, replaced, flags=re.I)):
                src = m.group(1)
                if not src or src.lower().startswith("cid:") or re.match(r'^(https?:)?//', src, flags=re.I): continue
                norm = src.strip()
                cid = used.get(norm)
                if not cid:
                    path = _resolve_signature_path(norm)
                    if not path: continue
                    att = mail.Attachments.Add(path)
                    cid = f"{uuid.uuid4().hex}@sig"
                    _mark_attachment_inline(att, cid, path, verbose=True)
                    used[norm]=cid
                replaced = re.sub(r'(<img\b[^>]*?\bsrc=["\'])'+re.escape(src)+r'(["\'])', r'\1cid:'+cid+r'\2', replaced, flags=re.I)
    return replaced

def _sanitize_bad_cids(html: str, verbose=False) -> str:
    if not isinstance(html, str) or not html: return html
    html = re.sub(r'<img\b[^>]*?\bsrc=["\']cid:([^"\']+)["\'][^>]*>',
                  lambda m: "" if any(x in m.group(1).lower() for x in BAD_CID_DENYLIST) else m.group(0),
                  html, flags=re.I)
    html = re.sub(r'<v:imagedata\b[^>]*?\bsrc=["\']cid:([^"\']+)["\'][^>]*>',
                  lambda m: "" if any(x in m.group(1).lower() for x in BAD_CID_DENYLIST) else m.group(0),
                  html, flags=re.I)
    def _css_sub(m):
        cid = m.group(2).lower()
        return "" if any(x in cid for x in BAD_CID_DENYLIST) else m.group(1)
    html = re.sub(r'(background(?:-image)?\s*:\s*url\(\s*["\']?cid:([^)\'"\s]+)["\']?\s*\)\s*;?)', _css_sub, html, flags=re.I)
    for bad in BAD_CID_DENYLIST:
        html = re.sub(rf'cid:{re.escape(bad)}', "", html, flags=re.I)
    return html

# ---------------- Send remind ----------------
def send_remind_for_recipients(app, item, subject, fallback_body, yard_code, state, dry_run=False, verbose=False):
    def _self_smtp():
        try:
            ae = app.Session.CurrentUser.AddressEntry
            exu = ae.GetExchangeUser()
            return (exu.PrimarySmtpAddress if exu else ae.Address) or None
        except Exception:
            return None

    try:
        recipients = []
        for r in item.Recipients:
            try:
                if r.Type in (1, 3):
                    addr = getattr(r, "Address", None) or getattr(r, "Name", None)
                    if addr: recipients.append((addr, r.Type))
            except Exception:
                continue
        if not recipients:
            if verbose: log("[WARN] no To/BCC recipients on original mail")
            return False

        cfg = load_config()
        tpl_map = {s["code"]: s for s in cfg["templates"]}
        default_text = cfg.get("remind_message") or tpl_map["T1"]["text"]

        me_addr = _self_smtp() or getattr(item, "SenderEmailAddress", None) or "me@example.com"
        sent_any = False

        st_snapshot = load_state()
        cancelled_keys = set(st_snapshot.get("__cancelled_keys__", []))

        for addr, rtype in recipients:
            state_key = make_state_key(item.EntryID, addr)
            if state_key in cancelled_keys:
                log(f"[CANCELLED-SKIP] {state_key} is cancelled; skip sending.")
                continue
            if state.get(state_key, {}).get("reply_received", False):
                continue

            # choose template per key
            code = state.get(state_key, {}).get("template_code")
            if code and code in tpl_map:
                remind_text = tpl_map[code]["text"]
            else:
                remind_text = default_text

            remind_html = format_body_text(remind_text)

            fwd = item.Forward()
            fwd.Subject = f"[Remind] {subject}"
            fwd.BodyFormat = 2
            fwd.HTMLBody = (
                "<div style='font-family:Malgun Gothic,Segoe UI,Arial,sans-serif; font-size:10pt;'>"
                f"{remind_html}</div><br>" + fwd.HTMLBody
            )
            fwd.HTMLBody = _sanitize_bad_cids(_attach_images_and_rewrite_html(fwd, fwd.HTMLBody))

            if rtype == 1:  # To
                fwd.To = addr
            else:           # BCC
                if not fwd.To:
                    fwd.To = me_addr
                recip = fwd.Recipients.Add(addr)
                recip.Type = 3

            try: fwd.Recipients.ResolveAll()
            except Exception: pass

            if dry_run:
                log(f"[DRY-RUN] Would send | {fwd.Subject} ({yard_code}) | To={fwd.To}, CC={fwd.CC}, BCC={fwd.BCC}")
                continue

            try:
                fwd.Save()
                fwd.Send()
                sent_any = True
                ts = now_naive().isoformat()

                # keep selected template meta for display
                label = None
                if code and code in tpl_map:
                    label = tpl_map[code]["label"]

                state[state_key] = {
                    "reply_received": False,
                    "last_sent": ts,
                    "subject": subject,
                    "template_code": code,
                    "template_label": label,
                }
                save_state(state)

                log(f"[SENT] To={fwd.To}, CC={fwd.CC}, BCC={fwd.BCC} | {fwd.Subject} ({yard_code})")
                log(f"[STATE-UPD] {state_key} reply_received=False last_sent={ts}")
            except Exception as e:
                log(f"[ERR-SEND] {e}")

        return True if (dry_run or sent_any) else False
    except Exception as e:
        log(f"[ERR-SEND] {e}")
        return False

# ---------------- Scan loop (subset) ----------------
def _has_newer_outgoing_with_same_subject(ns, canon_subj: str, sent_on: datetime, include_deleted=False, verbose=False):
    try:
        stores = list(ns.Stores)
    except Exception:
        return False
    deleted_roots = [] if not include_deleted else []

    for store in stores:
        try:
            sent = store.GetDefaultFolder(OL_FOLDER_SENT)
        except Exception:
            continue
        for f in _walk_folders(sent):
            try:
                items = f.Items
                items.Sort("[SentOn]", True)
                items = items.Restrict("[SentOn] >= '" + (sent_on.strftime('%m/%d/%Y %I:%M %p')) + "'")
            except Exception:
                continue
            try:
                enum = iter(items)
            except Exception:
                def enum_iter(it):
                    for i in range(1, it.Count+1): yield it.Item(i)
                enum = enum_iter(items)
            for it in enum:
                try:
                    if it.Class != OL_MAILITEM: continue
                    s = canonicalize_subject(getattr(it, "Subject", "") or "")
                    if s != canon_subj: continue
                    so = to_local_naive(getattr(it, "SentOn", None))
                    if so and so > sent_on:
                        subj = getattr(it, "Subject", "") or ""
                        if re.search(r"^\s*\[remind\]\s*", subj, flags=re.I): continue
                        if verbose: log(f"[SKIP-NEWER-OUT] newer outgoing found {so:%Y-%m-%d %H:%M}")
                        return True
                except Exception:
                    continue
    return False

def cycle_once(ns, app, state, lookback_days, dry_run, force_send, skip_reply_check, verbose,
               include_self, due_from_last, reply_mode, include_deleted, precheck_epsilon_sec, loop_budget_sec, max_age_hours, skip_if_newer_outgoing):
    sent = ns.GetDefaultFolder(OL_FOLDER_SENT)
    items = sent.Items; items.Sort("SentOn", True)
    cutoff = now_naive() - timedelta(days=lookback_days)
    found=0; sent_count=0

    loop_started = time.time()
    if verbose: log("[LOOP-START] budget timer reset")

    for mail in items:
        try:
            if mail.Class!=OL_MAILITEM: continue
            subject = (mail.Subject or "")
            if subject.lstrip().upper().startswith("[REMIND]"):
                if verbose: log("[SKIP] reminder mail itself")
                continue
            code, interval_days = parse_yard_tag(subject)
            if not code: continue

            sent_on = to_local_naive(getattr(mail,"SentOn",None))
            if not sent_on or sent_on < cutoff: continue
            found += 1

            if verbose:
                now_ts = now_naive()
                log(f"[CHK] subj='{subject}' code={code} tag={interval_days}d sent={sent_on:%Y-%m-%d %H:%M}")
                try:
                    delta_min = (now_ts - sent_on).total_seconds()/60.0
                    log(f"[TIME] now={now_ts:%Y-%m-%d %H:%M} | sent_on={sent_on:%Y-%m-%d %H:%M} | Δ={delta_min:.1f}min")
                except Exception as _e:
                    log(f"[TIME-ERR] {_e}")

            key = conv_key(mail)
            rec = state.get(key, {})
            last_sent_iso = rec.get("last_remind_at")

            base_time = sent_on
            if due_from_last and last_sent_iso:
                try:
                    last_dt_base = to_local_naive(datetime.fromisoformat(last_sent_iso))
                    if last_dt_base and last_dt_base > base_time:
                        base_time = last_dt_base
                except Exception:
                    pass
            due_time = base_time + timedelta(days=interval_days)
            now_ts = now_naive()
            due_ok  = now_ts >= due_time

            if max_age_hours and not force_send:
                age_h = (now_ts - sent_on).total_seconds() / 3600.0
                if age_h > max_age_hours:
                    if verbose: log(f"[SKIP-STALE] tag too old: {age_h:.1f}h > {max_age_hours}h"); continue

            if verbose:
                log(f"[DUE] base={'last_remind_at' if (due_from_last and last_sent_iso and base_time!=sent_on) else 'sent_on'} "
                    f"| base_time={base_time:%Y-%m-%d %H:%M} | due_time={due_time:%Y-%m-%d %H:%M} | due_ok={due_ok}")

            if (not force_send) and (not due_ok):
                remaining = (due_time - now_ts).total_seconds()
                if remaining > precheck_epsilon_sec:
                    if verbose: log(f"[PRECHECK-SKIP] due in {remaining:.1f}s (> {precheck_epsilon_sec}s)")
                    if (time.time() - loop_started) > loop_budget_sec:
                        log(f"[LOOP-BUDGET] elapsed={time.time()-loop_started:.1f}s > {loop_budget_sec}s, defer"); break
                    continue

            if last_sent_iso and not force_send:
                try:
                    last_dt = to_local_naive(datetime.fromisoformat(last_sent_iso))
                    if last_dt and now_ts - last_dt < timedelta(days=interval_days):
                        if verbose: log("[SKIP] within interval since last remind")
                        if (time.time() - loop_started) > loop_budget_sec:
                            log(f"[LOOP-BUDGET] elapsed={time.time()-loop_started:.1f}s > {loop_budget_sec}s, defer"); break
                        continue
                except Exception:
                    pass

            if skip_if_newer_outgoing:
                canon = canonicalize_subject(subject or "")
                if _has_newer_outgoing_with_same_subject(ns, canon, sent_on, include_deleted=include_deleted, verbose=verbose):
                    if verbose: log("[SKIP] newer outgoing exists in same thread")
                    if (time.time() - loop_started) > loop_budget_sec:
                        log(f"[LOOP-BUDGET] elapsed={time.time()-loop_started:.1f}s > {loop_budget_sec}s, defer"); break
                    continue

            if not skip_reply_check:
                try:
                    if verbose:
                        log(f"[DEBUG-REPLYCHK] subj='{subject}' conv_id={mail.ConversationID} topic='{mail.ConversationTopic}' check_after={sent_on:%Y-%m-%d %H:%M}")
                    check_and_update_replies(app, mail, state, verbose=verbose)
                    save_state(state)
                except Exception as e:
                    log(f"[ERR-REPLYCHK] {e}")

            if (not force_send) and (not due_ok):
                if verbose: log("[SKIP] not yet due")
                if (time.time() - loop_started) > loop_budget_sec:
                    log(f"[LOOP-BUDGET] elapsed={time.time()-loop_started:.1f}s > {loop_budget_sec}s, defer"); break
                continue

            if dry_run:
                log(f"[DRY-RUN] Would send | {subject} ({code})")
            else:
                ok = send_remind_for_recipients(
                    app, mail, subject,
                    load_config().get("remind_message",""),
                    code, state, dry_run=dry_run, verbose=verbose
                )
                if ok:
                    sent_count += 1
                    state[key] = state.get(key, {})
                    state[key]["last_remind_at"] = now_ts.isoformat()
                    save_state(state)
                else:
                    log("[WARN] send failed; state not updated")
        except Exception as e:
            log(f"[ERR] {e}")

    if verbose:
        log(f"[INFO] Candidates processed: {found}, sent: {sent_count}")

# ---------------- Tray-bound UI (Settings / List) ----------------
root = tk.Tk()
set_window_icon(root)     # set icon BEFORE withdraw for taskbar icon
root.withdraw()

def create_and_show_gui():
    top = tk.Toplevel(root)
    set_window_icon(top)
    top.title("설정 - 템플릿/기본문구 (Outlook 스타일)")
    top.geometry("1280x880")  # 전체 창 크기 확대
    top.grid_columnconfigure(1, weight=1)
    top.grid_rowconfigure(1, weight=1)

    cfg = load_config()

    # 기본 리마인드 메시지
    tk.Label(top, text="기본 리마인드 메시지(템플릿 미지정 시 사용):", font=("Malgun Gothic", 10, "bold")).grid(
        row=0, column=0, sticky="nw", padx=10, pady=(10,4)
    )
    txt = tk.Text(top, height=8, wrap=tk.WORD, font=("Malgun Gothic", 10))
    txt.grid(row=0, column=1, sticky="nsew", padx=10, pady=(10,4))
    txt.insert("1.0", cfg.get("remind_message",""))

    # 템플릿 3개
    lf = tk.LabelFrame(top, text="Templates (3개)", font=("Malgun Gothic", 10, "bold"))
    lf.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=8)
    lf.grid_columnconfigure(1, weight=1)

    slots = cfg["templates"]
    tpl_vars = []

    for i, slot in enumerate(slots):
        r0 = 2 + i * 2
        tk.Label(lf, text=f"템플릿 {i+1} 라벨:", font=("Malgun Gothic", 9)).grid(
            row=r0, column=0, sticky="w", padx=8, pady=4
        )
        lbl_var = tk.StringVar(value=slot.get("label", ""))
        tk.Entry(lf, textvariable=lbl_var, width=32, font=("Malgun Gothic", 9)).grid(
            row=r0, column=1, sticky="w", padx=8
        )

        tk.Label(lf, text=f"코드:", font=("Malgun Gothic", 9)).grid(row=r0, column=2, sticky="e", padx=8)
        tk.Label(lf, text=slot.get("code", f"T{i+1}"), width=8, fg="#555", font=("Malgun Gothic", 9)).grid(
            row=r0, column=3, sticky="w"
        )

        r1 = r0 + 1
        # Outlook 스타일의 큰 편집 칸
        t = tk.Text(lf, height=10, wrap=tk.WORD, font=("Malgun Gothic", 10))
        t.grid(row=r1, column=0, columnspan=4, sticky="nsew", padx=8, pady=(0,12))
        t.insert("1.0", slot.get("text", ""))

        # 스크롤바 추가
        sb = tk.Scrollbar(lf, command=t.yview)
        sb.grid(row=r1, column=4, sticky="ns", padx=(0,8))
        t.config(yscrollcommand=sb.set)

        tpl_vars.append((lbl_var, t, slot.get("code", f"T{i+1}")))

    # 옵션
    auto_start_var = tk.BooleanVar(value=cfg.get("auto_start", False))
    verbose_var = tk.BooleanVar(value=cfg.get("verbose", False))
    tk.Checkbutton(top, text="Windows 시작 시 자동 실행", variable=auto_start_var, font=("Malgun Gothic", 9)).grid(
        row=2, column=0, columnspan=2, sticky="w", padx=10
    )
    tk.Checkbutton(top, text="DEBUG 로그 출력 (Verbose)", variable=verbose_var, font=("Malgun Gothic", 9)).grid(
        row=3, column=0, columnspan=2, sticky="w", padx=10
    )

    # 저장 버튼
    def do_save():
        cfg2 = load_config()
        cfg2["remind_message"] = txt.get("1.0", "end-1c")
        cfg2["auto_start"] = auto_start_var.get()
        cfg2["verbose"] = verbose_var.get()

        new_slots = []
        for lbl_var, tw, code in tpl_vars:
            new_slots.append({
                "code": code,
                "label": (lbl_var.get() or code).strip(),
                "text": tw.get("1.0", "end-1c"),
            })
        cfg2["templates"] = new_slots
        save_config(cfg2)

        global VERBOSE
        VERBOSE = cfg2["verbose"]
        try:
            if cfg2["auto_start"]:
                register_startup_reg(APP_RUN_NAME)
            else:
                unregister_startup_reg(APP_RUN_NAME)
        except Exception as e:
            log(f"[WARN] Startup 토글 오류: {e}", level="WARN")

        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        top.destroy()

    fbtn = tk.Frame(top)
    fbtn.grid(row=4, column=1, sticky="e", padx=10, pady=10)
    tk.Button(fbtn, text="저장", command=do_save, font=("Malgun Gothic", 10, "bold")).pack(side="left", padx=6)
    tk.Button(fbtn, text="닫기", command=top.destroy, font=("Malgun Gothic", 10)).pack(side="left")

    # 저장 버튼
    def do_save():
        cfg2 = load_config()
        cfg2["remind_message"] = txt.get("1.0", "end-1c")
        cfg2["auto_start"] = auto_start_var.get()
        cfg2["verbose"] = verbose_var.get()

        new_slots = []
        for lbl_var, tw, code in tpl_vars:
            new_slots.append({
                "code": code,
                "label": (lbl_var.get() or code).strip(),
                "text": tw.get("1.0", "end-1c"),
            })
        cfg2["templates"] = new_slots
        save_config(cfg2)

        global VERBOSE
        VERBOSE = cfg2["verbose"]
        try:
            if cfg2["auto_start"]:
                register_startup_reg(APP_RUN_NAME)
            else:
                unregister_startup_reg(APP_RUN_NAME)
        except Exception as e:
            log(f"[WARN] Startup 토글 오류: {e}", level="WARN")

        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        top.destroy()

    fbtn = tk.Frame(top)
    fbtn.grid(row=4, column=1, sticky="e", padx=10, pady=10)
    tk.Button(fbtn, text="저장", command=do_save).pack(side="left", padx=6)
    tk.Button(fbtn, text="닫기", command=top.destroy).pack(side="left")

def open_settings_window(icon=None, item=None):
    root.after(0, create_and_show_gui)

def open_remind_list_window(icon=None, item=None):
    def _show():
        top = tk.Toplevel(root)
        set_window_icon(top)
        top.title("리마인드 리스트")
        top.geometry("1460x600")

        # 레이아웃: 좌측 테이블, 우측 컨트롤 패널
        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(0, weight=1)

        # ----- 좌측: 트리뷰 -----
        tree = ttk.Treeview(
            top,
            columns=("recipient","subject","template","last_sent"),
            show="headings", selectmode="extended"
        )
        tree.heading("recipient", text="Recipient")
        tree.heading("subject",   text="Subject")
        tree.heading("template",  text="Template")
        tree.heading("last_sent", text="Last Sent")
        tree.column("recipient", width=150, anchor="w")
        tree.column("subject",   width=380, anchor="w")
        tree.column("template",  width=70, anchor="center")
        tree.column("last_sent", width=80, anchor="center")
        tree.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(10,6), pady=10)

        ysb = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        ysb.grid(row=0, column=0, rowspan=2, sticky="nse", padx=(0,10), pady=10)
        tree.configure(yscrollcommand=ysb.set)

        # ----- 우측: 컨트롤 패널 -----
        panel = tk.Frame(top)
        panel.grid(row=0, column=1, sticky="nsew", padx=(0,10), pady=10)
        panel.grid_columnconfigure(0, weight=1)

        # 상태 표시
        status_var = tk.StringVar(value="변경 없음")
        def set_dirty(dirty: bool):
            status_var.set("미저장 변경 있음" if dirty else "변경 없음")

        dirty = {"flag": False}
        def mark_dirty():
            dirty["flag"] = True
            set_dirty(True)

        tk.Label(panel, textvariable=status_var, fg="#0a84ff").grid(row=0, column=0, sticky="w", pady=(0,8))

        # 템플릿 선택 콤보 + 버튼들
        cfg = load_config()
        tpl_labels = [s["label"] for s in cfg["templates"]]
        tpl_by_label = {s["label"]: s for s in cfg["templates"]}

        tk.Label(panel, text="템플릿 선택:").grid(row=1, column=0, sticky="w")
        tpl_choice = tk.StringVar(value=tpl_labels[0] if tpl_labels else "")
        cb = ttk.Combobox(panel, textvariable=tpl_choice, values=tpl_labels, state="readonly")
        cb.grid(row=2, column=0, sticky="ew", pady=(2,8))

        def apply_to_selected():
            label = tpl_choice.get()
            if label not in tpl_by_label:
                messagebox.showwarning("알림", "유효한 템플릿을 선택하세요.")
                return
            for rowid in tree.selection():
                # 트리뷰 표시만 갱신(임시 변경)
                tree.set(rowid, "template", label)
                # pending_changes 에 기록
                pending_changes[rowid] = label
            if tree.selection():
                mark_dirty()

        tk.Button(panel, text="선택 행에 적용", command=apply_to_selected).grid(row=3, column=0, sticky="ew", pady=(0,8))

        def delete_selected():
            sel = tree.selection()
            if not sel:
                return
            if not messagebox.askyesno("확인", "선택 항목을 삭제(발송 취소) 하시겠습니까?"):
                return
            st = load_state()
            cancelled = set(st.get("__cancelled_keys__", []))
            for key in sel:
                st.pop(key, None)
                cancelled.add(key)
                tree.delete(key)
                # pending에 있었으면 제거
                pending_changes.pop(key, None)
            st["__cancelled_keys__"] = sorted(cancelled)
            save_state(st)
            # 테이블 삭제는 즉시 저장되었으므로 dirty 판단 재계산
            set_dirty(bool(pending_changes))

        tk.Button(panel, text="선택 삭제(발송 취소)", command=delete_selected).grid(row=4, column=0, sticky="ew", pady=(0,8))

        def save_changes():
            if not pending_changes:
                messagebox.showinfo("저장", "저장할 변경이 없습니다.")
                return
            st = load_state()
            # rowid == state key 형태: "<entry_id>|<addr>"
            for rowid, label in pending_changes.items():
                code = None
                for s in cfg["templates"]:
                    if s["label"] == label:
                        code = s["code"]; break
                if rowid in st:
                    st[rowid]["template_code"]  = code
                    st[rowid]["template_label"] = label
            save_state(st)
            pending_changes.clear()
            dirty["flag"] = False
            set_dirty(False)
            messagebox.showinfo("저장 완료", "변경 사항이 저장되었습니다.")

        tk.Button(panel, text="변경 저장", command=save_changes).grid(row=5, column=0, sticky="ew", pady=(0,8))

        def refresh_table():
            populate()  # 현재 state 기준으로 재로딩
            pending_changes.clear()
            dirty["flag"] = False
            set_dirty(False)

        tk.Button(panel, text="새로고침", command=refresh_table).grid(row=6, column=0, sticky="ew")

        # 닫기
        def close_window():
            if dirty["flag"]:
                if not messagebox.askyesno("확인", "미저장 변경이 있습니다. 저장하지 않고 닫을까요?"):
                    return
            top.destroy()
        tk.Button(panel, text="닫기", command=close_window).grid(row=7, column=0, sticky="ew", pady=(12,0))

        # ----- 데이터 로딩/편집 -----
        pending_changes = {}  # rowid -> label

        def populate():
            tree.delete(*tree.get_children())
            st = load_state()
            for key, val in st.items():
                if "|" not in key: 
                    continue
                if val.get("reply_received", False):
                    continue
                addr = key.split("|",1)[1]
                subj = val.get("subject","-")
                last = _pretty_ts(val.get("last_sent","-"))
                label = val.get("template_label") or "-"
                tree.insert("", "end", iid=key, values=(addr, subj, label, last))

        # 인라인 더블클릭 편집(즉시 저장 대신 pending으로)
        editor = None
        def edit_template_cell(event=None):
            nonlocal editor
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            col = tree.identify_column(event.x)
            if col != "#3":  # template column
                return
            rowid = tree.identify_row(event.y)
            if not rowid:
                return
            bbox = tree.bbox(rowid, col)
            if not bbox:
                return
            x,y,w,h = bbox
            if editor:
                editor.destroy()

            # 현재 config 기준 라벨 목록
            cfg_local = load_config()
            labels_local = [s["label"] for s in cfg_local["templates"]]
            var = tk.StringVar(value=tree.set(rowid, "template"))
            editor = ttk.Combobox(tree, textvariable=var, values=labels_local, state="readonly")
            editor.place(x=x, y=y, width=max(w, 140), height=h)

            def apply_choice(*_):
                label = var.get()
                tree.set(rowid, "template", label)
                pending_changes[rowid] = label
                mark_dirty()
                if editor:
                    editor.destroy()

            def cancel_edit(*_):
                if editor:
                    editor.destroy()

            editor.bind("<<ComboboxSelected>>", apply_choice)
            editor.bind("<Return>", apply_choice)
            editor.bind("<Escape>", cancel_edit)
            editor.focus_set()

        tree.bind("<Double-1>", edit_template_cell)

        populate()
    root.after(0, _show)


# ---------------- Worker / Tray / Main ----------------
exit_event = threading.Event()

def check_single_instance(mutex_name="AutoRemindCS_Mutex"):
    _ = win32event.CreateMutex(None, False, mutex_name)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        ctypes.windll.user32.MessageBoxW(0, "이미 Auto_Reminder가 실행 중입니다.", "Auto Reminder", 0x40)
        sys.exit(0)

def show_startup_notification():
    ctypes.windll.user32.MessageBoxW(0, "백그라운드에서 Auto Reminder가 실행 중입니다.", "Auto Reminder 실행됨", 0x40)

def start_mail_check_loop(args):
    st = load_state()
    while not exit_event.is_set():
        try:
            st = load_state()
            log("[INFO] Starting new scan cycle.")
            app = get_outlook()
            ns = app.GetNamespace("MAPI")
            cycle_once(ns, app, st, args.lookback_days, args.dry_run, args.force_send,
                       args.skip_reply_check, args.verbose, args.include_self, args.due_from_last,
                       args.reply_mode, args.include_deleted, args.precheck_epsilon_sec,
                       args.loop_budget_sec, args.max_age_hours, args.skip_if_newer_outgoing)
        except Exception as e:
            log(f"[ERROR] An error occurred in the mail check loop: {e}")
        log(f"[INFO] Cycle finished. Waiting for {args.interval_min} minute(s).")
        exit_event.wait(args.interval_min * 60)

def exit_action(ic=None, it=None):
    log("[INFO] Exit requested. Shutting down.")
    exit_event.set()
    try:
        ic.stop()
    except Exception:
        pass
    try:
        root.after(0, root.quit)
    except Exception:
        pass

def main():
    check_single_instance()
    cfg = load_config()

    parser = argparse.ArgumentParser(description="Auto Reminder (Templates)")
    parser.add_argument("--precheck-epsilon-sec", type=int, default=10)
    parser.add_argument("--loop-budget-sec", type=int, default=45)
    parser.add_argument("--max-age-hours", type=float, default=0.0)
    parser.add_argument("--skip-if-newer-outgoing", action="store_true")
    parser.add_argument("--lookback-days", type=int, default=60)
    parser.add_argument("--interval-min", type=int, default=1)
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--force-send", action="store_true")
    parser.add_argument("--skip-reply-check", action="store_true")
    parser.add_argument("--include-self", action="store_true")
    parser.add_argument("--due-from-last", action="store_true")
    parser.add_argument("--include-deleted", action="store_true")
    parser.add_argument("--reply-mode", choices=["hdr-only","hdr-first","conv-first"], default="conv-first")
    args = parser.parse_args()

    global VERBOSE
    VERBOSE = args.verbose or cfg.get("verbose", False)
    log(f"[INFO] Verbose mode = {VERBOSE}")

    # startup toggle
    try:
        if cfg.get("auto_start", False):
            register_startup_reg(APP_RUN_NAME)
        else:
            unregister_startup_reg(APP_RUN_NAME)
    except Exception as e:
        log(f"[WARN] Startup 초기화 오류: {e}", level="WARN")

    # Background worker
    mail_thread = threading.Thread(target=start_mail_check_loop, args=(args,), daemon=True)
    mail_thread.start()
    log("[INFO] Mail check background thread started.")

    # Tray
    def resource_path(filename):
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, filename)
        return os.path.join(os.path.abspath("."), filename)
    try:
        img = Image.open(resource_path("icon.png"))
    except FileNotFoundError:
        log("[WARN] 'icon.png' not found. Using default blue square.")
        img = Image.new('RGB', (64,64), color='blue')

    menu = (item('설정 열기', open_settings_window),
            item('리마인드 리스트', open_remind_list_window),
            item('종료', exit_action))
    tray = icon('AutoMailSystem', img, "자동 메일 리마인더", menu)

    show_startup_notification()
    tray.run_detached()
    root.mainloop()

if __name__ == "__main__":
    main()
