
# auto_mail_system_final_complete_fixed.py
# 2025-10-02 — Critical stability fixes for Settings GUI + Tray + Threads.
# - Single Tk root (main thread), pystray runs detached
# - Settings window opens reliably (no undefined variables / duplicate defs)
# - Args parsed once before threads; background loop started with args
# - Removed duplicate functions and duplicate Tk roots
# - Fixed open_settings_window callback (no immediate call)
# - Added missing ZoneInfo import
# - Removed premature thread start that caused TypeError
# - Exit from tray now also quits Tk mainloop cleanly

import os, re, json, time, uuid, argparse, urllib.parse, pythoncom, threading
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo  # FIX: used by to_local_naive
import sys
import winreg
import win32com.client as win32
import win32event
import win32api
import winerror
from win10toast import ToastNotifier  # (optional; not used directly)
import ctypes

# GUI and Tray
import tkinter as tk
import tkinter.ttk as ttk
from datetime import datetime
from tkinter import messagebox
from PIL import Image
from pystray import Icon as icon, MenuItem as item

# ---- Global / Base Paths ----
LAST_CLEANUP = 0
CLEANUP_INTERVAL_SEC = 300
BODY_MAP = {}

APPDATA_DIR = os.path.join(os.environ.get("APPDATA", os.getcwd()), "AutoRemindCS")
os.makedirs(APPDATA_DIR, exist_ok=True)

STATE_FILE  = os.path.join(APPDATA_DIR, "state.json")
LOG_FILE    = os.path.join(APPDATA_DIR, "remind.log")
CONFIG_FILE = os.path.join(APPDATA_DIR, "config.json")

# ---- Outlook constants
OL_MAILITEM = 43
OL_FOLDER_SENT = 5
OL_DEFAULT_ITEM_MAIL = 0
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_DELETED = 3

# ===== 시작 프로그램 등록/해제 =====
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "AutoRemindCS"  # 시작프로그램에 표시될 이름

def _res_path(name: str) -> str:
    """PyInstaller/실행폴더 모두 호환되는 리소스 경로"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

_APP_ICO = _res_path("icon.ico")      # 씨넷 ico
_APP_PNG = _res_path("icon.png")      # 씨넷 png
_APP_ICONIMG = None                   # PhotoImage 캐시(가비지컬렉션 방지)

def set_window_icon(win: tk.Misc):
    """해당 창의 타이틀 아이콘 + 작업표시줄 아이콘 지정(.ico 우선, png 보조)"""
    global _APP_ICONIMG
    try:
        # 작업표시줄/제목 표시줄 아이콘(Windows는 ico가 가장 확실)
        if os.path.exists(_APP_ICO):
            try:
                win.iconbitmap(_APP_ICO)
            except Exception:
                pass
        # 타이틀바 파비콘(플랫폼 공통) - png PhotoImage
        if os.path.exists(_APP_PNG):
            if _APP_ICONIMG is None:
                _APP_ICONIMG = tk.PhotoImage(file=_APP_PNG)
            # True = 이 창과 미래의 Toplevel에 상속
            win.wm_iconphoto(True, _APP_ICONIMG)
    except Exception:
        pass
# ==== /ICON HELPER ====

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

def is_startup_registered_reg(app_name=APP_RUN_NAME):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ) as k:
            val, _ = winreg.QueryValueEx(k, app_name)
            return True, val
    except FileNotFoundError:
        return False, ""
    except Exception:
        return False, ""

def check_single_instance(mutex_name="AutoRemindCS_Mutex"):
    _ = win32event.CreateMutex(None, False, mutex_name)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        ctypes.windll.user32.MessageBoxW(0, "이미 Auto_Reminder가 실행 중입니다.", "Auto Reminder", 0x40)
        sys.exit(0)

def show_startup_notification():
    ctypes.windll.user32.MessageBoxW(
        0,
        "백그라운드에서 Auto Reminder가 실행 중입니다.",
        "Auto Reminder 실행됨",
        0x40
    )

# ===== 설정 파일 로드 =====
def load_body_map():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except FileNotFoundError:
        cfg = {
            "remind_message": "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다.",
            "auto_start": False,
            "verbose": False
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

    # Backfill
    cfg.setdefault("remind_message", "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다.")
    cfg.setdefault("auto_start", False)
    cfg.setdefault("verbose", False)
    return cfg

REMIND_SUBJECT_PREFIX = "[Remind] "
REMIND_BODY_HTML_TOP_DEFAULT = (
    "<p>안녕하십니까, 씨넷 CS팀 ㅇㅇㅇ엔지니어입니다.</p>"
    "<p>지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드 드립니다.</p>"
    "<p>확인 부탁드립니다. 감사합니다.</p>"
)

# ---- MAPI property tags
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_TRANSPORT_HEADERS   = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
PR_ATTACH_CONTENT_ID   = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"

BAD_CID_DENYLIST = (
    "filelist.html", "filelist.xml", "themedata.thmx",
    "colorschememapping.xml", "editdata.mso",
)

VERBOSE = False

from datetime import datetime  # 이미 있으면 생략

def _pretty_ts(ts: str) -> str:
    """ISO/텍스트 형태의 타임스탬프를 'YYYY-MM-DD HH:MM' 로 변환."""
    if not ts or ts == "-":
        return "-"
    try:
        # '2025-10-13T09:40:01.799235' → datetime → '2025-10-13 09:40'
        return datetime.fromisoformat(ts.replace("Z","")).strftime("%Y-%m-%d %H:%M")
    except Exception:
        # 혹시 비표준 문자열이어도 최대한 정리
        s = ts.replace("T", " ")
        if "." in s:
            s = s.split(".", 1)[0]
        return s[:16]  # YYYY-MM-DD HH:MM 까지만


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if level == "DEBUG" and not VERBOSE:
        return
    line = f"{ts} [{level}] {msg}"
    print(line)
    week_str = datetime.now().strftime("%Y-W%U")
    log_file = os.path.join(APPDATA_DIR, f"app_{week_str}.log")
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"[WARN] 로그 파일 기록 실패: {e}")

def format_body_text(text):
    if not text: return ""
    processed_text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    processed_text = processed_text.replace('\n', '<br>')
    return f"<p>{processed_text}</p>"

def to_local_naive(dt):
    if dt is None or not isinstance(dt, datetime): return None
    if dt.tzinfo is None: return dt
    try:
        return dt.astimezone(ZoneInfo("Asia/Seoul")).replace(tzinfo=None)
    except Exception:
        return dt.replace(tzinfo=None)

def now_naive():
    return datetime.now()

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_state(st):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_FILE)

def parse_yard_tag(subject):
    """[SHI3D], [HMD12H], [HHI1W], [HSHI30MIN] → (yard, interval_days)"""
    if not subject: return None, None
    s = subject.upper().replace("［", "[").replace("］", "]")
    m = re.search(r"\[(SHI|HMD|HHI|HSHI|HO|HJSC)(\d+)(MIN|H|D|W|M)\]", s)
    if not m: return None, None
    yard, num, unit = m.group(1), int(m.group(2)), m.group(3)
    if unit == "MIN": interval_days = num / 1440.0
    elif unit == "H": interval_days = num / 24.0
    elif unit == "D": interval_days = float(num)
    elif unit == "W": interval_days = float(num) * 7.0
    elif unit == "M": interval_days = float(num) * 30.0
    else: interval_days = None
    return yard, interval_days

PREFIXES = ["re:", "fw:", "fwd:", "답장:", "회신:", "전달:", "참조:", "回覆:", "転送:"]

def strip_brackets_tags(subj: str) -> str:
    s = re.sub(r"^\s*\[remind\]\s*","", subj or "", flags=re.I)
    s = re.sub(r"\s*\[(?:D[ANH]|F[UIP])\s*\d+\s*(?:MIN|H|D|W|M)\]\s*"," ", s, flags=re.I)
    return re.sub(r"\s+"," ", s).strip()

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
    s = strip_brackets_tags(s)
    return s.lower()

def _walk_folders(folder):
    yield folder
    for i in range(1, folder.Folders.Count+1):
        sub = folder.Folders.Item(i)
        for f in _walk_folders(sub):
            yield f

def _get_deleted_roots(ns):
    roots = []
    try:
        for store in ns.Stores:
            try:
                di = store.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
            except Exception:
                di = None
            if di:
                try:
                    roots.append(di.FolderPath)
                except Exception:
                    pass
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

def get_internet_message_id(item):
    try: return item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    except Exception: return None

def get_transport_headers(item):
    try: return item.PropertyAccessor.GetProperty(PR_TRANSPORT_HEADERS) or ""
    except Exception: return ""

def headers_contain_original(headers, orig_msgid):
    if not headers or not orig_msgid: return False
    return orig_msgid.strip().lower() in headers.lower()

def get_sender_address(mail_item):
    try:
        addr = getattr(mail_item, "SenderEmailAddress", None)
        if addr:
            return addr.lower()
    except Exception:
        pass
    try:
        ae = getattr(mail_item, "Sender", None)
        if ae:
            return (ae.Address or ae.Name).lower()
    except Exception:
        pass
    return None

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
            log(f"[CANCELLED-SKIP] {state_key} is cancelled; skip sending.")
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
                    if m.Class != OL_MAILITEM:
                        continue
                    if is_from_me(m, me_set):
                        continue
                    rt = to_local_naive(getattr(m, "ReceivedTime", None))
                    if not rt or rt <= orig_sent:
                        continue

                    can = canonicalize_subject(getattr(m, "Subject", "") or "")
                    base = canonicalize_subject(orig_subject)
                    if not base:
                        continue
                    ok = (can == base) if len(base) < 8 else ((can == base) or (base in can) or (can in base))
                    if not ok:
                        continue

                    sender_addr = (m.SenderEmailAddress or "").lower()
                    if addr.lower() in sender_addr:
                        if verbose:
                            log(f"[REPLY*:{rtype}] {folder.FolderPath} / {rt:%Y-%m-%d %H:%M:%S} / "
                                f"{m.SenderName} / {m.Subject} / matched={addr}")
                        state[state_key] = {
                            "reply_received": True,
                            "last_sent": state.get(state_key, {}).get("last_sent"),
                            "detected_at": rt.isoformat()
                        }
                        break
                except Exception:
                    continue

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
    try: pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid)
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
    if not isinstance(html, str) or not html:
        return html
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
    if verbose:
        try: log("[SANITIZE] stripped office junk cid refs")
        except: pass
    return html

def _ensure_existing_cids_have_attachments(mail, html, verbose=False):
    if not html: return html
    cids=set()
    for m in re.finditer(r'cid:([^\s"\'>)]+)', html, flags=re.I):
        cids.add(m.group(1))
    if not cids: return html
    existing=set()
    try:
        atts=mail.Attachments
        for i in range(1, atts.Count+1):
            att=atts.Item(i)
            try:
                cid=att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
                if cid: existing.add(cid)
            except Exception:
                continue
    except Exception:
        pass
    for cid in cids:
        low = cid.lower()
        if any(x in low for x in BAD_CID_DENYLIST):
            if verbose: log(f"[CID-FIX-SKIP] denylist cid={cid}")
            continue
        if cid in existing: continue
        path=_resolve_signature_path(cid)
        if not path or not os.path.exists(path):
            base = os.path.expanduser(r"~\AppData\Roaming\Microsoft\Signatures")
            if os.path.isdir(base):
                for root,_,files in os.walk(base):
                    for fn in files:
                        if fn.lower().startswith(("image00","logo")) and os.path.splitext(fn)[1].lower() in (".png",".jpg",".jpeg",".gif",".bmp"):
                            test=os.path.join(root,fn)
                            if os.path.exists(test): path=test; break
                    if path: break
        if not path or not os.path.exists(path): continue
        try:
            att=mail.Attachments.Add(path)
            _mark_attachment_inline(att, cid, path, verbose=True)
            if verbose: log(f"[CID-FIX] attached for missing cid={cid} -> {path}")
        except Exception as e:
            if verbose: log(f"[CID-FIX-ERR] cid={cid} err={e}")
    return html

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
        try:
            os.startfile("outlook.exe")
        except Exception:
            pass
    t0 = time.time()
    while time.time()-t0 < 30:
        try:
            return win32.Dispatch("Outlook.Application")
        except Exception:
            time.sleep(2)
    raise RuntimeError("Outlook COM attach failed")

def _safe_recipients_from(original_item):
    def _names_from(recips, t=1):
        out=[]
        if not recips: return out
        for i in range(1, recips.Count+1):
            r = recips.Item(i)
            if getattr(r,"Type",1)!=t: continue
            addr=None
            try:
                ae=r.AddressEntry
                if ae and ae.Type=="EX":
                    ex=ae.GetExchangeUser()
                    if ex and ex.PrimarySmtpAddress: addr=ex.PrimarySmtpAddress
                elif ae:
                    addr=ae.Address
            except Exception:
                pass
            if not addr: addr=getattr(r,"Name",None)
            if addr: out.append(addr)
        return out
    recips = getattr(original_item,"Recipients",None)
    to_list = _names_from(recips,1)
    cc_list = _names_from(recips,2)
    raw_to = (getattr(original_item,"To","") or "").strip()
    raw_cc = (getattr(original_item,"CC","") or "").strip()
    if (not to_list) and raw_to: to_list=[raw_to]
    if (not cc_list) and raw_cc: cc_list=[raw_cc]
    return "; ".join([x for x in to_list if x]), "; ".join([x for x in cc_list if x])

def send_remind_for_recipients(app, item, subject, body, yard_code, state, dry_run=False, verbose=False):
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
                    if addr:
                        recipients.append((addr, r.Type))
            except Exception:
                continue

        if not recipients:
            if verbose: log("[WARN] no To/BCC recipients on original mail")
            return False

        remind_text = load_body_map().get("remind_message", body or "")
        remind_html = format_body_text(remind_text)

        me_addr = _self_smtp() or getattr(item, "SenderEmailAddress", None) or "me@example.com"
        sent_any = False

                # ✅ [추가] 발송 취소된 key 목록 불러오기
        try:
            st_snapshot = load_state()
        except Exception:
            st_snapshot = {}
        cancelled_keys = set(st_snapshot.get("__cancelled_keys__", []))

        for addr, rtype in recipients:
            state_key = make_state_key(item.EntryID, addr)
             # ✅ 이번 메일(EntryID|email)만 취소되어 있으면 무조건 스킵
            if state_key in cancelled_keys:
                log(f"[CANCELLED-SKIP] {state_key} is cancelled; skip sending.")
                continue

            if state.get(state_key, {}).get("reply_received", False):
                continue

            fwd = item.Forward()
            fwd.Subject = f"[Remind] {subject}"
            fwd.BodyFormat = 2  # HTML
            fwd.HTMLBody = (
                "<div style='font-family:Malgun Gothic,Segoe UI,Arial,sans-serif; font-size:10pt;'>"
                f"{remind_html}"
                "</div><br>" + fwd.HTMLBody
            )
            fwd.HTMLBody = _sanitize_bad_cids(_attach_images_and_rewrite_html(fwd, fwd.HTMLBody))

            if rtype == 1:      # To
                fwd.To = addr
            else:               # BCC
                if not fwd.To:
                    fwd.To = me_addr
                recip = fwd.Recipients.Add(addr)
                recip.Type = 3

            try:
                fwd.Recipients.ResolveAll()
            except Exception:
                pass

            if dry_run:
                log(f"[DRY-RUN] Would send | {fwd.Subject} ({yard_code}) | To={fwd.To}, CC={fwd.CC}, BCC={fwd.BCC}")
                continue

            try:
                fwd.Save()
                fwd.Send()
                sent_any = True
                ts = now_naive().isoformat()
                state[state_key] = {"reply_received": False, "last_sent": ts, "subject": subject}
                save_state(state)  # <- 반드시 즉시 디스크 반영

                # ② 로그는 INFO 등급으로 항상 찍어 가시성 확보 (verbose 없이도 보임)
                log(f"[SENT] To={fwd.To}, CC={fwd.CC}, BCC={fwd.BCC} | {fwd.Subject} ({yard_code})")
                log(f"[STATE-UPD] {state_key} reply_received=False last_sent={ts}")
            except Exception as e:
                log(f"[ERR-SEND] {e}")

        return True if (dry_run or sent_any) else False

    except Exception as e:
        log(f"[ERR-SEND] {e}")
        return False

def is_empty_draft(item):
    try:
        if getattr(item, "Class", None) != OL_MAILITEM:
            return False
        subj = (item.Subject or "").strip().lower()
        body = (item.Body or "").strip()
        to   = (item.To or "").strip()
        cc   = (item.CC or "").strip()
        if not body and not to and not cc:
            return True
        return False
    except Exception:
        return False

def cleanup_empty_drafts_and_deleted(ns, verbose=False):
    removed = 0
    scanned = 0
    try:
        drafts = ns.GetDefaultFolder(OL_FOLDER_DRAFTS)
        for item in list(drafts.Items):
            scanned += 1
            if is_empty_draft(item):
                if verbose: log(f"[CLEANUP] Removing from Drafts: {item.Subject}")
                item.Delete()
                removed += 1
    except Exception as e:
        if verbose: log(f"[CLEANUP-ERR] Drafts: {e}")

    try:
        deleted = ns.GetDefaultFolder(OL_FOLDER_DELETED)
        for item in list(deleted.Items):
            scanned += 1
            if is_empty_draft(item):
                if verbose: log(f"[CLEANUP] Purging from Deleted Items: {item.Subject}")
                item.Delete()
                removed += 1
    except Exception as e:
        if verbose: log(f"[CLEANUP-ERR] Deleted: {e}")

    if verbose:
        log(f"[CLEANUP] scanned={scanned}, fully removed={removed}")

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
                    if verbose: log(f"[SKIP-STALE] tag too old: {age_h:.1f}h > {max_age_hours}h")
                    continue

            if verbose:
                log(f"[DUE] base={'last_remind_at' if (due_from_last and last_sent_iso and base_time!=sent_on) else 'sent_on'} "
                    f"| base_time={base_time:%Y-%m-%d %H:%M} | due_time={due_time:%Y-%m-%d %H:%M} | due_ok={due_ok}")

            if (not force_send) and (not due_ok):
                remaining = (due_time - now_ts).total_seconds()
                if remaining > precheck_epsilon_sec:
                    if verbose:
                        log(f"[PRECHECK-SKIP] due in {remaining:.1f}s (> {precheck_epsilon_sec}s)")
                    if (time.time() - loop_started) > loop_budget_sec:
                        log(f"[LOOP-BUDGET] elapsed={time.time() - loop_started:.1f}s > {loop_budget_sec}s, defer rest to next scan")
                        break
                    continue

            if last_sent_iso and not force_send:
                try:
                    last_dt = to_local_naive(datetime.fromisoformat(last_sent_iso))
                    if last_dt and now_ts - last_dt < timedelta(days=interval_days):
                        if verbose: log("[SKIP] within interval since last remind")
                        if (time.time() - loop_started) > loop_budget_sec:
                            log(f"[LOOP-BUDGET] elapsed={time.time() - loop_started:.1f}s > {loop_budget_sec}s, defer rest to next scan")
                            break
                        continue
                except Exception:
                    pass

            if skip_if_newer_outgoing:
                canon = canonicalize_subject(subject or "")
                if _has_newer_outgoing_with_same_subject(ns, canon, sent_on,
                                                         include_deleted=include_deleted,
                                                         verbose=verbose):
                    if verbose: log("[SKIP] newer outgoing exists in same thread")
                    if (time.time() - loop_started) > loop_budget_sec:
                        log(f"[LOOP-BUDGET] elapsed={time.time() - loop_started:.1f}s > {loop_budget_sec}s, defer rest to next scan")
                        break
                    continue

            if not skip_reply_check:
                try:
                    if verbose:
                        log(f"[DEBUG-REPLYCHK] subj='{subject}' conv_id={mail.ConversationID} "
                            f"topic='{mail.ConversationTopic}' check_after={sent_on:%Y-%m-%d %H:%M}")

                    check_and_update_replies(app, mail, state, verbose=verbose)
                    save_state(state)
                except Exception as e:
                    log(f"[ERR-REPLYCHK] {e}")

            if (not force_send) and (not due_ok):
                if verbose: log("[SKIP] not yet due")
                if (time.time() - loop_started) > loop_budget_sec:
                    log(f"[LOOP-BUDGET] elapsed={time.time() - loop_started:.1f}s > {loop_budget_sec}s, defer rest to next scan")
                    break
                continue

            if dry_run:
                log(f"[DRY-RUN] Would send | {subject} ({code})")
            else:
                ok = send_remind_for_recipients(
                    app,
                    mail,
                    subject,
                    load_body_map().get("remind_message", ""),
                    code,
                    state,
                    dry_run=dry_run,
                    verbose=verbose
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

def _has_newer_outgoing_with_same_subject(ns, canon_subj: str, sent_on: datetime, include_deleted=False, verbose=False):
    try:
        stores = list(ns.Stores)
    except Exception:
        return False
    deleted_roots = [] if include_deleted else _get_deleted_roots(ns)

    for store in stores:
        try:
            sent = store.GetDefaultFolder(OL_FOLDER_SENT)
        except Exception:
            continue
        for f in _walk_folders(sent):
            try:
                if (not include_deleted) and _is_under_deleted(f, deleted_roots):
                    continue
                items = f.Items
                items.Sort("[SentOn]", True)
                items = items.Restrict("[SentOn] >= '" + (sent_on.strftime('%m/%d/%Y %I:%M %p')) + "'")
            except Exception:
                continue

            try:
                enum = iter(items)
            except Exception:
                def enum_iter(it):
                    for i in range(1, it.Count+1):
                        yield it.Item(i)
                enum = enum_iter(items)

            for it in enum:
                try:
                    if it.Class != OL_MAILITEM:
                        continue
                    s = canonicalize_subject(getattr(it, "Subject", "") or "")
                    if s != canon_subj:
                        continue
                    so = to_local_naive(getattr(it, "SentOn", None))
                    if so and so > sent_on:
                        subj = getattr(it, "Subject", "") or ""
                        if re.search(r"^\s*\[remind\]\s*", subj, flags=re.I):
                            continue
                        if verbose: log(f"[SKIP-NEWER-OUT] newer outgoing found at {so:%Y-%m-%d %H:%M}")
                        return True
                except Exception:
                    continue
    return False

# ---- App wiring ----
exit_event = threading.Event()

def start_mail_check_loop(args):
    st = load_state()
    while not exit_event.is_set():
        try:
            # ✅ 트레이에서 취소/설정 변경 반영을 위해 매 사이클마다 최신 state 로드
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

# Tk root (main thread) — single instance for all Toplevels
root = tk.Tk()
set_window_icon(root)
root.withdraw()

def create_and_show_gui():
    top = tk.Toplevel(root)
    set_window_icon(top)
    top.title("리마인드 메시지 설정")
    top.geometry("640x420")
    top.transient()  # keep above if possible
    top.grab_set()

    current_config = load_body_map()

    guide_txt = (
        "■ 제목 코드 작성 예시\n"
        "  • [SHI1DAY] SN0000 Vendor 정보 요청 드립니다 → SHI 삼성중공업, 하루 마다 리마인드 요청\n"
        "  • [HMD3H] H0000 Vendor 정보 요청 드립니다  → HMD 현대미포조선, 3시간 뒤 리마인드 요청\n"
        "  • 메일 제목의 호선 이름은 항상 Full name으로 작성바랍니다. (ex:SN2693/H3525/H.2378)\n"
        "  • 본문은 아래 입력된 메시지 하나로 고정되며, 사용자가 자유롭게 수정할 수 있습니다.\n"
        "  • 서식: **굵게**, *기울임*, 줄바꿈은 Enter. 목록은 - 또는 • 사용 가능."
    )
    guide = tk.Label(top, text=guide_txt, justify="left", anchor="w",
                     fg="#555555", wraplength=600)
    guide.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6))

    tk.Label(top, text="리마인드 메시지:").grid(row=1, column=0, sticky="nw", padx=10, pady=5)
    frame = tk.Frame(top)
    frame.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")
    top.grid_rowconfigure(1, weight=1)
    top.grid_columnconfigure(1, weight=1)

    text_widget = tk.Text(frame, width=80, height=12, wrap=tk.WORD)
    scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    text_widget.insert("1.0", current_config.get("remind_message", ""))

    auto_start_var = tk.BooleanVar(value=current_config.get("auto_start", False))
    verbose_var    = tk.BooleanVar(value=current_config.get("verbose", False))

    cb_autostart = tk.Checkbutton(top, text="Windows 시작 시 자동 실행", variable=auto_start_var)
    cb_autostart.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 0))

    cb_verbose = tk.Checkbutton(top, text="DEBUG 로그 출력 (Verbose 모드)", variable=verbose_var)
    cb_verbose.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 2))

    def save_config():
        new_cfg = {
            "remind_message": text_widget.get("1.0", tk.END).strip(),
            "auto_start": auto_start_var.get(),
            "verbose": verbose_var.get()
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(new_cfg, f, ensure_ascii=False, indent=2)

            # apply immediately
            global VERBOSE
            VERBOSE = new_cfg["verbose"]
            try:
                if new_cfg["auto_start"]:
                    register_startup_reg(APP_RUN_NAME)
                else:
                    unregister_startup_reg(APP_RUN_NAME)
            except Exception as e:
                log(f"[WARN] Startup 토글 중 오류: {e}", level="WARN")

            messagebox.showinfo("저장 완료", "설정이 성공적으로 저장되었습니다.")
            top.destroy()
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    btn_frame = tk.Frame(top)
    btn_frame.grid(row=4, column=1, sticky="e", padx=10, pady=10)
    tk.Button(btn_frame, text="저장", command=save_config).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="닫기", command=top.destroy).pack(side=tk.LEFT)

def open_settings_window(icon, item):
    # Schedule on Tk main thread
    try:
        root.after(0, create_and_show_gui)  # NOTE: do NOT call create_and_show_gui()
    except Exception as e:
        log(f"[ERR] open_settings_window: {e}")

def open_remind_list_window(icon, item):
    def _show():
        top = tk.Toplevel(root)
        set_window_icon(top)
        top.title("현재 Remind 리스트")
        top.geometry("700x420")

        tree = ttk.Treeview(
            top,
            columns=("recipient", "subject", "last_sent"),
            show="headings",
            selectmode="extended"
        )
        tree.heading("recipient", text="Recipient")
        tree.heading("subject",   text="Subject")     # ✅ 제목 컬럼 추가
        tree.heading("last_sent", text="Last Sent")
        # (선택) 폭 약간 조정
        tree.column("recipient", width=230, anchor="w")
        tree.column("subject",   width=320, anchor="w")
        tree.column("last_sent", width=140, anchor="center")

        tree.pack(fill="both", expand=True)

        def populate():
            tree.delete(*tree.get_children())

            st = load_state()

            # 수신인별 '가장 최신' 미회신(=계속 리마인드) 항목만 집계
            latest_by_addr = {}
            for key, val in st.items():
                if "|" not in key:
                    continue
                entry_id, addr = key.split("|", 1)

                # 회신 온 건 제외
                if val.get("reply_received", False):
                    continue

                sent_time = val.get("last_sent")
                try:
                    sent_dt = datetime.fromisoformat(sent_time) if sent_time else None
                except Exception:
                    sent_dt = None

                prev = latest_by_addr.get(addr)
                if not prev:
                    latest_by_addr[addr] = (sent_dt, val, key)
                else:
                    prev_dt, _, _ = prev
                    # 더 최신 발송 기준으로 교체
                    if (sent_dt or datetime.min) > (prev_dt or datetime.min):
                        latest_by_addr[addr] = (sent_dt, val, key)

            # 화면에 뿌리기
            for addr, (sent_dt, val, key) in latest_by_addr.items():
                # tree.insert(
                #     "",
                #     "end",
                #     iid=key,
                #     values=(addr, val.get("last_sent", "-"))
                # )
                # 화면에 뿌리기
                subj = val.get("subject", "-")                 # ✅ 제목 사용
                last = _pretty_ts(val.get("last_sent", "-"))   # ✅ 시간 포맷 적용
                tree.insert("", "end", iid=key, values=(addr, subj, last))



        def delete_selected():
            sel = tree.selection()      # sel = state_key 들
            st = load_state()

            # 전역 차단 목록(키 단위) 준비
            cancelled_keys = set(st.get("__cancelled_keys__", []))

            for key in sel:
                # 1) 현재 키(state_key = EntryID|email) 히스토리 제거(선택)
                st.pop(key, None)
                # 2) 이 키만 차단 목록에 추가
                cancelled_keys.add(key)
                # 3) UI에서 제거
                tree.delete(key)

            # 차단 목록 저장
            st["__cancelled_keys__"] = sorted(cancelled_keys)
            save_state(st)


        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="새로고침", command=populate).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="선택 삭제(발송 취소)", command=delete_selected).pack(side=tk.LEFT, padx=6)

        populate()

    root.after(0, _show)

def exit_action(icon, item):
    log("[INFO] Exit requested. Shutting down.")
    exit_event.set()
    try:
        icon.stop()
    except Exception:
        pass
    try:
        root.after(0, root.quit)  # stop Tk mainloop
    except Exception:
        pass

def make_state_key(entry_id, recipient_addr):
    return f"{entry_id}|{(recipient_addr or '').lower()}"

def main():
    check_single_instance()
    cfg = load_body_map()

    parser = argparse.ArgumentParser(description="Automated Outlook Mail Reminder System (Fixed)")
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
    parser.add_argument("--reply-mode", choices=["hdr-only", "hdr-first", "conv-first"], default="conv-first")
    args = parser.parse_args()

    global VERBOSE
    VERBOSE = args.verbose or cfg.get("verbose", False)

    # startup toggle from config at boot
    try:
        if cfg.get("auto_start", False):
            register_startup_reg(APP_RUN_NAME)
        else:
            unregister_startup_reg(APP_RUN_NAME)
    except Exception as e:
        log(f"[WARN] Startup 토글 초기화 오류: {e}", level="WARN")

    log(f"[INFO] Verbose mode = {VERBOSE}")

    # background worker
    mail_thread = threading.Thread(target=start_mail_check_loop, args=(args,), daemon=True)
    mail_thread.start()
    log("[INFO] Mail check background thread started.")

    # Tray icon (detached so Tk mainloop can run on main thread)
    def resource_path(filename):
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, filename)
        return os.path.join(os.path.abspath("."), filename)

    try:
        img = Image.open(resource_path("icon.png"))
    except FileNotFoundError:
        log("[WARN] 'icon.png' not found. Using default blue square.")
        img = Image.new('RGB', (64, 64), color='blue')

    menu = (
        item('설정 열기', open_settings_window),
        item('리마인드 리스트', open_remind_list_window),
        item('종료', exit_action))
    tray = icon('AutoMailSystem', img, "자동 메일 리마인더", menu)

    show_startup_notification()
    tray.run_detached()  # <-- key difference

    # Enter Tk main loop (for settings windows)
    root.mainloop()

if __name__ == "__main__":
    main()
