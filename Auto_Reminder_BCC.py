# auto_mail_system_final_complete.py
# 2025-09 — Loop-budget reset fix, precheck-before-heavy, due calc reorder,
#            sanitize/signature CID handling, reply detection modes, deleted exclusion
# Integrated with System Tray, GUI Config, and Markup Support.
#
# Tags: DN/DH/DA + FU/FI/FP compatibility
#
# Usage examples:
#   (This script is now designed to run as a background application)

import os, re, json, time, uuid, argparse, urllib.parse, pythoncom, threading
from datetime import datetime, timedelta
import winreg
import sys
import win32com.client as win32
import win32event
import win32api
import winerror
import win32event
from win10toast import ToastNotifier
import ctypes

# GUI and Tray Icon libraries
import tkinter as tk
from tkinter import messagebox
from PIL import Image
from pystray import Icon as icon, MenuItem as item

# ---- Global / Base Paths ----
import os
LAST_CLEANUP = 0
CLEANUP_INTERVAL_SEC = 300
BODY_MAP = {}

# AppData 디렉토리
APPDATA_DIR = os.path.join(os.environ.get("APPDATA", os.getcwd()), "AutoRemindCS")
os.makedirs(APPDATA_DIR, exist_ok=True)

# 주요 파일 경로
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
import winreg
import os
import os, sys, winreg

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "AutoRemindCS"  # 시작프로그램에 표시될 이름

def _exe_path():
    # PyInstaller onefile이면 sys.executable, 스크립트 실행이면 argv[0]
    try:
        return sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
    except Exception:
        return os.path.abspath(sys.argv[0])

def register_startup_reg(app_name=APP_RUN_NAME, extra_args=""):
    """작업관리자 시작프로그램에 보이는 HKCU\\...\\Run 등록"""
    path = _exe_path()
    cmd = f'"{path}" {extra_args}'.strip()   # 공백 경로 대비 따옴표 포함
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
        # 등록 안 되어 있었음
        log("[INFO] Startup 해제(HKCU\\Run): 기존 등록 없음")
    except Exception as e:
        log(f"[ERR] Startup 해제 실패(HKCU\\Run): {e}", level="ERR")

def is_startup_registered_reg(app_name=APP_RUN_NAME):
    """디버그용: 현재 등록 상태/명령어 반환"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ) as k:
            val, _ = winreg.QueryValueEx(k, app_name)
            return True, val
    except FileNotFoundError:
        return False, ""
    except Exception:
        return False, ""
    
def check_single_instance(mutex_name="AutoRemindCS_Mutex"):
    handle = win32event.CreateMutex(None, False, mutex_name)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, "이미 Auto_Reminder가 실행 중입니다.", "Auto Reminder", 0x40)
        sys.exit(0)

def show_startup_notification():
    ctypes.windll.user32.MessageBoxW(
        0,
        "백그라운드에서 Auto Reminder가 실행 중입니다.",
        "Auto Reminder 실행됨",
        0x40  # MB_ICONINFORMATION
    )



# ===== 설정 파일 로드 =====
def load_body_map():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except FileNotFoundError:
        # 기본 템플릿 생성
        cfg = {
            "remind_message": "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다.",
            "auto_start": False,
            "verbose": False
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

    # 누락 키 보완
    if "remind_message" not in cfg:
        cfg["remind_message"] = "지난 메일 관련하여 아직 회신이 확인되지 않아 정중히 리마인드드립니다."
    if "auto_start" not in cfg:
        cfg["auto_start"] = False
    if "verbose" not in cfg:
        cfg["verbose"] = False

    return cfg


# ===== 설정창 GUI =====
def create_and_show_gui():
    top = tk.Tk()
    top.title("리마인드 메시지 설정")

    current_config = load_body_map()

    # --- 안내 라벨 ---
    guide_txt = (
        "■ 제목 코드 작성 예시\n"
        "  • [SHI1DAY] SN0000 Vendor 정보 요청 드립니다 → SHI 삼성중공업, 하루 마다 리마인드 요청\n"
        "  • [HMD3H] H0000 Vendor 정보 요청 드립니다  → HMD 현대미포조선, 3시간 뒤 리마인드 요청\n"
        "  • 메일 제목의 호선 이름은 항상 Full name으로 작성바랍니다. (ex:SN2693/H3525/H.2378)\n"
        "  • 본문은 아래 입력된 메시지 하나로 고정되며, 사용자가 자유롭게 수정할 수 있습니다.\n"
        "  • 서식: **굵게**, *기울임*, 줄바꿈은 Enter. 목록은 - 또는 • 사용 가능."
    )
    guide = tk.Label(top, text=guide_txt, justify="left", anchor="w",
                     fg="#555555", wraplength=760)
    guide.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6))

    # --- 리마인드 메시지 입력칸 ---
    tk.Label(top, text="리마인드 메시지:").grid(row=1, column=0, sticky="nw", padx=10, pady=5)
    frame = tk.Frame(top)
    frame.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")
    top.grid_rowconfigure(1, weight=1)
    top.grid_columnconfigure(1, weight=1)

    text_widget = tk.Text(frame, width=80, height=10, wrap=tk.WORD)
    scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)

    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    text_widget.insert("1.0", current_config.get("remind_message", ""))

    # --- 체크박스 ---
    auto_start_var = tk.BooleanVar(value=current_config.get("auto_start", False))
    verbose_var    = tk.BooleanVar(value=current_config.get("verbose", False))

    cb_autostart = tk.Checkbutton(top, text="Windows 시작 시 자동 실행", variable=auto_start_var)
    cb_autostart.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 0))

    cb_verbose = tk.Checkbutton(top, text="DEBUG 로그 출력 (Verbose 모드)", variable=verbose_var)
    cb_verbose.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 2))

    # --- 저장 버튼 ---
    def save_config():
        new_cfg = {
            "remind_message": text_widget.get("1.0", tk.END).strip(),
            "auto_start": auto_start_var.get(),
            "verbose": verbose_var.get()
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(new_cfg, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", "설정이 성공적으로 저장되었습니다.")
            top.destroy()
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    btn_frame = tk.Frame(top)
    btn_frame.grid(row=4, column=1, sticky="e", padx=10, pady=10)
    tk.Button(btn_frame, text="저장", command=save_config).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="취소", command=top.destroy).pack(side=tk.LEFT)

    # --- 저장/취소 버튼 ---
    def save_config():
        global BODY_MAP, VERBOSE
        new_cfg = {}

        # DN/DH/DA 텍스트 취합
        for k, widget in entries.items():
            new_cfg[k] = widget.get("1.0", tk.END).strip()

        # 체크박스 반영
        new_cfg["auto_start"] = auto_start_var.get()
        new_cfg["verbose"]    = verbose_var.get()

        try:
            # config 저장
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(new_cfg, f, ensure_ascii=False, indent=2)

            # 전역 반영
            BODY_MAP = {k: v for k, v in new_cfg.items() if k in ("DN", "DH", "DA")}
            VERBOSE  = new_cfg["verbose"]

            # 자동실행 토글 (레지스트리 방식 사용 시)
            # 자동실행 토글
            try:
                if new_cfg["auto_start"]:
                    register_startup_reg(APP_RUN_NAME)      # ← 레지스트리 방식
                else:
                    unregister_startup_reg(APP_RUN_NAME)    # ← 레지스트리 방식
            except Exception as e:
                log(f"[WARN] Startup 토글 중 오류: {e}", level="WARN")


            log("[INFO] Configuration saved and updated.")
            messagebox.showinfo("저장 완료", "설정이 성공적으로 저장되었습니다.")
            top.destroy()

        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {e}")

    btn_frame = tk.Frame(top)
    btn_frame.grid(row=4, column=1, sticky="e", padx=10, pady=10)

    tk.Button(btn_frame, text="저장", command=save_config).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="취소", command=top.destroy).pack(side=tk.LEFT)




BODY_MAP = load_body_map()
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

# ---- Office junk CID denylist + sanitizer
BAD_CID_DENYLIST = (
    "filelist.html", "filelist.xml", "themedata.thmx",
    "colorschememapping.xml", "editdata.mso",
)

from datetime import datetime
import os

# 전역 플래그: argparse에서 --verbose 받아서 VERBOSE=True/False로 세팅하세요
VERBOSE = False  

def log(msg, level="INFO"):
    """
    레벨별 로그 출력 + AppData 주간 로그 파일 기록
    - INFO / WARN / ERR : 항상 출력
    - DEBUG             : VERBOSE=True일 때만 출력
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # DEBUG는 verbose 옵션일 때만 허용
    if level == "DEBUG" and not VERBOSE:
        return

    line = f"{ts} [{level}] {msg}"

    # 콘솔 출력
    print(line)

    # 파일 저장 (주간 로그)
    week_str = datetime.now().strftime("%Y-W%U")
    log_file = os.path.join(APPDATA_DIR, f"app_{week_str}.log")
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"[WARN] 로그 파일 기록 실패: {e}")

# ---- Markup to HTML Converter
def format_body_text(text):
    """Converts simple markup to HTML for the email body."""
    if not text: return ""
    # 1. Convert **bold** to <b>bold</b>
    processed_text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    # 2. Convert newlines to <br> tags
    processed_text = processed_text.replace('\n', '<br>')
    # 3. Wrap the whole text in a paragraph tag
    return f"<p>{processed_text}</p>"

# ---- START OF ORIGINAL SCRIPT FUNCTIONS ----

def to_local_naive(dt):
    if dt is None or not isinstance(dt, datetime):
        return None
    if dt.tzinfo is None:
        return dt
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
    """
    제목에서 [조선소코드+기간] 패턴 인식
    예: [SHI3D], [HMD12H], [HHI1W], [HSHI30MIN]
    반환: (yard_code, interval_days)
    """
    if not subject:
        return None, None

    s = subject.upper().replace("［", "[").replace("］", "]")
    m = re.search(r"\[(SHI|HMD|HHI|HSHI|HO|HJSC)(\d+)(MIN|H|D|W|M)\]", s)
    if not m:
        return None, None

    yard = m.group(1)              # 조선소 코드
    num = int(m.group(2))          # 숫자
    unit = m.group(3)              # 시간 단위

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
    """
    Outlook MailItem에서 발신자 SMTP 주소를 안전하게 추출
    """
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
    """
    각 To/BCC 수신자별 회신 여부 업데이트
    """
    ns = app.GetNamespace("MAPI")
    me_set = my_addresses(ns)

    orig_subject = orig_mail.Subject or ""
    orig_sent = to_local_naive(getattr(orig_mail, "SentOn", None))
    recipients = [(r.Address, r.Type) for r in orig_mail.Recipients if r.Type in (1, 3)]

    for addr, rtype in recipients:
        state_key = make_state_key(orig_mail.EntryID, addr)

        # 이미 회신 처리된 경우 스킵
        if state.get(state_key, {}).get("reply_received", False):
            continue

        # 받은 편지함/스팸/삭제됨 등 전체 탐색
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

                    # 제목 매칭 (FUZZY)
                    can = canonicalize_subject(getattr(m, "Subject", "") or "")
                    base = canonicalize_subject(orig_subject)
                    if not base:
                        continue
                    if len(base) < 8:
                        ok = (can == base)
                    else:
                        ok = (can == base) or (base in can) or (can in base)

                    if not ok:
                        continue

                    # ✅ 발신자가 대상 수신자 주소인지 확인
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
    # Outlook이 이미 떠 있는지 확인 후 없을 때만 실행
    for proc in os.popen('tasklist').read().splitlines():
        if "OUTLOOK.EXE" in proc.upper():
            break
    else:
        try:
            os.startfile("outlook.exe")
        except Exception:
            pass
    # attach 재시도
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
    """
    To/BCC 각각을 개별 추적하여 리마인드 발송.
    - 원본 메일은 Forward()로 붙여 인라인 이미지/첨부 유지
    - BCC-only 시 To에 본인 주소를 채워 안전 발송
    """
    def _self_smtp():
        # 가능한 한 SMTP 주소 얻기
        try:
            ae = app.Session.CurrentUser.AddressEntry
            exu = ae.GetExchangeUser()
            return (exu.PrimarySmtpAddress if exu else ae.Address) or None
        except Exception:
            return None

    try:
        # 수신자 목록 수집 (To=1, BCC=3만)
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

        # Remind 상단 메시지 준비
        remind_text = load_body_map().get("remind_message", body or "")
        remind_html = format_body_text(remind_text)

        me_addr = _self_smtp() or getattr(item, "SenderEmailAddress", None) or "jhunpark@sea-net.co.kr"
        sent_any = False

        for addr, rtype in recipients:
            # 개별 수신자별 상태 키
            state_key = make_state_key(item.EntryID, addr)
            if state.get(state_key, {}).get("reply_received", False):
                continue

            # 원본을 포워드하여 히스토리/이미지 보존
            fwd = item.Forward()
            fwd.Subject = f"[Remind] {subject}"
            fwd.BodyFormat = 2  # olFormatHTML
            # Remind 메시지를 최상단에 삽입 (그 아래에 -----Original Message----- 가 붙음)
            fwd.HTMLBody = (
                "<div style='font-family:Malgun Gothic,Segoe UI,Arial,sans-serif; font-size:10pt;'>"
                f"{remind_html}"
                "</div><br>" + fwd.HTMLBody
            )
            fwd.HTMLBody = _sanitize_bad_cids(_attach_images_and_rewrite_html(fwd, fwd.HTMLBody))

            # 수신자 세팅
            if rtype == 1:      # To
                fwd.To = addr
            else:               # BCC
                if not fwd.To:  # BCC-only 보호
                    fwd.To = me_addr
                recip = fwd.Recipients.Add(addr)
                recip.Type = 3  # BCC

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
                if verbose:
                    log(f"[SENT] To={fwd.To}, CC={fwd.CC}, BCC={fwd.BCC} | {fwd.Subject} ({yard_code})")
                state[state_key] = {"reply_received": False, "last_sent": now_naive().isoformat()}
            except Exception as e:
                log(f"[ERR-SEND] {e}")  # 이 주소만 실패, 다음 주소 계속

        return True if (dry_run or sent_any) else False

    except Exception as e:
        log(f"[ERR-SEND] {e}")
        return False




def mark_reply_received(original_item, reply_item, state):
    """
    To + BCC 각각 개별적으로 reply_received = True 처리
    """
    try:
        replied_addr = reply_item.SenderEmailAddress.lower().strip()
        for recip in original_item.Recipients:
            if recip.Type in (1, 3):  # To, BCC만 추적
                if recip.Address and recip.Address.lower().strip() == replied_addr:
                    state_key = make_state_key(original_item.EntryID, recip.Address)
                    if state.get(state_key):
                        state[state_key]["reply_received"] = True
                        log(f"[INFO] Marked reply received for {replied_addr}")
    except Exception as e:
        log(f"[ERR-MARK-REPLY] {e}")


def check_reply_exists(original_item, replies, state):
    """
    수신된 reply가 To/BCC 중 누구의 것인지 확인해서 state 업데이트
    """
    try:
        for reply in replies:
            replied_addr = reply.SenderEmailAddress.lower().strip()
            for recip in original_item.Recipients:
                if recip.Type in (1, 3):
                    if recip.Address and recip.Address.lower().strip() == replied_addr:
                        state_key = make_state_key(original_item.EntryID, recip.Address)
                        if state.get(state_key) and not state[state_key]["reply_received"]:
                            state[state_key]["reply_received"] = True
                            log(f"[REPLY] {replied_addr} replied → stop reminders for this recipient")
    except Exception as e:
        log(f"[ERR-CHK-REPLY] {e}")


def is_empty_draft(item):
    try:
        if getattr(item, "Class", None) != OL_MAILITEM:
            return False
        subj = (item.Subject or "").strip().lower()
        body = (item.Body or "").strip()
        to   = (item.To or "").strip()
        cc   = (item.CC or "").strip()
        # Draft 찌꺼기의 특징: 제목 없음, 본문 없음, 받는사람 없음
        if not body and not to and not cc:
            return True
        return False
    except Exception:
        return False

def cleanup_empty_drafts_and_deleted(ns, verbose=False):
    removed = 0
    scanned = 0

    # 1) Drafts 폴더 정리
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

    # 2) Deleted Items 폴더 정리
    try:
        deleted = ns.GetDefaultFolder(OL_FOLDER_DELETED)
        for item in list(deleted.Items):
            scanned += 1
            if is_empty_draft(item):
                if verbose: log(f"[CLEANUP] Purging from Deleted Items: {item.Subject}")
                item.Delete()  # Deleted Items 안에서도 Delete → 완전 제거
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

            # --- reply check (개별 수신자 단위) ---
                        # --- 회신 감지 ---
            if not skip_reply_check:
                try:
                    if verbose:
                        log(f"[DEBUG-REPLYCHK] subj='{subject}' conv_id={mail.ConversationID} "
                            f"topic='{mail.ConversationTopic}' check_after={sent_on:%Y-%m-%d %H:%M}")

                    # 수신자별로 회신 여부 업데이트
                    check_and_update_replies(app, mail, state, verbose=verbose)

                    # ✅ 반드시 저장
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

# ---- END OF ORIGINAL SCRIPT FUNCTIONS ----


# ---- Main Application Logic for Background Execution ----
exit_event = threading.Event()

def mark_reply_received(original_item, reply_item, state):
    try:
        sender = reply_item.SenderEmailAddress.lower()
    except Exception:
        return

    for r in original_item.Recipients:
        if r.Address and r.Address.lower() == sender:
            key = make_state_key(original_item.EntryID, r.Address)
            if key not in state:
                state[key] = {}
            state[key]["replied"] = True
            log(f"[REPLY-DETECTED] {r.Address} replied, stop reminders for this recipient")
            return


def start_mail_check_loop(args):
    """The main loop that runs in a background thread."""
    st = load_state()
    while not exit_event.is_set():
        try:
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

root = tk.Tk()
root.withdraw()

def open_settings_window(icon, item):
    create_and_show_gui()

def exit_action(icon, item):
    """Callback to gracefully exit the application."""
    log("[INFO] Exit requested. Shutting down.")
    exit_event.set()
    icon.stop()

def make_state_key(entry_id, recipient_addr):
    return f"{entry_id}|{recipient_addr.lower()}"

if __name__ == "__main__":
    check_single_instance()
    show_startup_notification()
    cfg = load_body_map()
    VERBOSE = cfg.get("verbose", False)

    if cfg.get("auto_start", False):
        register_startup_reg(APP_RUN_NAME)
    else:
        unregister_startup_reg(APP_RUN_NAME)
    
    p = argparse.ArgumentParser(description="Automated Outlook Mail Reminder System")
    p.add_argument("--precheck-epsilon-sec", type=int, default=10)
    p.add_argument("--loop-budget-sec", type=int, default=45)
    p.add_argument("--max-age-hours", type=float, default=0.0)
    p.add_argument("--skip-if-newer-outgoing", action="store_true")
    p.add_argument("--lookback-days", type=int, default=60)
    p.add_argument("--show-settings", action="store_true", help="Show settings window and exit")
    p.add_argument("--interval-min", type=int, default=1, help="Interval in minutes for background check.")
    p.add_argument("--verbose", action="store_true", help="Enable verbose logging (override config)")  # 기본값 제거)
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--force-send", action="store_true")
    p.add_argument("--skip-reply-check", action="store_true")
    p.add_argument("--include-self", action="store_true")
    p.add_argument("--due-from-last", action="store_true")
    p.add_argument("--include-deleted", action="store_true")
    p.add_argument("--reply-mode", choices=["hdr-only", "hdr-first", "conv-first"], default="conv-first")
    args = p.parse_args()

    if args.verbose:
        VERBOSE = True

    log(f"[INFO] Verbose mode = {VERBOSE}")

    mail_thread = threading.Thread(target=start_mail_check_loop, args=(args,))
    mail_thread.daemon = True
    mail_thread.start()
    log("[INFO] Mail check background thread started.")

    import sys, os
    from PIL import Image

    def resource_path(filename):
        """PyInstaller로 빌드된 exe 내부와 개발 환경 모두에서 리소스 파일 경로 찾기"""
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, filename)
        return os.path.join(os.path.abspath("."), filename)

    try:
        image = Image.open(resource_path("icon.png"))
    except FileNotFoundError:
        log("[ERROR] 'icon.png' not found. A default icon will be used.")
        image = Image.new('RGB', (64, 64), color='blue')

    menu = (
        item('설정 열기', open_settings_window),
        item('종료', exit_action)
    )
    tray_icon = icon('AutoMailSystem', image, "자동 메일 리마인더", menu)

    log("[INFO] System tray icon is running.")
    tray_icon.run()
