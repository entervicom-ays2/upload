"""
YouTube Auto Upload Script v2.0
- Auto-detect CHANNEL_CODE từ file .exe trong thư mục cha
- Auto-update từ URL online
- Cache + Retry cho Google Sheets API (fix quota 429)
"""

import os, sys, logging, time, random, shutil, ctypes, hashlib
from types import SimpleNamespace
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pyautogui
import pyperclip
import requests

# ================== VERSION & AUTO-UPDATE ==================
VERSION = "2.0.3"

# Cấu hình GitHub repo để auto-update
GITHUB_USER = "entervicom-ays2"      # Điền username GitHub, ví dụ: "criggerbrannon-hash"
GITHUB_REPO = "upload"      # Điền tên repo, ví dụ: "upload"
GITHUB_BRANCH = "main"

# Files/folders không được ghi đè khi update (giữ nguyên của máy local)
UPDATE_EXCLUDE = ["creds.json", "upload.log"]

UPDATE_CHECK_INTERVAL = 3600  # Kiểm tra update mỗi 1 giờ

# ================== LOGGING ==================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("upload.log", encoding="utf-8")
    ]
)

# ================== DPI AWARE ==================
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

pyautogui.FAILSAFE = False

# ================== AUTO-DETECT CONFIG ==================
def detect_config():
    """
    Tự động detect CHANNEL_CODE và các đường dẫn dựa vào cấu trúc thư mục.
    
    Cấu trúc mong đợi:
    C:\\Users\\{user}\\Documents\\{CHANNEL_CODE}\\
    ├── upload-{SPREADSHEET}/    <- thư mục chứa script này
    │   ├── main.py
    │   ├── icon/
    │   └── creds.json
    ├── {CHANNEL_CODE}.exe       <- trình duyệt
    └── ...
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    
    # Tìm file .exe trong thư mục cha
    exe_files = [f for f in os.listdir(parent_dir) if f.lower().endswith('.exe')]
    if not exe_files:
        raise RuntimeError(f"Không tìm thấy file .exe trong {parent_dir}")
    
    # Lấy tên file đầu tiên (không có .exe) làm CHANNEL_CODE
    channel_code = os.path.splitext(exe_files[0])[0]
    browser_exe = os.path.join(parent_dir, exe_files[0])
    
    # Detect SPREADSHEET_NAME từ tên thư mục script (upload-AYS2 -> AYS2)
    folder_name = os.path.basename(script_dir)
    spreadsheet_name = folder_name.replace("upload-", "") if folder_name.startswith("upload-") else "AYS2"
    
    # Đường dẫn DONE
    user_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    local_done = os.path.join(user_desktop, "DONE")
    server_done = r"\\tsclient\D\AUTO\done"
    
    # Tạo thư mục DONE nếu chưa có
    os.makedirs(local_done, exist_ok=True)
    
    config = {
        "CHANNEL_CODE": channel_code,
        "RUN_BROWSER_EXE": browser_exe,
        "SPREADSHEET_NAME": spreadsheet_name,
        "LOCAL_DONE_ROOT": local_done,
        "SERVER_DONE_ROOT": server_done,
        "SCRIPT_DIR": script_dir,
        "ICON_DIR": os.path.join(script_dir, "icon"),
        "CREDENTIAL_PATH": os.path.join(script_dir, "creds.json"),
    }
    
    logging.info(f"📋 Config detected:")
    logging.info(f"   CHANNEL_CODE: {config['CHANNEL_CODE']}")
    logging.info(f"   BROWSER: {config['RUN_BROWSER_EXE']}")
    logging.info(f"   SPREADSHEET: {config['SPREADSHEET_NAME']}")
    logging.info(f"   LOCAL_DONE: {config['LOCAL_DONE_ROOT']}")
    
    return config

# Load config
try:
    CFG = detect_config()
except Exception as e:
    logging.error(f"Lỗi detect config: {e}")
    sys.exit(1)

# ================== CONSTANTS ==================
INPUT_SHEET = "INPUT"
SOURCE_SHEET = "NGUON"
STATUS_OK = "EDIT XONG"
STATUS_COL = 48  # AV

# Column indices (zero-based)
IDX_TITLE_BB = 53
IDX_DESC_BC = 54
IDX_LINK_BD = 55
IDX_LINK_BE = 56
IDX_LINK_BF = 57
IDX_LINK_BG = 58
IDX_DATE_BI = 60
IDX_TIME_BJ = 61

UPLOAD_URL = "https://www.youtube.com/upload"
FOLDER_PATTERN = os.path.join(CFG["LOCAL_DONE_ROOT"], "{code}")

# Icon templates
ICON_DIR = CFG["ICON_DIR"]
TEMPLATES = {
    "SELECT_BTN": "chonfile.png",
    "DANHSACHPHAT": "danhsachphat.png",
    "DANGKY": "dangky.png",
    "NEXT_BTN": "tiep.png",
    "OPEN_READY": "open.png",
    "BUOC2": "buoc2.png",
    "CHON_ENDSCREEN": "chonmanhinhketthuc.png",
    "STEP2_THEM": "them.png",
    "DONE": "xong.png",
    "SAVE": "luu.png",
    "ENDSCREEN": "manhinhketthuc.png",
    "CHONVIDEO_CUTHE": "chonmotvideocuthe.png",
    "THE1": "the1.png",
    "HENLICH": "henlich.png",
    "SCHEDULE_PUBLISH": "lenlich.png",
    "DAHIEU": "dahieu.png",
    "FILENAME": "filename.png",
    "TAITEPLEN": "taiteplen.png",
    "KETTHUC_OK": "ketthucok.png",
    "THE": "the.png",
    "TAGVIDEO": "tagvideo.png",
    "TIME": "time.png",
    "TIEPTUC": "tieptuc.png",
    "CHEDO_HIEN_THI": "chedohienthi.png",
    "THUNNGHIEM": "thunghiem.png",
}

def icon(name):
    return os.path.join(ICON_DIR, TEMPLATES.get(name, name))

# ================== RANDOM PARAMS ==================
RANDOM = SimpleNamespace(
    tiny=(0.5, 0.9),
    small=(1.2, 2.0),
    medium=(2.5, 4.0),
    long=(5.0, 8.0),
    mouse_move=(0.25, 0.45),
    retry_screen_interval=(1.2, 2.0),
    browser_launch_wait_sec=(12, 20),
    click_timeout_sec=(120, 180),
    click_confidence=(0.70, 0.90),
    step2_load_timeout_sec=(150, 240),
)

def r(a, b):
    return random.uniform(a, b)

def rsleep(bucket="small"):
    lo, hi = getattr(RANDOM, bucket)
    time.sleep(r(lo, hi))

# ================== AUTO-UPDATE ==================
_last_update_check = 0

def get_remote_version():
    """Lấy version từ file main.py trên GitHub."""
    if not GITHUB_USER or not GITHUB_REPO:
        return None
    
    try:
        url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/main.py"
        resp = requests.get(url, timeout=10)
        if resp.status_code != 200:
            return None
        
        for line in resp.text.split('\n')[:30]:
            if line.startswith('VERSION = '):
                return line.split('"')[1]
        return None
    except Exception:
        return None

def download_and_extract_repo():
    """Tải ZIP repo từ GitHub và giải nén vào thư mục script."""
    import zipfile
    import io
    
    zip_url = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/archive/refs/heads/{GITHUB_BRANCH}.zip"
    logging.info(f"📥 Tải repo từ: {zip_url}")
    
    try:
        resp = requests.get(zip_url, timeout=60)
        if resp.status_code != 200:
            logging.error(f"Không tải được ZIP: HTTP {resp.status_code}")
            return False
        
        # Giải nén vào bộ nhớ
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            # Tên thư mục gốc trong ZIP (thường là {repo}-{branch})
            root_folder = zf.namelist()[0].split('/')[0]
            
            script_dir = CFG["SCRIPT_DIR"]
            
            for member in zf.namelist():
                # Bỏ qua thư mục gốc
                if member == root_folder + '/':
                    continue
                
                # Đường dẫn tương đối (bỏ thư mục gốc)
                rel_path = member[len(root_folder) + 1:]
                if not rel_path:
                    continue
                
                # Kiểm tra có trong danh sách exclude không
                skip = False
                for exclude in UPDATE_EXCLUDE:
                    if rel_path == exclude or rel_path.startswith(exclude + '/'):
                        skip = True
                        break
                
                if skip:
                    logging.info(f"⏭️ Bỏ qua (exclude): {rel_path}")
                    continue
                
                target_path = os.path.join(script_dir, rel_path)
                
                # Nếu là thư mục
                if member.endswith('/'):
                    os.makedirs(target_path, exist_ok=True)
                else:
                    # Tạo thư mục cha nếu chưa có
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    
                    # Ghi file
                    with open(target_path, 'wb') as f:
                        f.write(zf.read(member))
                    logging.info(f"✅ Cập nhật: {rel_path}")
        
        return True
        
    except Exception as e:
        logging.error(f"Lỗi khi tải/giải nén repo: {e}")
        return False

def check_for_updates():
    """Kiểm tra và tự động cập nhật nếu có version mới."""
    global _last_update_check
    
    if not GITHUB_USER or not GITHUB_REPO:
        logging.debug("Chưa cấu hình GitHub repo, bỏ qua check update")
        return False
    
    now = time.time()
    if now - _last_update_check < UPDATE_CHECK_INTERVAL:
        return False
    _last_update_check = now
    
    try:
        logging.info("🔍 Kiểm tra cập nhật...")
        
        remote_version = get_remote_version()
        if not remote_version:
            logging.warning("Không lấy được version từ GitHub")
            return False
        
        logging.info(f"📋 Version hiện tại: {VERSION}, Version mới nhất: {remote_version}")
        
        if remote_version == VERSION:
            logging.info(f"✅ Đang dùng version mới nhất: {VERSION}")
            return False
        
        # Có version mới
        logging.info(f"📥 Phát hiện version mới: {VERSION} → {remote_version}")
        
        # Backup file main.py cũ
        script_path = os.path.abspath(__file__)
        backup_path = script_path + ".backup"
        try:
            shutil.copy(script_path, backup_path)
            logging.info(f"💾 Đã backup: {backup_path}")
        except Exception:
            pass
        
        # Tải và giải nén repo mới
        if download_and_extract_repo():
            logging.info("✅ Cập nhật thành công! Khởi động lại script...")
            time.sleep(2)
            os.execv(sys.executable, [sys.executable] + sys.argv)
            return True
        else:
            logging.error("❌ Cập nhật thất bại")
            return False
        
    except Exception as e:
        logging.warning(f"Lỗi kiểm tra update: {e}")
        return False

# ================== GOOGLE SHEETS (với Cache + Retry) ==================
_CACHE = {}
_CACHE_TTL = 120  # Cache 2 phút

def retry_api_call(func, max_retries=5, base_delay=10):
    """Retry với exponential backoff khi gặp lỗi 429."""
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            err_str = str(e)
            if '429' in err_str or 'Quota' in err_str:
                delay = base_delay * (2 ** attempt)
                logging.warning(f"⏳ Quota exceeded, đợi {delay}s (lần {attempt+1}/{max_retries})...")
                time.sleep(delay)
            else:
                raise e
    raise Exception(f"Hết {max_retries} lần retry")

def cached_get_all_values(ws, cache_key):
    """Lấy dữ liệu từ cache nếu còn hạn."""
    now = time.time()
    if cache_key in _CACHE:
        data, ts = _CACHE[cache_key]
        if now - ts < _CACHE_TTL:
            logging.debug(f"📦 Cache hit: {cache_key}")
            return data
    
    data = retry_api_call(ws.get_all_values)
    _CACHE[cache_key] = (data, now)
    return data

def invalidate_cache(cache_key=None):
    """Xóa cache."""
    global _CACHE
    if cache_key:
        _CACHE.pop(cache_key, None)
    else:
        _CACHE.clear()

def gs_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CFG["CREDENTIAL_PATH"], scope)
    return gspread.authorize(creds)

def get_rows(client, sheet_name):
    ws = client.open(CFG["SPREADSHEET_NAME"]).worksheet(sheet_name)
    return cached_get_all_values(ws, f"rows_{sheet_name}")

def update_source_status(client, code, status="ĐÃ ĐĂNG"):
    """Cập nhật trạng thái với cache + retry."""
    try:
        ws = client.open(CFG["SPREADSHEET_NAME"]).worksheet(SOURCE_SHEET)
        rows = cached_get_all_values(ws, f"source_{SOURCE_SHEET}")
        
        for i, row in enumerate(rows[1:], start=2):
            if len(row) > 12 and norm(row[6]) == code:
                retry_api_call(lambda: ws.update_cell(i, 13, status))
                logging.info(f"✅ Đã cập nhật '{status}' cho mã {code}")
                invalidate_cache(f"source_{SOURCE_SHEET}")
                return True
        
        logging.warning(f"Không tìm thấy mã {code} trong sheet {SOURCE_SHEET}")
        return False
    except Exception as e:
        logging.error(f"Lỗi update status: {e}")
        return False

# ================== HELPERS ==================
def _get_scale():
    sw, sh = pyautogui.size()
    iw, ih = pyautogui.screenshot().size
    return iw / (sw or 1), ih / (sh or 1)

def _to_logical(x, y):
    sx, sy = _get_scale()
    return int(x / sx), int(y / sy)

def norm(s):
    return s.strip() if isinstance(s, str) else None

def click_once(x, y):
    lx, ly = _to_logical(x, y)
    pyautogui.moveTo(lx, ly, duration=r(*RANDOM.mouse_move))
    pyautogui.click(lx, ly)

def move_click(x, y):
    click_once(x, y)

def paste_text(text):
    if text is None:
        return
    pyperclip.copy(text)
    rsleep("tiny")
    pyautogui.hotkey('ctrl', 'v')
    rsleep("tiny")

def _parse_date(s):
    for f in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), f).date()
        except:
            pass
    return None

def _parse_time(s):
    for f in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s.strip(), f).time()
        except:
            pass
    return None

# ================== BROWSER CONTROL ==================
def open_run_and_execute(cmd):
    pyautogui.hotkey('win', 'r')
    rsleep("small")
    try:
        pyperclip.copy(cmd)
        rsleep("tiny")
        pyautogui.hotkey('ctrl', 'v')
        rsleep("tiny")
    except Exception as e:
        logging.warning(f"Paste lỗi: {e}")
        pyautogui.typewrite(cmd, interval=0.02)
    pyautogui.press('enter')
    rsleep("medium")

def close_browsers():
    logging.info("🧹 Đóng browsers...")
    open_run_and_execute('cmd /c del /q /f /s "%temp%\\*.*" >nul 2>&1')
    rsleep("small")
    
    exebase = os.path.splitext(os.path.basename(CFG["RUN_BROWSER_EXE"]))[0]
    exename = os.path.basename(CFG["RUN_BROWSER_EXE"])
    
    # PowerShell close
    ps_close = f"$names=@('chrome','msedge','firefox','{exebase}');$procs=Get-Process -EA 0|?{{$names -contains $_.ProcessName}};foreach($p in $procs){{if($p.MainWindowHandle -ne 0){{$null=$p.CloseMainWindow()}}}}"
    open_run_and_execute(f'powershell -NoProfile -WindowStyle Hidden -Command "{ps_close}"')
    rsleep("small")
    
    # Force kill
    skill = f'cmd /c taskkill /F /IM chrome.exe /T 2>nul & taskkill /F /IM msedge.exe /T 2>nul & taskkill /F /IM firefox.exe /T 2>nul & taskkill /F /IM "{exename}" /T 2>nul'
    open_run_and_execute(skill)
    rsleep("small")

# ================== IMAGE RECOGNITION ==================
def wait_image(img_path, timeout_sec=30, confidence=0.85):
    """Chờ ảnh xuất hiện, trả về vị trí hoặc None."""
    logging.info(f"Chờ ảnh: {os.path.basename(img_path)}...")
    end = time.time() + timeout_sec
    
    while time.time() < end:
        try:
            pos = pyautogui.locateCenterOnScreen(img_path, confidence=confidence)
            if pos:
                logging.info(f"✓ Thấy ảnh tại ({pos.x}, {pos.y})")
                return pos
        except Exception:
            pass
        time.sleep(r(*RANDOM.retry_screen_interval))
    
    logging.warning(f"✗ Không thấy ảnh: {os.path.basename(img_path)}")
    return None

def wait_and_click_image(img_path, timeout_sec=30, confidence=0.85):
    """Chờ ảnh và click với giảm dần confidence."""
    logging.info(f"Chờ + click: {os.path.basename(img_path)}...")
    end = time.time() + timeout_sec
    levels = [confidence, 0.8, 0.75, 0.7, 0.65, 0.6]
    
    while time.time() < end:
        for conf in levels:
            try:
                pos = pyautogui.locateCenterOnScreen(img_path, confidence=conf)
                if pos:
                    click_once(pos.x, pos.y)
                    logging.info(f"✓ Click ảnh tại ({pos.x}, {pos.y}) conf={conf:.2f}")
                    return True
            except Exception:
                pass
        time.sleep(r(*RANDOM.retry_screen_interval))
    
    logging.warning(f"✗ Không click được: {os.path.basename(img_path)}")
    return False

# ================== FILE HANDLING ==================
IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp"}

def has_required_files(dir_path):
    """Kiểm tra thư mục có đủ mp4+srt+ảnh."""
    if not os.path.isdir(dir_path):
        return False
    names = os.listdir(dir_path)
    has_mp4 = any(n.lower().endswith(".mp4") for n in names)
    has_srt = any(n.lower().endswith(".srt") for n in names)
    has_img = any(os.path.splitext(n)[1].lower() in IMG_EXTS for n in names)
    return has_mp4 and has_srt and has_img

def get_required_stats(dir_path):
    """Trả về (count, bytes) của các file bắt buộc."""
    if not os.path.isdir(dir_path):
        return (0, 0)
    total, count = 0, 0
    for root, _, files in os.walk(dir_path):
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext in (".mp4", ".srt") or ext in IMG_EXTS:
                try:
                    total += os.path.getsize(os.path.join(root, name))
                    count += 1
                except Exception:
                    pass
    return (count, total)

def ensure_local_folder(code, delete_server=True):
    """Đảm bảo thư mục local có đủ file."""
    local_folder = os.path.join(CFG["LOCAL_DONE_ROOT"], code)
    server_folder = os.path.join(CFG["SERVER_DONE_ROOT"], code)
    
    local_ok = os.path.isdir(local_folder) and has_required_files(local_folder)
    server_ok = has_required_files(server_folder)
    
    if local_ok:
        if server_ok:
            lc, sc = get_required_stats(local_folder), get_required_stats(server_folder)
            if lc == sc:
                logging.info(f"✅ Local đủ: {local_folder}")
                return True
            logging.info(f"♻️ Local khác server → refresh")
        else:
            logging.info(f"✅ Local đủ, server không có")
            return True
    
    if not server_ok:
        logging.error(f"❌ Server thiếu: {server_folder}")
        return False
    
    try:
        if os.path.exists(local_folder):
            shutil.rmtree(local_folder, ignore_errors=True)
        shutil.copytree(server_folder, local_folder)
        logging.info(f"📥 Đã copy: {server_folder} → {local_folder}")
    except Exception as e:
        logging.error(f"❌ Lỗi copy: {e}")
        return False
    
    if not has_required_files(local_folder):
        logging.error(f"❌ Sau copy vẫn thiếu: {local_folder}")
        return False
    
    if delete_server:
        try:
            shutil.rmtree(server_folder)
            logging.info(f"🗑️ Đã xóa server: {server_folder}")
        except Exception as e:
            logging.warning(f"Không xóa được server: {e}")
    
    return True

def cleanup_posted_codes():
    """Xóa thư mục local của mã đã đăng."""
    logging.info("🧹 Dọn mã đã đăng...")
    try:
        client = gs_client()
        ws = client.open(CFG["SPREADSHEET_NAME"]).worksheet(INPUT_SHEET)
        rows = cached_get_all_values(ws, f"cleanup_{INPUT_SHEET}")
        
        for row in rows[1:]:
            code = row[0].strip() if len(row) > 0 else ""
            status = row[STATUS_COL-1].strip() if len(row) >= STATUS_COL else ""
            if code and status.upper() == "ĐÃ ĐĂNG":
                folder = os.path.join(CFG["LOCAL_DONE_ROOT"], code)
                if os.path.isdir(folder):
                    try:
                        shutil.rmtree(folder)
                        logging.info(f"🗑️ Đã xóa: {folder}")
                    except Exception as e:
                        logging.warning(f"Không xóa được {folder}: {e}")
    except Exception as e:
        logging.warning(f"Lỗi cleanup: {e}")

def find_row_by_code(rows, code):
    for row in rows[1:]:
        if row and len(row) > 0 and norm(row[0]) == code:
            return row
    return None

def get_all_ready_codes(rows):
    """Lấy mã cần đăng hôm nay."""
    now = datetime.now()
    out = []
    for row in rows[1:]:
        if len(row) > 61 and norm(row[34]) == CFG["CHANNEL_CODE"] and norm(row[47]) == STATUS_OK:
            d = _parse_date(norm(row[60]) or "")
            t = _parse_time(norm(row[61]) or "")
            if d and t:
                target = datetime.combine(d, t)
                if d == now.date() and target > now:
                    code = norm(row[0])
                    if code:
                        out.append(code)
    return out

def get_tomorrow_codes(rows):
    """Lấy mã ngày mai để pre-stage."""
    tomorrow = datetime.now().date() + timedelta(days=1)
    out = []
    for row in rows[1:]:
        if len(row) > 61 and norm(row[34]) == CFG["CHANNEL_CODE"] and norm(row[47]) == STATUS_OK:
            d = _parse_date(norm(row[60]) or "")
            if d and d == tomorrow:
                code = norm(row[0])
                if code:
                    out.append(code)
    return out

# ================== FILE DIALOGS ==================
def file_dialog_select_first_mp4(target_folder):
    rsleep("long")
    
    if wait_and_click_image(icon("FILENAME"), timeout_sec=60, confidence=0.75):
        rsleep("medium")
    
    pyautogui.hotkey('ctrl', 'l'); rsleep("tiny")
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text(target_folder)
    pyautogui.press('enter'); rsleep("medium")
    
    pyautogui.keyDown('alt'); pyautogui.press('n'); pyautogui.keyUp('alt'); rsleep("tiny")
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text('*.mp4')
    pyautogui.press('enter'); rsleep("long")
    
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.press('space'); rsleep("tiny")
    
    for _ in range(2):
        pyautogui.press('tab'); rsleep("small")
    pyautogui.press('enter'); rsleep("long")

def file_dialog_select_thumbnail():
    rsleep("medium")
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.press('space'); rsleep("small")
    for _ in range(4):
        pyautogui.press('tab'); rsleep("tiny")
    pyautogui.press('enter'); rsleep("long")

def file_dialog_select_srt():
    if wait_and_click_image(icon("FILENAME"), timeout_sec=60, confidence=0.75):
        rsleep("small")
    
    paste_text('*.srt'); rsleep("tiny")
    pyautogui.press('enter'); rsleep("small")
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.hotkey('shift', 'tab'); rsleep("tiny")
    pyautogui.press('space'); rsleep("medium")
    for _ in range(4):
        pyautogui.press('tab'); rsleep("tiny")
    pyautogui.press('enter'); rsleep("long")

# ================== UPLOAD PROGRESS CHECK ==================
def wait_for_upload_complete(timeout_minutes=10):
    """
    Chờ video upload xong trước khi tiếp tục.
    Đơn giản: chờ cứng timeout_minutes phút cho an toàn.
    """
    logging.info(f"⏳ Chờ {timeout_minutes} phút để đảm bảo video upload xong...")
    
    for minute in range(timeout_minutes):
        remaining = timeout_minutes - minute
        logging.info(f"⏳ Còn {remaining} phút...")
        time.sleep(60)  # Chờ 1 phút
    
    logging.info(f"✅ Đã chờ đủ {timeout_minutes} phút, sẵn sàng tiếp tục")
    return True

def safe_fallback_step2():
    """
    Fallback an toàn khi Step 2 lỗi:
    1. Chờ cứng 10 phút để đảm bảo upload xong
    2. F5 refresh
    3. Enter để confirm dialog (nếu có)
    """
    logging.warning("⚠️ Step 2 lỗi - Bắt đầu fallback an toàn...")
    
    # Chờ cứng 10 phút
    wait_for_upload_complete(timeout_minutes=10)
    
    # F5 refresh
    try:
        logging.info("🔄 F5 để refresh trang...")
        pyautogui.press('f5')
        rsleep("long")  # Chờ trang load
        
        # Enter để đóng dialog confirm (nếu có)
        pyautogui.press('enter')
        rsleep("medium")
        
        # Chờ thêm cho trang ổn định
        time.sleep(5)
        
        logging.info("✅ Đã F5 + Enter, sẵn sàng tiếp tục")
        return True
        
    except Exception as e:
        logging.error(f"Lỗi khi fallback: {e}")
        return False

# ================== UPLOAD FLOW ==================
def press(key, n=1, bucket="tiny"):
    for _ in range(n):
        pyautogui.press(key); rsleep(bucket)

def handle_metadata_flow(active_row):
    """Nhập metadata: tiêu đề, mô tả, thumbnail, playlist."""
    title = norm(active_row[IDX_TITLE_BB]) if len(active_row) > IDX_TITLE_BB else ""
    desc = norm(active_row[IDX_DESC_BC]) if len(active_row) > IDX_DESC_BC else ""
    
    TIMEOUT = int(r(*RANDOM.click_timeout_sec))
    CONF = r(*RANDOM.click_confidence)
    
    logging.info(f"Nhập TIÊU ĐỀ: {title[:50]}...")
    rsleep("long")
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text(title or "")
    
    # Check UI thử nghiệm
    try:
        test_pos = pyautogui.locateCenterOnScreen(icon("THUNNGHIEM"), confidence=0.80)
        tab_count = 3 if test_pos else 2
    except Exception:
        tab_count = 2
    
    press('tab', tab_count, "tiny")
    rsleep("small")
    
    logging.info("Nhập MÔ TẢ...")
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text(desc or "")
    
    pyautogui.press('enter'); rsleep("tiny")
    press('tab', 2, "tiny")
    rsleep("small")
    
    # Cuộn xuống + chọn thumbnail
    press('end', 2, "small")
    rsleep("medium")
    pyautogui.press('enter'); rsleep("small")
    
    if wait_image(icon("OPEN_READY"), timeout_sec=TIMEOUT, confidence=CONF):
        file_dialog_select_thumbnail()
    else:
        logging.error("Không thấy hộp thoại Open thumbnail")
        return
    
    # Chọn playlist
    pos_dsp = wait_image(icon("DANHSACHPHAT"), timeout_sec=TIMEOUT, confidence=CONF)
    if pos_dsp:
        move_click(pos_dsp.x, pos_dsp.y); rsleep("small")
        pyautogui.press('tab'); rsleep("tiny")
        pyautogui.press('enter'); rsleep("small")
        press('tab', 2, "tiny")
        pyautogui.press('enter'); rsleep("small")
    
    # Click Tiếp
    pos = wait_image(icon("NEXT_BTN"), timeout_sec=TIMEOUT, confidence=CONF)
    if pos:
        click_once(pos.x, pos.y)
    else:
        logging.warning("Không thấy nút Tiếp")

def handle_step2_flow(active_row):
    """Step 2: phụ đề, end screen, thẻ."""
    TIMEOUT = int(r(*RANDOM.click_timeout_sec))
    CONF = r(*RANDOM.click_confidence)
    STEP2_TIMEOUT = int(r(*RANDOM.step2_load_timeout_sec))
    
    # Vào Bước 2
    logging.info("Vào Bước 2...")
    pos_buoc2 = wait_image(icon("BUOC2"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
    if not pos_buoc2:
        pos_buoc2 = wait_image(icon("STEP2_THEM"), timeout_sec=30, confidence=CONF)
        if not pos_buoc2:
            logging.error("Không vào được Bước 2")
            return False
    
    # Click và chờ taiteplen.png
    for attempt in range(5):
        move_click(pos_buoc2.x, pos_buoc2.y); rsleep("small")
        press('tab', 4, "tiny")
        pyautogui.press('enter'); rsleep("small")
        
        if wait_image(icon("TAITEPLEN"), timeout_sec=10, confidence=CONF):
            break
        
        pos_buoc2 = wait_image(icon("BUOC2"), timeout_sec=15, confidence=CONF) or \
                    wait_image(icon("STEP2_THEM"), timeout_sec=5, confidence=CONF)
        if not pos_buoc2:
            return False
    else:
        return False
    
    # Click taiteplen với retry
    time.sleep(15)
    for attempt in range(3):
        for conf in [CONF, 0.80, 0.75, 0.70]:
            try:
                pos = pyautogui.locateCenterOnScreen(icon("TAITEPLEN"), confidence=conf)
                if pos:
                    move_click(pos.x, pos.y)
                    break
            except Exception:
                pass
        
        time.sleep(15)
        try:
            if pyautogui.locateCenterOnScreen(icon("TIEPTUC"), confidence=0.70):
                break
        except Exception:
            pass
    else:
        return False
    
    # Click tieptuc
    if not wait_and_click_image(icon("TIEPTUC"), timeout_sec=STEP2_TIMEOUT, confidence=CONF):
        press('tab', 3, "tiny")
        pyautogui.press('enter'); rsleep("long")
    else:
        rsleep("long")
    
    # Chọn SRT
    if not wait_image(icon("OPEN_READY"), timeout_sec=STEP2_TIMEOUT, confidence=CONF):
        return False
    file_dialog_select_srt()
    
    # Đợi xong
    pos_done = wait_image(icon("DONE"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
    if not pos_done:
        return False
    rsleep("medium")
    move_click(pos_done.x, pos_done.y); rsleep("medium")
    
    # End screen
    if not wait_image(icon("ENDSCREEN"), timeout_sec=STEP2_TIMEOUT, confidence=CONF):
        return False
    
    press('tab', 2, "tiny")
    pyautogui.press('enter'); rsleep("medium")
    
    rsleep("medium")
    if not wait_and_click_image(icon("CHON_ENDSCREEN"), timeout_sec=STEP2_TIMEOUT, confidence=CONF):
        return False
    
    press('tab', 3, "tiny")
    press('enter', 2, "small")  # Video 1
    press('enter', 2, "small")  # Video 2
    press('enter', 1, "small")
    pyautogui.press('d'); rsleep("tiny")
    press('enter', 1, "small")
    press('tab', 3, "tiny")
    press('enter', 1, "small")
    press('enter', 1, "small")
    
    pos_dangky = wait_image(icon("DANGKY"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
    if pos_dangky:
        move_click(pos_dangky.x, pos_dangky.y); rsleep("small")
    
    # Lưu end screen
    pos_save = wait_image(icon("SAVE"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
    if not pos_save:
        return False
    move_click(pos_save.x, pos_save.y); rsleep("medium")
    
    # Thêm thẻ (Cards)
    if not wait_image(icon("KETTHUC_OK"), timeout_sec=STEP2_TIMEOUT, confidence=CONF):
        return False
    
    rsleep("small")
    press('tab', 1, "tiny")
    pyautogui.press('enter'); rsleep("small")
    
    def click_the_button():
        try:
            pyautogui.moveTo(10, 10, duration=0.1)
        except Exception:
            pass
        pos = wait_image(icon("THE"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
        if pos:
            lx, ly = _to_logical(pos.x, pos.y)
            pyautogui.moveTo(lx, ly, duration=0.15)
            pyautogui.click()
            rsleep("small")
            return True
        return False
    
    def click_the1_button():
        pos = wait_image(icon("THE1"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
        if pos:
            click_once(pos.x, pos.y)
            rsleep("tiny")
            return True
        return False
    
    # Thêm playlist card
    if click_the_button():
        press('tab', 4, "tiny")
        pyautogui.press('enter'); rsleep("small")
        rsleep("small")
        press('tab', 3, "tiny")
        pyautogui.press('enter'); rsleep("medium")
    
    # Thêm video cards (BD, BE, BF, BG)
    video_ok = []
    for idx, col_name in [(IDX_LINK_BD, "BD"), (IDX_LINK_BE, "BE"), (IDX_LINK_BF, "BF"), (IDX_LINK_BG, "BG")]:
        link = norm(active_row[idx]) if len(active_row) > idx else ""
        if not link:
            continue
        
        if not click_the1_button():
            continue
        
        rsleep("tiny")
        press('tab', 1, "tiny")
        pyautogui.press('enter'); rsleep("medium")
        rsleep("small")
        
        pos_choose = wait_image(icon("CHONVIDEO_CUTHE"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
        if not pos_choose:
            continue
        click_once(pos_choose.x, pos_choose.y)
        
        press('tab', 3, "tiny")
        paste_text(link); rsleep("small")
        
        pos_tag = wait_image(icon("TAGVIDEO"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
        if pos_tag:
            click_once(pos_tag.x, pos_tag.y)
            video_ok.append(col_name)
        rsleep("medium")
    
    if not video_ok:
        return False
    
    # Thêm timestamps
    for ts in ["30:00:00", "10:00:00", "15:00:00", "20:00:00", "25:00:00"]:
        if click_the_button():
            press('tab', 5, "tiny")
            paste_text(ts); rsleep("tiny")
            pyautogui.press('tab'); rsleep("small")
    
    # Lưu thẻ
    pos_save = wait_image(icon("SAVE"), timeout_sec=STEP2_TIMEOUT, confidence=CONF)
    if pos_save:
        move_click(pos_save.x, pos_save.y); rsleep("medium")
    
    logging.info("Step 2 hoàn thành")
    return True

def handle_step3_4_flow(active_row, client, code):
    """Step 3-4: hẹn lịch và đăng."""
    TIMEOUT = int(r(*RANDOM.click_timeout_sec))
    
    # Click Chế độ hiển thị
    if not wait_and_click_image(icon("CHEDO_HIEN_THI"), timeout_sec=TIMEOUT):
        return False
    rsleep("medium")
    
    # Click Hẹn lịch
    if not wait_and_click_image(icon("HENLICH"), timeout_sec=TIMEOUT):
        return False
    rsleep("medium")
    
    press('tab', 8, "tiny")
    pyautogui.press('enter'); rsleep("small")
    
    # Dán ngày
    date_val = norm(active_row[IDX_DATE_BI]) if len(active_row) > IDX_DATE_BI else ""
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text(date_val or "")
    pyautogui.press('enter'); rsleep("small")
    
    # Dán giờ
    time_val = norm(active_row[IDX_TIME_BJ]) if len(active_row) > IDX_TIME_BJ else ""
    pos_time = wait_image(icon("TIME"), timeout_sec=TIMEOUT)
    if not pos_time:
        return False
    move_click(pos_time.x, pos_time.y); rsleep("small")
    pyautogui.hotkey('ctrl', 'a'); rsleep("tiny")
    paste_text(time_val or "")
    pyautogui.press('enter'); rsleep("small")
    
    # Click Lên lịch
    pos_publish = wait_image(icon("SCHEDULE_PUBLISH"), timeout_sec=TIMEOUT)
    if not pos_publish:
        return False
    move_click(pos_publish.x, pos_publish.y); rsleep("medium")
    
    # Xử lý popup Đã hiểu
    try:
        if wait_and_click_image(icon("DAHIEU"), timeout_sec=15, confidence=0.80):
            logging.info("Đã click 'Đã hiểu'")
    except Exception:
        pass
    
    # Cập nhật trạng thái
    try:
        update_source_status(client, code, "ĐÃ ĐĂNG")
    except Exception as e:
        logging.warning(f"Lỗi update status: {e}")
    
    # Đợi 10 phút
    logging.info("⏳ Đợi 10 phút...")
    time.sleep(10 * 60)
    
    return True

# ================== MAIN ==================
def main():
    random.seed()
    
    # Kiểm tra update
    check_for_updates()
    
    # Dọn mã đã đăng
    cleanup_posted_codes()
    
    BROWSER_WAIT = int(r(*RANDOM.browser_launch_wait_sec))
    TIMEOUT = int(r(*RANDOM.click_timeout_sec))
    CONF = r(*RANDOM.click_confidence)
    
    client = gs_client()
    input_rows = get_rows(client, INPUT_SHEET)
    
    # Lấy mã cần đăng
    ready_codes = get_all_ready_codes(input_rows)
    if not ready_codes:
        logging.info(f"Không có mã cho {CFG['CHANNEL_CODE']} hôm nay")
        
        # Pre-stage ngày mai
        tomorrow = get_tomorrow_codes(input_rows)
        for c in tomorrow:
            try:
                ensure_local_folder(c)
            except Exception:
                pass
        return
    
    # Lọc mã có file
    ready_codes = [c for c in ready_codes if 
                   has_required_files(os.path.join(CFG["LOCAL_DONE_ROOT"], c)) or 
                   has_required_files(os.path.join(CFG["SERVER_DONE_ROOT"], c))]
    
    if not ready_codes:
        logging.info("Không còn mã hợp lệ")
        return
    
    logging.info(f"📋 Đăng {len(ready_codes)} mã: {ready_codes}")
    
    # Pre-stage
    for c in ready_codes:
        try:
            ensure_local_folder(c)
        except Exception:
            pass
    
    # Mở browser
    logging.info(f"🌐 Mở browser: {CFG['RUN_BROWSER_EXE']}")
    open_run_and_execute(CFG["RUN_BROWSER_EXE"])
    time.sleep(BROWSER_WAIT)
    
    # Upload từng mã
    first_time = True
    processed = set()
    
    for idx, code in enumerate(ready_codes, 1):
        if code in processed:
            continue
        
        logging.info(f"=== [{idx}/{len(ready_codes)}] CODE: {code} ===")
        
        active_row = find_row_by_code(input_rows, code)
        if not active_row:
            continue
        
        target_folder = FOLDER_PATTERN.format(code=code)
        if not ensure_local_folder(code):
            continue
        
        # Điều hướng
        if not first_time:
            pyautogui.hotkey('ctrl', 't'); rsleep("small")
        
        pyautogui.hotkey('ctrl', 'l'); rsleep("tiny")
        paste_text(UPLOAD_URL)
        pyautogui.press('enter'); rsleep("medium")
        
        # Phóng to
        try:
            pyautogui.keyDown('alt'); pyautogui.press('space'); pyautogui.keyUp('alt'); rsleep("tiny")
            pyautogui.press('x'); rsleep("small")
        except Exception:
            pass
        
        pyautogui.press('f5'); rsleep("medium")
        
        # Click Select files
        if not wait_and_click_image(icon("SELECT_BTN"), timeout_sec=TIMEOUT, confidence=CONF):
            pyautogui.press('f5'); rsleep("medium")
            if not wait_and_click_image(icon("SELECT_BTN"), timeout_sec=60, confidence=CONF):
                continue
        
        # Chọn video
        if not wait_image(icon("OPEN_READY"), timeout_sec=TIMEOUT, confidence=CONF):
            first_time = False
            continue
        
        file_dialog_select_first_mp4(target_folder)
        
        # Metadata
        if not wait_image(icon("NEXT_BTN"), timeout_sec=TIMEOUT, confidence=CONF):
            first_time = False
            continue
        
        handle_metadata_flow(active_row)
        
        # Step 2
        if not handle_step2_flow(active_row):
            # Fallback an toàn: chờ upload xong rồi mới F5
            safe_fallback_step2()
        
        # Step 3-4
        if handle_step3_4_flow(active_row, client, code):
            processed.add(code)
        
        first_time = False
    
    logging.info(f"✅ Hoàn thành {len(processed)}/{len(ready_codes)} mã")
    
    # Pre-stage ngày mai
    try:
        tomorrow = get_tomorrow_codes(input_rows)
        for c in tomorrow:
            ensure_local_folder(c)
    except Exception:
        pass

if __name__ == "__main__":
    while True:
        try:
            close_browsers()
            rsleep("small")
        except Exception as e:
            logging.warning(f"Lỗi đóng browser: {e}")
        
        try:
            main()
        except Exception as e:
            err = str(e)
            if '429' in err or 'Quota' in err:
                logging.error("🚫 Quota exceeded - đợi 5 phút...")
                time.sleep(5 * 60)
            else:
                logging.error(f"Lỗi main(): {e}")
        
        # Nghỉ 3 tiếng
        time.sleep(3 * 60 * 60)
