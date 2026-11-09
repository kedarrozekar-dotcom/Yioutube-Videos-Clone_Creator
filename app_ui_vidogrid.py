import asyncio
import os
import cv2
import sys
import time
import json
import logging
import math
import re
import html
from datetime import datetime
import subprocess
import threading
import pandas as pd
import openpyxl
import streamlit as st
import yt_dlp
from google import genai
from google.genai import types
from playwright.sync_api import sync_playwright

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

st.set_page_config(
    page_title="VidoGrid Muse Video Studio",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 26px; font-weight: 700; color: #1F497D; margin-bottom: 2px; }
    .sub-header { font-size: 14px; color: #555; margin-bottom: 18px; }
    .clip-card { background: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; padding: 12px; margin-bottom: 14px; }
    .clip-title { font-weight: 700; font-size: 14px; color: #1F497D; margin-bottom: 4px; }
    .clip-meta { font-size: 12px; color: #666; margin-bottom: 8px; font-weight: 500; }
    .prompt-box { background: #F8F9FA; border-left: 4px solid #1F497D; padding: 10px 14px; font-size: 13px; color: #222; border-radius: 4px; line-height: 1.4; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGS_DIR = os.path.join(BASE_DIR, "Logs")
WORKSPACE_DIR = os.path.join(BASE_DIR, "workspace_vidogrid")
CLIPS_DIR = os.path.join(WORKSPACE_DIR, "clips")
FRAMES_DIR = os.path.join(WORKSPACE_DIR, "frames")
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "final_videos")
PROMPTS_DIR = os.path.join(WORKSPACE_DIR, "prompts")

for d in [LOGS_DIR, CLIPS_DIR, FRAMES_DIR, OUTPUT_DIR, PROMPTS_DIR]:
    os.makedirs(d, exist_ok=True)

LOGIN_URL = os.getenv("VIDOGRID_LOGIN_URL", "https://vidogrid.ai/auth/login")
MUSE_URL = os.getenv("VIDOGRID_MUSE_URL", "https://vidogrid.ai/app/video-gen/muse")
VIDOGRID_USER = os.getenv("VIDOGRID_USER", "")
VIDOGRID_PASS = os.getenv("VIDOGRID_PASS", "")
PROMPT_MAX_CHARS = int(os.getenv("PROMPT_MAX_CHARS", "5000"))
INTRO_OUTRO_TRIM_SEC = int(os.getenv("INTRO_OUTRO_TRIM_SEC", "22"))
DEFAULT_DURATION_SEC = 10
GEMINI_MODELS = [m.strip() for m in os.getenv(
    "GEMINI_MODELS", "gemini-3.8-flash,gemini-3.6-flash,gemini-2.5-flash").split(",") if m.strip()]
MUSE_DURATION_SELECTOR = os.getenv("MUSE_DURATION_SELECTOR", "")


def parse_seconds(v, default=None):
    if v is None or str(v).strip() == "":
        return default
    t = str(v).strip().lower()
    if ":" in t:
        a, b = t.split(":")[-2:]
        a = int(re.sub(r"\D", "", a) or 0)
        b = int(re.sub(r"\D", "", b) or 0)
        return float(b if b > 59 else a * 60 + b)
    m = re.search(r"\d+(?:\.\d+)?", t)
    return float(m.group()) if m else default


def get_cell(ws, row, col_map, name, default=None):
    col = col_map.get(name.upper())
    return ws.cell(row=row, column=col).value if col else default


def trim_prompt(text, limit=None):
    limit = limit or PROMPT_MAX_CHARS
    text = str(text).strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    idx = max(cut.rfind(". "), cut.rfind("; "))
    if idx > limit - 600:
        cut = cut[:idx + 1]
    return cut.strip()


def build_final_prompt(p_text, cfg=None):
    if str(p_text).lstrip().startswith("STYLE:"):
        return trim_prompt(p_text)
    return trim_prompt(f"3D animated scene, {orientation_text(cfg)}, cinematic lighting, vivid colors, smooth motion. " + str(p_text))


def normalize_prompts(data):
    if isinstance(data, dict):
        data = next((v for v in data.values() if isinstance(v, list)), list(data.values()))
    out = []
    for item in data:
        if isinstance(item, dict):
            item = item.get("prompt") or next((v for v in item.values() if isinstance(v, str)), "")
        out.append(str(item))
    return out


def load_prompts_from_queue(wb, vid_id, seg_sec, orientation=None):
    if "Prompts" not in wb.sheetnames:
        return []
    ws = wb["Prompts"]
    hdr = {str(c.value).strip().upper(): i for i, c in enumerate(ws[1]) if c.value}
    if not {"VIDEO_ID", "CLIP_NO", "PROMPT"} <= set(hdr):
        return []
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[hdr["VIDEO_ID"]]).strip() != str(vid_id):
            continue
        if "DURATION_SEC" in hdr and row[hdr["DURATION_SEC"]] is not None:
            try:
                if abs(float(row[hdr["DURATION_SEC"]]) - float(seg_sec)) > 0.01:
                    continue
            except (TypeError, ValueError):
                continue
        if orientation and "ORIENTATION" in hdr and row[hdr["ORIENTATION"]] and \
                str(row[hdr["ORIENTATION"]]).strip().lower() != str(orientation).strip().lower():
            continue
        if row[hdr["PROMPT"]]:
            rows.append((int(row[hdr["CLIP_NO"]]), str(row[hdr["PROMPT"]])))
    return [t for _, t in sorted(rows)]


def set_muse_duration(page, seconds):
    s = int(seconds) if float(seconds).is_integer() else seconds
    try:
        if MUSE_DURATION_SELECTOR:
            loc = page.locator(MUSE_DURATION_SELECTOR.format(s=s)).first
            if loc.is_visible():
                loc.click(force=True)
                page.wait_for_timeout(400)
                return True
        pat = re.compile(rf"^\s*{re.escape(str(s))}\s*(s|sec|secs|seconds?)\s*$", re.I)
        for attempt in range(2):
            opt = page.get_by_text(pat).first
            if opt.is_visible():
                opt.click(force=True)
                page.wait_for_timeout(400)
                return True
            trigger = page.locator("[aria-label*='duration' i], button:has-text('Duration'), [role='combobox']").first
            if attempt == 0 and trigger.is_visible():
                trigger.click(force=True)
                page.wait_for_timeout(400)
            else:
                break
    except Exception:
        pass
    return False


PROVIDER_URLS = {
    "museai": os.getenv("VIDOGRID_MUSE_URL", "https://vidogrid.ai/app/video-gen/muse"),
    "muse": os.getenv("VIDOGRID_MUSE_URL", "https://vidogrid.ai/app/video-gen/muse"),
    "grok": os.getenv("VIDOGRID_GROK_URL", "https://vidogrid.ai/app/video-gen/grok?mode=create-new"),
    "gveo": os.getenv("VIDOGRID_VEO_URL", "https://vidogrid.ai/app/video-gen/veo"),
    "veo": os.getenv("VIDOGRID_VEO_URL", "https://vidogrid.ai/app/video-gen/veo"),
}


def _norm(x):
    return re.sub(r"[^a-z0-9]", "", str(x).lower())


def read_active_config(wb):
    ws = next((wb[n] for n in wb.sheetnames if _norm(n) in
               ("prerequisites", "prerequisite", "videoconfig", "videoconfiguration", "videoconfigurations")), None)
    if ws is None:
        return None
    hdr = {_norm(c.value): c.column for c in ws[1] if c.value}
    col = lambda prefix: next((v for k, v in hdr.items() if k.startswith(prefix)), None)
    cP, cO, cR, cD, cE = col("provider"), col("orientation"), col("resolution"), col("duration"), col("execute")
    if not (cP and cE):
        return None
    active = []
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, cE).value or "").strip().lower() in ("yes", "y", "true", "1"):
            g = lambda c: ws.cell(r, c).value if c else None
            active.append({
                "provider": str(g(cP) or "").strip(),
                "orientation": str(g(cO)).strip() if g(cO) else None,
                "resolution": str(g(cR)).strip() if g(cR) else None,
                "duration": parse_seconds(g(cD)),
            })
    return active[0] if active else None


def read_login_details(wb, provider=None):
    ws = next((wb[n] for n in wb.sheetnames if _norm(n) in
               ("logindetails", "logindetail", "logins", "logincredentials", "credentials")), None)
    if ws is None:
        return None
    hdr = {_norm(c.value): c.column for c in ws[1] if c.value}

    def col(*prefixes):
        for pre in prefixes:
            for k, v in hdr.items():
                if k.startswith(pre):
                    return v
        return None

    cE = col("execute")
    cU = col("username", "userid", "user", "email", "loginid", "login")
    cW = col("password", "pwd", "pass")
    cP = col("provider", "site", "website", "platform", "application", "app")
    if not (cU and cW):
        return None
    rows = []
    for r in range(2, ws.max_row + 1):
        user, pwd = ws.cell(r, cU).value, ws.cell(r, cW).value
        if not user or pwd is None:
            continue
        if cE and str(ws.cell(r, cE).value or "").strip().lower() not in ("yes", "y", "true", "1"):
            continue
        prov = str(ws.cell(r, cP).value or "") if cP else ""
        rows.append({"row": r, "user": str(user).strip(), "password": str(pwd).strip(), "provider": prov})
    if not rows:
        return None
    if provider:
        match = [x for x in rows if x["provider"] and (_norm(provider) in _norm(x["provider"]) or _norm(x["provider"]) in _norm(provider))]
        if match:
            rows = match
    return rows[0]


def mask_user(u):
    return (u[:2] + "***") if len(u) > 2 else "***"


def provider_start_url(provider):
    key = _norm(provider)
    if key in PROVIDER_URLS:
        return PROVIDER_URLS[key]
    slug = re.sub(r"[^a-z0-9]+", "-", str(provider).lower()).strip("-") or "muse"
    return f"https://vidogrid.ai/app/video-gen/{slug}"


def orientation_text(cfg):
    portrait = bool(cfg and cfg.get("orientation") and "port" in cfg["orientation"].lower())
    txt = "9:16 vertical portrait" if portrait else "16:9 landscape"
    if cfg and cfg.get("resolution"):
        txt += ", target resolution " + cfg["resolution"]
    return txt


def _click_option(page, pat, kind, hints):
    try:
        for attempt in range(2):
            opts = page.get_by_text(pat)
            for i in range(min(opts.count(), 6)):
                o = opts.nth(i)
                if o.is_visible():
                    o.click(force=True)
                    page.wait_for_timeout(400)
                    return True
            trig = page.locator(", ".join(f"[aria-label*='{h}' i], button:has-text('{h}')" for h in hints)).first
            if attempt == 0 and trig.count() and trig.is_visible():
                trig.click(force=True)
                page.wait_for_timeout(400)
            else:
                break
    except Exception:
        pass
    return False


def apply_page_settings(page, seconds, cfg):
    set_muse_duration(page, seconds)
    cfg = cfg or {}
    if cfg.get("orientation"):
        pat = re.compile(r"^\s*(portrait|9\s*:\s*16)\s*$", re.I) if "port" in cfg["orientation"].lower() else re.compile(r"^\s*(landscape|16\s*:\s*9)\s*$", re.I)
        _click_option(page, pat, "Orientation", ["orientation", "aspect", "ratio"])
    if cfg.get("resolution"):
        res = cfg["resolution"]
        alts = [re.escape(res).replace("\\ ", r"\s*")] + [rf"{d}\s*p?" for d in re.findall(r"\d{3,4}", res)]
        pat = re.compile(r"^\s*(" + "|".join(alts) + r")\s*$", re.I)
        _click_option(page, pat, "Resolution", ["resolution", "quality"])


GENERATE_BUTTON_SELECTOR = os.getenv("GENERATE_BUTTON_SELECTOR", "")


def _scroll_all_containers(page):
    try:
        page.evaluate("""() => {
            document.querySelectorAll('*').forEach(el => {
                const oy = getComputedStyle(el).overflowY;
                if (el.scrollHeight > el.clientHeight + 20 && oy !== 'hidden' && oy !== 'visible') el.scrollTop = el.scrollHeight;
            });
            window.scrollTo(0, document.body.scrollHeight);
        }""")
    except Exception:
        pass


def _generate_candidates(page):
    include = re.compile(r"generate", re.I)
    exclude = re.compile(r"generated|generations?|history|regenerate|again|grok", re.I)
    found = []
    if GENERATE_BUTTON_SELECTOR:
        for scope in page.frames:
            loc = scope.locator(GENERATE_BUTTON_SELECTOR)
            found += [loc.nth(i) for i in range(loc.count())]
        if found:
            return found
    for scope in page.frames:
        loc = scope.locator("button, [role='button'], input[type='submit']").filter(has_text=include)
        for i in range(min(loc.count(), 20)):
            el = loc.nth(i)
            try:
                txt = (el.inner_text(timeout=500) or "").strip()
            except Exception:
                txt = ""
            if txt and exclude.search(txt):
                continue
            found.append(el)
    return found


def click_generate(page, vid_id, idx):
    for attempt in range(1, 5):
        cands = _generate_candidates(page)
        for el in reversed(cands):
            try:
                el.scroll_into_view_if_needed(timeout=3000)
                page.wait_for_timeout(300)
                if not el.is_visible() or not el.is_enabled():
                    continue
                try:
                    el.click(timeout=4000)
                except Exception:
                    el.click(force=True, timeout=4000)
                return True
            except Exception:
                pass
        _scroll_all_containers(page)
        page.wait_for_timeout(1000)
    return True


STATUS_FILE = os.path.join(WORKSPACE_DIR, "live_status.json")
HISTORY_FILE = os.path.join(WORKSPACE_DIR, "execution_history.json")


def load_execution_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_to_history(vid_id, clips):
    if not clips:
        return
    history = load_execution_history()
    run_record = {
        "video_id": vid_id,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "clips": clips
    }
    history.insert(0, run_record)
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass


def read_live_status():
    if os.path.exists(STATUS_FILE):
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "is_running": False, "is_paused": False, "waiting_for_sign_in": False,
        "current_video_id": "", "current_stage": "Idle", "current_clip": 0, "total_clips": 0,
        "current_prompt": "", "clip_started_at": None, "batch_started_at": None,
        "stage_timings": {}, "current_stage_start": None, "current_log_file": None, "completed_clips": []
    }


def update_live_status(data):
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


if not os.path.exists(STATUS_FILE):
    update_live_status(read_live_status())

logger = logging.getLogger("VidoGridPipelineUI")
logger.setLevel(logging.INFO)


def setup_run_logger(log_file_path: str):
    logger.handlers.clear()
    f_handler = logging.FileHandler(log_file_path, mode="w", encoding="utf-8")
    f_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    c_handler = logging.StreamHandler(sys.stdout)
    c_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(f_handler)
    logger.addHandler(c_handler)


st.sidebar.title("⚙️ VidoGrid Muse Settings")
api_key_input = st.sidebar.text_input("Gemini API Key", value=os.getenv("GEMINI_API_KEY", ""), type="password")
headless_mode = st.sidebar.checkbox("Run Browser Headless (Silent)", value=False)
clip_cap = st.sidebar.number_input("Max Clips Per Video (0 = Full Original Match)", min_value=0, max_value=150, value=0, step=1)
PROMPT_SOURCES = [
    "Auto: Prompts sheet -> cache -> Gemini",
    "Always regenerate with Gemini",
    "Prompts sheet only (never call Gemini)",
]
prompt_source_opt = st.sidebar.selectbox("Prompt / Storyboard source", PROMPT_SOURCES, index=0)
excel_file_path = st.sidebar.text_input("Excel Queue Path", value=os.path.join(BASE_DIR, "youtube_video_pipeline_queue.xlsx"))
pause_for_login_opt = st.sidebar.checkbox("Pause for Manual Sign-In / Captcha", value=True)
auto_refresh = st.sidebar.checkbox("Auto-Refresh Progress (Every 3s)", value=True)

st.markdown('<div class="main-header">🎬 VidoGrid Muse AI Video Pipeline</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated Muse video generation, last-frame image anchoring, and queue monitoring.</div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔴 Live Runtime Progress", "📋 Queue & Launch Controls", "🎞️ Completed Video Gallery", "📜 Terminal Execution Logs", "🧾 Storyboard Prompts"
])


def fmt_duration(seconds: float) -> str:
    if seconds is None:
        return "--"
    s = int(round(seconds))
    if s < 60:
        return f"{s}s"
    return f"{s // 60}m {s % 60}s"


def log_console(msg):
    logger.info(msg)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def run_pipeline_task(api_key, max_clips, headless, excel_path, auto_pause_login, prompt_source="Auto"):
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    except Exception:
        pass

    session_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    active_log_file = os.path.join(LOGS_DIR, f"run_{session_ts}.log")
    setup_run_logger(active_log_file)

    status_data = read_live_status()
    status_data["is_running"] = True
    status_data["is_paused"] = False
    status_data["waiting_for_sign_in"] = False
    status_data["batch_started_at"] = time.time()
    status_data["current_stage"] = "Initializing Worker Thread"
    status_data["current_log_file"] = active_log_file
    status_data["stage_timings"] = {}
    status_data["completed_clips"] = []
    status_data["current_clip"] = 0
    update_live_status(status_data)

    def check_pause_loop():
        while True:
            cur = read_live_status()
            if not cur.get("is_running", False):
                return False
            if not cur.get("is_paused", False) and not cur.get("waiting_for_sign_in", False):
                return True
            time.sleep(1)

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["Video_Queue"]
        cfg = read_active_config(wb)
        provider_label = (cfg or {}).get("provider") or "MUSE AI"
        start_url = provider_start_url(provider_label)
        start_slug = start_url.split("?")[0].rstrip("/").split("/")[-1]
        login_user, login_pass, login_source = VIDOGRID_USER, VIDOGRID_PASS, ".env"
        creds = read_login_details(wb, provider_label)
        if creds:
            login_user, login_pass = creds["user"], creds["password"]
            login_source = f"'Login_Details' sheet (row {creds['row']}, Execute=Yes)"

        headers = [str(ws.cell(row=1, column=c).value).strip() for c in range(1, ws.max_column + 1)]
        col_map = {h.upper(): i + 1 for i, h in enumerate(headers)}

        for r in range(2, ws.max_row + 1):
            vid_id = ws.cell(row=r, column=col_map.get("VIDEO_ID", 1)).value
            url = ws.cell(row=r, column=col_map.get("YOUTUBE_URL", 2)).value
            execute_val = str(get_cell(ws, r, col_map, "EXECUTE") or "").strip().upper()
            status = str(get_cell(ws, r, col_map, "STATUS") or "").strip()
            theme = get_cell(ws, r, col_map, "STYLE_THEME") or "Premium 3D animated nursery rhyme"
            seg_sec = ((cfg or {}).get("duration")
                       or parse_seconds(get_cell(ws, r, col_map, "Each video duration"))
                       or parse_seconds(get_cell(ws, r, col_map, "BATCH_SEGMENT_SEC"))
                       or DEFAULT_DURATION_SEC)
            seg_sec = int(seg_sec) if float(seg_sec).is_integer() else seg_sec

            if not url or execute_val != "YES" or status in ["Completed", "Failed - Video Unavailable"]:
                continue

            old_live = read_live_status()
            if old_live.get("completed_clips"):
                save_to_history(old_live.get("current_video_id", "VID_001"), old_live.get("completed_clips"))

            status_data["current_video_id"] = vid_id
            status_data["completed_clips"] = []
            status_data["current_clip"] = 0
            update_live_status(status_data)

            status_data["current_stage"] = "Downloading Reference Video"
            status_data["current_stage_start"] = time.time()
            update_live_status(status_data)

            t0 = time.time()
            ref_path = os.path.join(CLIPS_DIR, f"{vid_id}_ref.mp4")

            local_ref_candidates = [
                ref_path,
                os.path.join(BASE_DIR, f"{vid_id}_ref.mp4"),
                os.path.join(BASE_DIR, f"{vid_id}_ref_2.mp4"),
                f"{vid_id}_ref.mp4"
            ]
            found_local = ""
            for cand in local_ref_candidates:
                if os.path.exists(cand) and os.path.getsize(cand) > 50000:
                    found_local = cand
                    break

            if found_local and found_local != ref_path:
                log_console(f"Using existing local reference video: {found_local}")
                import shutil
                shutil.copy(found_local, ref_path)

            if not (os.path.exists(ref_path) and os.path.getsize(ref_path) > 50000):
                log_console("Downloading reference video via yt-dlp...")
                ydl_opts_fallback = {
                    'format': '18/best[ext=mp4][height<=360]/best',
                    'outtmpl': ref_path,
                    'overwrites': True,
                    'quiet': True,
                    'no_warnings': True,
                    'extractor_args': {'youtube': {'player_client': ['web_creator', 'android']}},
                    'force_keyframes_at_cuts': True
                }
                success_dl = False
                try:
                    with yt_dlp.YoutubeDL(ydl_opts_fallback) as ydl:
                        ydl.extract_info(url, download=True)
                    success_dl = True
                except Exception as fatal_dl_err:
                    log_console(f"yt-dlp download notice ({fatal_dl_err}).")

                if not success_dl:
                    if found_local:
                        log_console(f"Fallback download failed, but local reference video was found at {found_local}. Proceeding...")
                        import shutil
                        shutil.copy(found_local, ref_path)
                    else:
                        log_console(f"Failed to download video {url} and no local reference found.")
                        ws.cell(row=r, column=col_map.get("STATUS", 6)).value = "Failed - Video Unavailable"
                        wb.save(excel_path)
                        continue

            cap = cv2.VideoCapture(ref_path)
            fps = cap.get(cv2.CAP_PROP_FPS) or 30
            frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            duration_sec = frame_count / fps if fps > 0 else 35.0
            cap.release()

            content_sec = duration_sec - INTRO_OUTRO_TRIM_SEC if duration_sec > 60 else duration_sec
            total_clips = max_clips if max_clips > 0 else max(1, int(round(content_sec / seg_sec)))

            status_data["stage_timings"]["Download Reference"] = fmt_duration(time.time() - t0)
            status_data["total_clips"] = total_clips
            ws.cell(row=r, column=col_map["TARGET_LENGTH"]).value = f"{int(duration_sec) // 60:02d}:{int(duration_sec) % 60:02d}"
            wb.save(excel_path)

            if not check_pause_loop():
                return

            prompts_excel_path = os.path.join(PROMPTS_DIR, f"{vid_id}_{seg_sec}s_prompts.xlsx")
            src_auto = prompt_source.startswith("Auto")
            src_gemini = prompt_source.startswith("Always")
            src_sheet_only = prompt_source.startswith("Prompts sheet")
            prompt_origin = ""
            prompts = [] if src_gemini else load_prompts_from_queue(wb, vid_id, seg_sec, (cfg or {}).get("orientation"))
            if prompts:
                prompt_origin = "Prompts sheet"
            elif src_sheet_only:
                raise RuntimeError(f"No 'Prompts' sheet rows for {vid_id} at {seg_sec}s.")

            if not prompts and src_auto and os.path.exists(prompts_excel_path):
                try:
                    df_saved = pd.read_excel(prompts_excel_path, sheet_name="Storyboard_Prompts", engine="openpyxl")
                    if "SCENE_DESCRIPTION" in df_saved.columns:
                        prompts = df_saved["SCENE_DESCRIPTION"].dropna().astype(str).tolist()
                        if prompts:
                            prompt_origin = "Cache file"
                except Exception:
                    prompts = []

            if not prompts:
                status_data["current_stage"] = f"Deconstructing Storyboard via Gemini ({total_clips} clips)"
                status_data["current_stage_start"] = time.time()
                update_live_status(status_data)

                t_gemini = time.time()
                client = genai.Client(api_key=api_key)
                with open(ref_path, "rb") as vf:
                    video_part = types.Part.from_bytes(data=vf.read(), mime_type="video/mp4")

                prompt_text = (
                    f"Analyze the attached video and recreate it as exactly {total_clips} sequential clips of exactly {seg_sec} seconds each. "
                    f"Theme: {theme}. Format: {orientation_text(cfg)}. "
                    "RULES: Invent original characters (Leo, Mia, Clara, David), no copyright branding. "
                    f"Return strictly a JSON array of exactly {total_clips} string prompts."
                )
                resp = client.models.generate_content(
                    model=GEMINI_MODELS[0],
                    contents=[video_part, prompt_text],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                prompts = [trim_prompt(x) for x in normalize_prompts(json.loads(resp.text))]
                prompt_origin = "Gemini"
                status_data["stage_timings"]["Gemini Storyboard"] = fmt_duration(time.time() - t_gemini)

            if max_clips > 0:
                prompts = prompts[:max_clips]
            status_data["total_clips"] = len(prompts)
            status_data["storyboard"] = {
                "video_id": vid_id, "source": prompt_origin, "duration_sec": seg_sec,
                "prompts": [{"clip": i + 1, "chars": len(x), "prompt": x} for i, x in enumerate(prompts)],
            }
            update_live_status(status_data)

            t_browser_init = time.time()
            status_data["current_stage"] = f"Launching {provider_label} Studio"
            status_data["current_stage_start"] = time.time()
            update_live_status(status_data)

            clip_paths = []
            browser_profile_dir = os.path.join(WORKSPACE_DIR, "browser_profile")
            os.makedirs(browser_profile_dir, exist_ok=True)

            with sync_playwright() as p:
                launch_kwargs = {
                    "user_data_dir": browser_profile_dir, "headless": headless, "accept_downloads": True,
                    "viewport": {"width": 1280, "height": 800}, "ignore_default_args": ["--enable-automation"],
                    "args": ["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-infobars", "--start-maximized"]
                }
                try:
                    context = p.chromium.launch_persistent_context(channel="chrome", **launch_kwargs)
                except Exception:
                    context = p.chromium.launch_persistent_context(**launch_kwargs)

                page = context.pages[0] if context.pages else context.new_page()

                page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(3000)

                try:
                    user_field = page.locator("input[placeholder*='email' i], input[name='email'], input[type='text']").first
                    pwd_field = page.locator("input[type='password']").first
                    if login_user and login_pass and user_field.is_visible() and pwd_field.is_visible():
                        user_field.fill(login_user)
                        pwd_field.fill(login_pass)
                        page.locator("button:has-text('Sign in'), button[type='submit']").first.click(force=True)
                        page.wait_for_timeout(4000)
                except Exception:
                    pass

                if start_slug not in page.url:
                    page.goto(start_url, wait_until="domcontentloaded", timeout=60000)
                    page.wait_for_timeout(4000)

                status_data["stage_timings"]["Browser Init"] = fmt_duration(time.time() - t_browser_init)

                if auto_pause_login and ("auth/login" in page.url or "Verify you're human" in page.content()):
                    status_data["waiting_for_sign_in"] = True
                    status_data["is_paused"] = True
                    status_data["current_stage"] = "⏸️ Waiting for Login / Verification"
                    update_live_status(status_data)
                    if not check_pause_loop():
                        context.close()
                        return
                    page.wait_for_timeout(2000)

                last_anchor_frame = None

                for idx, p_text in enumerate(prompts, start=1):
                    if not check_pause_loop():
                        context.close()
                        return

                    if start_slug not in page.url:
                        page.goto(start_url, wait_until="domcontentloaded", timeout=30000)
                        page.wait_for_timeout(2000)

                    clip_start_time = time.time()
                    clean_prompt = build_final_prompt(p_text, cfg)
                    c_path = os.path.join(CLIPS_DIR, f"{vid_id}_part_{idx:03d}.mp4")
                    f_path = os.path.join(FRAMES_DIR, f"{vid_id}_frame_{idx:03d}.png")

                    status_data["current_clip"] = idx
                    status_data["current_prompt"] = clean_prompt
                    status_data["clip_started_at"] = clip_start_time
                    status_data["current_stage"] = f"Generating Clip {idx}/{len(prompts)} on {provider_label}"
                    update_live_status(status_data)

                    if last_anchor_frame and os.path.exists(last_anchor_frame):
                        try:
                            file_input = page.locator('input[type="file"]').first
                            file_input.set_input_files(last_anchor_frame)
                            page.wait_for_timeout(2000)
                        except Exception:
                            pass

                    apply_page_settings(page, seg_sec, cfg)

                    try:
                        prompt_textarea = page.locator("textarea:visible, [contenteditable='true']:visible").last
                        prompt_textarea.click(force=True)
                        page.keyboard.press("Control+A")
                        page.keyboard.press("Backspace")
                        page.wait_for_timeout(200)
                        prompt_textarea.fill(clean_prompt)
                        page.wait_for_timeout(500)
                    except Exception:
                        pass

                    click_generate(page, vid_id, idx)
                    page.wait_for_timeout(3000)

                    poll_start = time.time()
                    download_ready = False
                    while time.time() - poll_start < 600:
                        if not check_pause_loop():
                            context.close()
                            return
                        try:
                            dl_btn = page.locator("button:has-text('Download Video'), [aria-label*='Download' i]").first
                            if dl_btn.is_visible():
                                download_ready = True
                                break
                        except Exception:
                            pass
                        time.sleep(3)

                    if not download_ready:
                        raise RuntimeError(f"Timeout waiting for video generation output on Clip {idx}")

                    try:
                        with page.expect_download(timeout=15000) as d_info:
                            page.locator("button:has-text('Download Video')").first.click(force=True)
                            page.wait_for_timeout(1000)
                            ellipses = page.locator("[aria-label*='More' i], button:has(svg)").last
                            if ellipses.is_visible():
                                ellipses.click(force=True)
                                page.wait_for_timeout(500)
                                download_option = page.locator("text=Download").last
                                if download_option.is_visible():
                                    download_option.click(force=True)
                        d_info.value.save_as(c_path)
                    except Exception:
                        v_src = page.locator("video").first.get_attribute("src")
                        if v_src:
                            resp = context.request.get(v_src)
                            with open(c_path, "wb") as vf:
                                vf.write(resp.body())

                    cap = cv2.VideoCapture(c_path)
                    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                    if total_frames > 0:
                        cap.set(cv2.CAP_PROP_POS_FRAMES, total_frames - 1)
                    success, last_frame_img = cap.read()
                    if success:
                        cv2.imwrite(f_path, last_frame_img)
                        last_anchor_frame = f_path
                    cap.release()

                    clip_paths.append(c_path)
                    status_data["completed_clips"].append({
                        "clip_num": idx, "duration_sec": seg_sec, "file_path": c_path, "frame_path": f_path,
                        "prompt": clean_prompt, "elapsed_sec": round(time.time() - clip_start_time, 1),
                        "completed_time": datetime.now().strftime("%H:%M:%S")
                    })
                    update_live_status(status_data)

                context.close()

            t_stitch = time.time()
            status_data["current_stage"] = f"Concatenating {len(clip_paths)} clips into final master video"
            update_live_status(status_data)

            final_mp4 = os.path.join(OUTPUT_DIR, f"{vid_id}_final_vidogrid.mp4")
            fl_txt = os.path.join(CLIPS_DIR, f"{vid_id}_filelist.txt")
            with open(fl_txt, "w") as f:
                for cp in clip_paths:
                    f.write("file '" + cp + "'\n")

            subprocess.run(["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", fl_txt, "-c", "copy", final_mp4], check=True)
            status_data["stage_timings"]["Video Assembly"] = fmt_duration(time.time() - t_stitch)

            ws.cell(row=r, column=col_map["STATUS"]).value = "Completed"
            wb.save(excel_path)

    except Exception as e:
        status_data = read_live_status()
        status_data["current_stage"] = f"❌ Error: {str(e)[:80]}"
        status_data["is_running"] = False
        update_live_status(status_data)
    finally:
        status_data = read_live_status()
        has_error = "❌ Error" in status_data.get("current_stage", "")
        status_data["is_running"] = False
        status_data["is_paused"] = False
        status_data["waiting_for_sign_in"] = False
        if not has_error:
            status_data["current_stage"] = "Idle / Completed"
        update_live_status(status_data)


with tab1:
    live = read_live_status()

    if live.get("waiting_for_sign_in"):
        st.warning("🔐 **Manual Sign-In / Captcha Required**: Please complete verification in the open Chromium browser window.")
        if st.button("✅ Verification Done — Resume Pipeline", key="btn_signin_done_tab1", type="primary", use_container_width=True):
            live["waiting_for_sign_in"] = False
            live["is_paused"] = False
            update_live_status(live)
            st.success("Resuming generation pipeline...")
            st.rerun()

    elif live.get("is_running"):
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            if live.get("is_paused"):
                if st.button("▶️ Resume Pipeline", key="btn_resume_pipeline_tab1", type="primary", use_container_width=True):
                    live["is_paused"] = False
                    update_live_status(live)
                    st.rerun()
            else:
                if st.button("⏸️ Pause Pipeline", key="btn_pause_pipeline_tab1", use_container_width=True):
                    live["is_paused"] = True
                    update_live_status(live)
                    st.rerun()
        with col_btn2:
            if st.button("🔑 Pause for Sign-In / Captcha", key="btn_pause_signin_tab1", use_container_width=True):
                live["waiting_for_sign_in"] = True
                live["is_paused"] = True
                update_live_status(live)
                st.rerun()

    col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns([1.2, 2, 1, 1, 1])
    with col_s1:
        if live.get("waiting_for_sign_in"):
            st.write("**State:** `🟡 Waiting Captcha`")
        elif live.get("is_paused"):
            st.write("**State:** `⏸️ Paused`")
        elif live.get("is_running"):
            st.write("**State:** `🟢 Generating`")
        else:
            st.write("**State:** `⚪ Idle`")
        st.write(f"**Task:** `{live.get('current_video_id', 'None')}`")
    with col_s2:
        st.write(f"**Action:** {live.get('current_stage', 'Ready')}")
        tot = live.get("total_clips", 0)
        curr = live.get("current_clip", 0)
        st.progress(min((curr / tot) if tot > 0 else 0.0, 1.0))
    with col_s3:
        st.metric(label="Clips Completed", value=f"{len(live.get('completed_clips', []))} / {tot}")
    with col_s4:
        stage_elapsed_str = fmt_duration(time.time() - live["current_stage_start"]) if live.get("is_running") and live.get("current_stage_start") else "--"
        st.metric(label="Stage Elapsed", value=stage_elapsed_str)
    with col_s5:
        total_batch_str = fmt_duration(time.time() - live["batch_started_at"]) if live.get("is_running") and live.get("batch_started_at") else "--"
        st.metric(label="Total Elapsed", value=total_batch_str)

    st.markdown("##### ⏱️ Pipeline Stages Duration Breakdown")
    timings = live.get("stage_timings", {})
    t_c1, t_c2, t_c3, t_c4, t_c5 = st.columns(5)
    curr_stage = live.get("current_stage", "").lower()
    curr_elapsed = fmt_duration(time.time() - live["current_stage_start"]) if live.get("is_running") and live.get("current_stage_start") else "--"

    with t_c1:
        if "Download Reference" in timings:
            st.metric("1. Video Download", value=timings["Download Reference"], delta="Completed")
        elif "downloading" in curr_stage:
            st.metric("1. Video Download", value=curr_elapsed, delta="In Progress")
        else:
            st.metric("1. Video Download", value="--", delta="Pending")

    with t_c2:
        if "Gemini Storyboard" in timings:
            _sb = timings["Gemini Storyboard"]
            st.metric("2. Gemini Storyboard", value=_sb, delta="Skipped (ready prompts)" if _sb == "Skipped" else "Completed")
        elif "gemini" in curr_stage or "deconstructing" in curr_stage:
            st.metric("2. Gemini Storyboard", value=curr_elapsed, delta="In Progress")
        else:
            st.metric("2. Gemini Storyboard", value="--", delta="Pending")

    with t_c3:
        if "Browser Init" in timings:
            st.metric("3. Browser Init", value=timings["Browser Init"], delta="Completed")
        elif "automation" in curr_stage or "browser" in curr_stage:
            st.metric("3. Browser Init", value=curr_elapsed, delta="In Progress")
        else:
            st.metric("3. Browser Init", value="--", delta="Pending")

    with t_c4:
        clips_done = len(live.get("completed_clips", []))
        total_clip_render_sec = sum(c.get("elapsed_sec", 0) for c in live.get("completed_clips", []))
        if "generating clip" in curr_stage:
            st.metric("4. Clip Render", value=curr_elapsed, delta=f"Clip {live.get('current_clip')} Rendering")
        elif clips_done > 0:
            st.metric("4. Clips Render Total", value=fmt_duration(total_clip_render_sec), delta=f"{clips_done} Done")
        else:
            st.metric("4. Clip Render", value="--", delta="Pending")

    with t_c5:
        if "Video Assembly" in timings:
            st.metric("5. Final Stitch", value=timings["Video Assembly"], delta="Completed")
        elif "concatenating" in curr_stage or "stitch" in curr_stage:
            st.metric("5. Final Stitch", value=curr_elapsed, delta="In Progress")
        else:
            st.metric("5. Final Stitch", value="--", delta="Pending")

    st.markdown("---")

    if live.get("is_running") and live.get("current_prompt"):
        st.markdown(f"**⚡ Current Prompt Running on VidoGrid Muse (Clip {live.get('current_clip')} of {live.get('total_clips')}):**")
        st.markdown(f'<div class="prompt-box">{html.escape(str(live.get("current_prompt")))}</div>', unsafe_allow_html=True)
        st.write("")

    st.subheader("🎞️ Generated Clips Stream (Chronological)")
    completed = live.get("completed_clips", [])

    if completed:
        cols_per_row = 2
        rows = [completed[i:i + cols_per_row] for i in range(0, len(completed), cols_per_row)]
        for row in rows:
            grid_cols = st.columns(cols_per_row)
            for c_idx, item in enumerate(row):
                with grid_cols[c_idx]:
                    st.markdown(
                        f'<div class="clip-card">'
                        f'<div class="clip-title">Clip {item["clip_num"]} ({item.get("duration_sec", 10):g}-second segment)</div>'
                        f'<div class="clip-meta">⏱️ <b>Render Time:</b> {item.get("elapsed_sec", "--")}s &nbsp;|&nbsp; 🕒 <b>Finished:</b> {item.get("completed_time", "--")}</div>',
                        unsafe_allow_html=True
                    )
                    if os.path.exists(item["file_path"]):
                        st.video(item["file_path"])
                    st.markdown("**Prompt Used:**")
                    st.markdown(f'<div class="prompt-box">{html.escape(str(item.get("prompt", "")))}</div>', unsafe_allow_html=True)
                    st.write("")
                    if os.path.exists(item.get("frame_path", "")):
                        with st.expander("🔍 View Continuity Anchor Frame (Last-Frame Seed)"):
                            st.image(item["frame_path"], caption=f"Last Frame of Clip {item['clip_num']}", use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("No clips rendered in current active session. Clips will stream here as they finish.")

    st.markdown("---")
    st.subheader("📚 Past Execution History (Video-Wise)")
    history_data = load_execution_history()
    if history_data:
        if st.button("🗑️ Clear All Execution History", key="btn_clear_history"):
            if os.path.exists(HISTORY_FILE):
                try:
                    os.remove(HISTORY_FILE)
                except Exception:
                    pass
            st.success("History cleared!")
            st.rerun()

        for run in history_data:
            with st.expander(f"📁 Video ID: {run.get('video_id')} | Finished at: {run.get('timestamp')} ({len(run.get('clips', []))} clips)"):
                h_clips = run.get("clips", [])
                h_cols_per_row = 2
                h_rows = [h_clips[i:i + h_cols_per_row] for i in range(0, len(h_clips), h_cols_per_row)]
                for h_row in h_rows:
                    h_grid = st.columns(h_cols_per_row)
                    for h_idx, h_item in enumerate(h_row):
                        with h_grid[h_idx]:
                            st.markdown(
                                f'<div class="clip-card">'
                                f'<div class="clip-title">Clip {h_item["clip_num"]} ({h_item.get("duration_sec", 10):g}-second segment)</div>'
                                f'<div class="clip-meta">⏱️ <b>Render Time:</b> {h_item.get("elapsed_sec", "--")}s &nbsp;|&nbsp; 🕒 <b>Finished:</b> {h_item.get("completed_time", "--")}</div>',
                                unsafe_allow_html=True
                            )
                            if os.path.exists(h_item["file_path"]):
                                st.video(h_item["file_path"])
                            st.markdown("**Prompt Used:**")
                            st.markdown(f'<div class="prompt-box">{html.escape(str(h_item.get("prompt", "")))}</div>', unsafe_allow_html=True)
                            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.caption("No past execution history archived yet.")

with tab2:
    col_q1, col_q2 = st.columns([2, 1])
    with col_q1:
        st.subheader("Queue Manager (Excel Source)")
        if os.path.exists(excel_file_path) and os.path.getsize(excel_file_path) > 0:
            try:
                df_queue = pd.read_excel(excel_file_path, sheet_name="Video_Queue", engine="openpyxl")
            except Exception as e:
                st.error(f"Error reading Excel file: {e}")
                df_queue = pd.DataFrame()
            edited_df = st.data_editor(
                df_queue,
                num_rows="dynamic",
                use_container_width=True,
                height=260
            )
            if st.button("💾 Save Excel Changes", key="btn_save_excel_changes"):
                with pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    edited_df.to_excel(writer, sheet_name="Video_Queue", index=False)
                st.success("Queue updated successfully!")
        else:
            st.warning(f"Excel queue not found at `{excel_file_path}`.")
            df_queue = pd.DataFrame()

    with col_q2:
        st.subheader("Launch Operations")
        
        live_state = read_live_status()
        is_busy = live_state.get("is_running", False)
        
        if is_busy:
            st.warning("⚠️ **Engine Status: RUNNING** (Automation in progress...)")
        else:
            st.success("🟢 **Engine Status: IDLE** (Ready to launch)")

        if live_state.get("waiting_for_sign_in"):
            if st.button("✅ Verification Done — Resume", key="btn_signin_done_tab2", type="primary", use_container_width=True):
                live_state["waiting_for_sign_in"] = False
                live_state["is_paused"] = False
                update_live_status(live_state)
                st.success("Resuming...")
                st.rerun()

        if st.button("🔓 Force Unlock / Reset Engine State", key="btn_force_unlock", use_container_width=True):
            live_state = read_live_status()
            live_state["is_running"] = False
            live_state["is_paused"] = False
            live_state["waiting_for_sign_in"] = False
            live_state["current_stage"] = "Idle / Unlocked"
            live_state["completed_clips"] = []
            live_state["current_clip"] = 0
            live_state["total_clips"] = 0
            live_state["current_prompt"] = ""
            update_live_status(live_state)
            st.success("State unlocked and reset successfully!")
            st.rerun()

        run_btn = st.button(
            "🚀 Start VidoGrid Muse Pipeline",
            key="btn_start_pipeline",
            type="primary",
            disabled=is_busy
        )

        if run_btn:
            init_status = read_live_status()
            init_status["is_running"] = True
            init_status["current_stage"] = "🚀 Pipeline Starting..."
            update_live_status(init_status)
            
            threading.Thread(
                target=run_pipeline_task,
                args=(api_key_input, clip_cap, headless_mode, excel_file_path, pause_for_login_opt, prompt_source_opt),
                daemon=True
            ).start()
            st.toast("Pipeline started! Switching to Live Runtime Progress...")
            st.rerun()

with tab3:
    st.subheader("Master Video Gallery")
    for f in os.listdir(OUTPUT_DIR):
        if f.endswith(".mp4"):
            fp = os.path.join(OUTPUT_DIR, f)
            st.video(fp)

with tab4:
    st.subheader("Console Logs")
    if st.button("🔄 Refresh Logs View", key="btn_refresh_logs"):
        st.rerun()

    active_log = live.get("current_log_file")
    if active_log and os.path.exists(active_log):
        with open(active_log, "r", encoding="utf-8") as lf:
            lines = lf.readlines()
        log_text = "".join(lines[-150:])
        st.caption(f"📁 Active Session Log: `{os.path.basename(active_log)}`")
    else:
        log_text = "No active execution session yet. Start the pipeline from Tab 2 to initialize fresh logs."
    
    st.text_area("Console Stream", value=log_text, height=450)

with tab5:
    sb = live.get("storyboard") or {}
    st.subheader("Storyboard Prompts & Continuity Guidelines")
    if sb.get("prompts"):
        st.caption(f"Video **{sb.get('video_id')}** | Source: **{sb.get('source')}** | "
                   f"{len(sb['prompts'])} clips x {sb.get('duration_sec')}s | limit {PROMPT_MAX_CHARS} chars per prompt")
        for item in sb["prompts"]:
            with st.expander(f"Clip {item['clip']} ({item['chars']} chars)"):
                st.markdown(f'<div class="prompt-box">{html.escape(str(item["prompt"]))}</div>', unsafe_allow_html=True)
    else:
        st.info("No storyboard yet. It appears here as soon as a run has its prompts (from the Prompts sheet, cache or Gemini).")

if auto_refresh and live.get("is_running") and not live.get("waiting_for_sign_in"):
    time.sleep(3)
    st.rerun()