"""
pdf_to_notebooklm.py — อัปโหลด PDF ของแต่ละวิชา (โฟลเดอร์ย่อย) ไปยัง NotebookLM วิชาละ 1 notebook
ถ้า notebook ที่จำไว้ถูกลบหรือเปิดไม่ได้ จะสร้างใหม่แล้วอัปโหลดไฟล์ทั้งหมดของวิชานั้นอีกครั้ง

ติดตั้ง (ครั้งเดียว):  python -m pip install playwright
รัน:                  py "<path ของไฟล์นี้>/pdf_to_notebooklm.py"

ก่อนใช้: ตั้งค่า OneDrive sync กับ Team's SharePoint และเช็ก SOURCE_DIR, CHROME_PROFILE ให้ตรงกับเครื่อง
"""

import json
import re
import socket
import subprocess
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Error as PWError, TimeoutError as PWTimeout

# ================= ตั้งค่า =================
SOURCE_DIR = r"D:\OneDrive_RMUTT\Study_Source\PDF\Rajamangala University of Technology Thanyaburi"  # โฟลเดอร์แม่ ข้างในคือโฟลเดอร์ย่อยแยกตามวิชา

# False = เช็กจากไฟล์ในเครื่องก่อน เข้าเว็บเฉพาะวิชาที่มีไฟล์ใหม่
#         (ไม่มีอะไรใหม่ = ไม่เปิด Chrome เลย เร็ว)
# True  = เปิดทุก notebook ทุกครั้งเพื่อตรวจว่ายังอยู่จริง (ช้า ใช้เป็นครั้งคราวเท่านั้น)
# ถ้า notebook ถูกลบ จะรู้ตอนที่วิชานั้นมีไฟล์ใหม่ แล้วสร้างใหม่ให้เอง
VERIFY_NOTEBOOKS = False

BASE_DIR = Path(__file__).resolve().parent  # โฟลเดอร์ของสคริปต์นี้ (ไฟล์สถานะเก็บที่นี่ ไม่ขึ้นกับว่ารันจากโฟลเดอร์ไหน)
RECORD_FILE = BASE_DIR / "notebooks.json"  # จำ notebook ของแต่ละวิชา + ไฟล์ที่อัปโหลดแล้ว

HOME_URL = "https://notebook.google.com/"

DEBUG_PORT = 9222
CHROME_PROFILE = r"C:\chrome-debug"  # โปรไฟล์ Chrome แยก (จำการล็อกอิน Google ไว้)
CHROME_PATH = ""  # ถ้าหา Chrome ไม่เจอ ใส่ path เต็มของ chrome.exe

LOGIN_WAIT_SECONDS = 300      # ถ้าต้องล็อกอิน Google จะรอสูงสุดกี่วินาที
EXISTING_CHECK_SECONDS = 30   # เปิด notebook ที่จำไว้แล้วไม่พร้อมใน N วินาที = ถือว่าถูกลบ/หาย
VERIFY_SECONDS = 90           # รอให้ชื่อไฟล์ขึ้นในแถบแหล่งข้อมูลสูงสุดกี่วินาที
SETTLE_QUIET_SECONDS = 8      # ต้องไม่เห็นสถานะ "กำลังอัปโหลด/ประมวลผล" ต่อเนื่องกี่วินาทีถึงถือว่าเสร็จ
SETTLE_MAX_SECONDS = 300      # รอให้ NotebookLM ประมวลผลไฟล์ไม่เกินกี่วินาที (ต่อวิชา)
SETTLE_BASE_SECONDS = 15      # เวลารอขั้นต่ำ = ค่านี้ + SETTLE_PER_FILE_SECONDS x จำนวนไฟล์
SETTLE_PER_FILE_SECONDS = 6

ADD_SOURCE_BTN = re.compile(r"เพิ่มแหล่งข้อมูล|Add source", re.I)
UPLOAD_FILE_BTN = re.compile(r"อัปโหลดไฟล์|Upload files?", re.I)
NEW_NOTEBOOK_BTN = re.compile(r"Notebook ใหม่|สร้าง Notebook|New notebook|Create new", re.I)
BUSY_TEXT = re.compile(r"กำลัง(อัปโหลด|ประมวลผล|นำเข้า|เพิ่ม)|Uploading|Processing|Importing", re.I)
NOTEBOOK_URL_RE = re.compile(r"/notebook/[0-9a-fA-F-]{36}")
DEBUG_SHOT = BASE_DIR / "debug_upload.png"  # ภาพหน้าจอเก็บไว้ดูถ้าขั้นตอนไหนพลาด
# ==========================================


# ---------- บันทึกสถานะ ----------

def load_record() -> dict:
    try:
        data = json.loads(RECORD_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_record(record: dict):
    RECORD_FILE.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------- ส่วนที่ 1: หาโฟลเดอร์วิชา + เตรียมไฟล์ ----------

def discover_subjects(root: Path) -> dict[str, list[Path]]:
    """โฟลเดอร์ย่อยแต่ละอันที่มี PDF = 1 วิชา (ไล่หาในโฟลเดอร์ซ้อนด้วย)"""
    if not root.exists():
        raise SystemExit(f"ไม่พบโฟลเดอร์: {root} — แก้ SOURCE_DIR ในสคริปต์")
    subjects: dict[str, list[Path]] = {}
    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        pdfs = sorted(d.rglob("*.pdf"))
        if pdfs:
            subjects[d.name] = pdfs
        else:
            others = sorted({f.suffix.lower() or "(ไม่มีนามสกุล)" for f in d.rglob("*") if f.is_file()})
            print(f"⚠ ข้ามโฟลเดอร์ '{d.name}': ไม่พบไฟล์ .pdf"
                  + (f" (เจอไฟล์ชนิด: {', '.join(others)})" if others else " (โฟลเดอร์ว่าง หรือไฟล์ยังเป็นออนไลน์อย่างเดียวบน OneDrive)"))
    return subjects


def prepare_subject(pdfs: list[Path]) -> list[Path]:
    """ไฟล์ที่จะอัปโหลดของวิชานั้น (ไฟล์ต้นฉบับ) ชื่อซ้ำกันในวิชาเดียวกัน = ข้ามตัวหลังและเตือน"""
    by_name: dict[str, Path] = {}
    for pdf in pdfs:
        if pdf.name in by_name:
            print(f"  ⚠ ข้ามไฟล์ชื่อซ้ำ: {pdf}")
        else:
            by_name[pdf.name] = pdf
    return list(by_name.values())


def new_files(entry: dict, files: list[Path]) -> list[Path]:
    """ไฟล์ที่ยังไม่เคยอัปโหลดตามที่บันทึกไว้ใน entry ของวิชานั้น"""
    done = set(entry.get("uploaded", []))
    return [f for f in files if f.name not in done]


# ---------- ส่วนที่ 2: เปิด/เชื่อมต่อ Chrome ----------

def port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


def ensure_chrome():
    if port_open(DEBUG_PORT):
        return
    candidates = [
        CHROME_PATH,
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        str(Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "Application" / "chrome.exe"),
    ]
    exe = next((c for c in candidates if c and Path(c).exists()), None)
    if not exe:
        raise SystemExit("ไม่พบ chrome.exe — ใส่ path ที่ตัวแปร CHROME_PATH")
    subprocess.Popen(
        [
            exe,
            f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={CHROME_PROFILE}",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized",
        ]
    )
    for _ in range(40):
        if port_open(DEBUG_PORT):
            return
        time.sleep(0.5)
    raise SystemExit("เปิด Chrome ไม่สำเร็จ (debug port ไม่ตอบสนอง)")


# ---------- ส่วนที่ 3: ควบคุมหน้า NotebookLM ----------

def visible(locator) -> bool:
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except PWTimeout:
        return False


def pick_upload_button(page):
    """หาปุ่ม 'อัปโหลดไฟล์' ที่มองเห็นอยู่ในหน้า (ถ้ามี)"""
    for loc in (
        page.get_by_role("button", name=UPLOAD_FILE_BTN),
        page.get_by_text(UPLOAD_FILE_BTN),
    ):
        if visible(loc):
            return loc.first
    return None


def find_new_notebook_button(page):
    for loc in (
        page.get_by_role("button", name=NEW_NOTEBOOK_BTN),
        page.get_by_role("link", name=NEW_NOTEBOOK_BTN),
        page.get_by_text(NEW_NOTEBOOK_BTN),
    ):
        if visible(loc):
            return loc.first
    return None


def notebook_ready(page) -> bool:
    return visible(page.get_by_role("button", name=ADD_SOURCE_BTN)) or pick_upload_button(page) is not None


def wait_until(page, ready_fn, what: str):
    """รอจน ready_fn() จริง แล้วคืนค่านั้น ถ้าถูกพาไปหน้าล็อกอิน Google จะรอให้ล็อกอินในหน้าต่างนั้น"""
    warned = False
    deadline = time.time() + LOGIN_WAIT_SECONDS
    while time.time() < deadline:
        result = ready_fn()
        if result:
            return result
        if "accounts.google.com" in page.url and not warned:
            print(f"⏳ ต้องล็อกอิน Google — ล็อกอินในหน้าต่าง Chrome (รอสูงสุด {LOGIN_WAIT_SECONDS} วินาที)")
            warned = True
        page.wait_for_timeout(1000)
    page.screenshot(path=str(DEBUG_SHOT))
    raise RuntimeError(f"รอ {what} ไม่สำเร็จ (ดูภาพ {DEBUG_SHOT.resolve()})")


def open_existing_notebook(ctx, url: str):
    """เปิด notebook ที่จำไว้ คืน page ถ้าใช้งานได้ หรือ None ถ้าถูกลบ/เปิดไม่ได้"""
    page = ctx.new_page()
    page.goto(url)
    deadline = time.time() + EXISTING_CHECK_SECONDS
    while time.time() < deadline:
        if "accounts.google.com" in page.url:  # ต้องล็อกอินก่อน รอให้ล็อกอินเสร็จแล้วค่อยตัดสิน
            wait_until(page, lambda: "accounts.google.com" not in page.url, "การล็อกอิน Google")
            deadline = time.time() + EXISTING_CHECK_SECONDS
        if notebook_ready(page):
            page.wait_for_timeout(1500)
            return page
        if not NOTEBOOK_URL_RE.search(page.url):  # ถูกเด้งออกไปหน้าอื่น = notebook ไม่อยู่แล้ว
            break
        page.wait_for_timeout(1000)
    page.screenshot(path=str(DEBUG_SHOT))
    page.close()
    return None


def create_notebook(ctx, subject: str):
    """สร้าง notebook ใหม่ แล้วคืน (page, url) — ชื่อ notebook NotebookLM ตั้งให้เอง"""
    page = ctx.new_page()
    page.goto(HOME_URL)
    wait_until(page, lambda: find_new_notebook_button(page), "ปุ่มสร้าง notebook ใหม่").click()
    page.wait_for_url(NOTEBOOK_URL_RE, timeout=60_000)
    page.wait_for_timeout(2500)
    # ไม่ตั้งชื่อเอง: NotebookLM จะตั้งชื่อให้อัตโนมัติหลังอัปโหลดแหล่งข้อมูล
    print(f"  ✓ สร้าง notebook ใหม่สำหรับ '{subject}'")
    return page, page.url.split("?")[0]


def is_busy(page) -> bool:
    """NotebookLM ยังอัปโหลด/ประมวลผลไฟล์อยู่หรือไม่"""
    return visible(page.get_by_role("progressbar")) or visible(page.get_by_text(BUSY_TEXT))


def wait_until_settled(page, n_files: int) -> bool:
    """รอให้ NotebookLM อัปโหลด/ประมวลผลไฟล์เสร็จ ก่อนจะปิดแท็บหรือไปวิชาถัดไป
    (ถ้าปิดแท็บเร็วไป การอัปโหลดที่ค้างอยู่จะถูกตัดกลางคัน)"""
    min_wait = SETTLE_BASE_SECONDS + SETTLE_PER_FILE_SECONDS * n_files
    start = time.time()
    quiet_since = None
    while time.time() - start < SETTLE_MAX_SECONDS:
        now = time.time()
        if is_busy(page):
            quiet_since = None
        elif quiet_since is None:
            quiet_since = now
        if quiet_since is not None and now - quiet_since >= SETTLE_QUIET_SECONDS and now - start >= min_wait:
            return True
        page.wait_for_timeout(1000)
    return False


def upload_files(page, files: list[Path]) -> set[str]:
    """อัปโหลดไฟล์เข้า notebook ที่เปิดอยู่ แล้วคืนชื่อไฟล์ที่ยืนยันว่าขึ้นในแถบแหล่งข้อมูลแล้ว"""
    todo = [f for f in files if page.get_by_text(f.name, exact=True).count() == 0]
    already = {f.name for f in files} - {f.name for f in todo}
    if not todo:
        print("  ไฟล์ทั้งหมดอยู่ใน notebook แล้ว")
        return already

    paths = [str(f.resolve()) for f in todo]
    try:
        btn = pick_upload_button(page)
        if btn is None:
            page.get_by_role("button", name=ADD_SOURCE_BTN).first.click(timeout=10_000)
            page.wait_for_timeout(1500)
            btn = pick_upload_button(page)
        if btn is not None:
            with page.expect_file_chooser(timeout=15_000) as fc:
                btn.click()
            fc.value.set_files(paths)
        else:
            page.locator("input[type=file]").first.set_input_files(paths)
    except PWTimeout:
        page.screenshot(path=str(DEBUG_SHOT))
        print("  ✗ อัปโหลดไม่สำเร็จ — เก็บภาพหน้าจอไว้ที่", DEBUG_SHOT.resolve())
        return already

    # 1) รอให้ชื่อไฟล์ขึ้นในแถบแหล่งข้อมูล
    confirmed = set(already)
    pending = {f.name for f in todo}
    deadline = time.time() + VERIFY_SECONDS
    while pending and time.time() < deadline:
        for name in list(pending):
            if page.get_by_text(name, exact=True).count() > 0:
                pending.discard(name)
                confirmed.add(name)
        if pending:
            page.wait_for_timeout(2000)

    # 2) ชื่อขึ้นแล้วไม่ได้แปลว่าเสร็จ ต้องรอให้ประมวลผลจบก่อนถึงจะออกจากหน้านี้
    print("  ⏳ รอให้ NotebookLM ประมวลผลไฟล์ให้เสร็จ...")
    if wait_until_settled(page, len(todo)):
        print("  ✓ ประมวลผลเสร็จ")
    else:
        print(f"  ⚠ รอครบ {SETTLE_MAX_SECONDS} วินาทีแล้วยังเห็นสถานะกำลังประมวลผล — ไปต่อ (ตรวจในเว็บอีกที)")

    print(f"  ✓ อัปโหลดสำเร็จ {len(todo) - len(pending)}/{len(todo)} ไฟล์")
    if pending:
        print("    ยังไม่เห็นในแถบแหล่งข้อมูล:", ", ".join(sorted(pending)))
    return confirmed


# ---------- main ----------

def main():
    record = load_record()

    subjects = discover_subjects(Path(SOURCE_DIR))
    print(f"พบ {len(subjects)} วิชา: \n{'\n'.join(subjects) or '-'}\n")

    prepared: dict[str, list[Path]] = {}
    has_new = False
    for subject, pdfs in subjects.items():
        files = prepare_subject(pdfs)
        prepared[subject] = files
        entry = record.get(subject, {})
        new = new_files(entry, files)
        no_notebook = not entry.get("url")
        print(f"[{subject}] PDF {len(files)} | อัปโหลดแล้ว {len(files) - len(new)} | ใหม่ {len(new)}"
              + (" | ยังไม่มี notebook" if no_notebook else ""))
        has_new = has_new or bool(new) or no_notebook

    if not prepared:
        print("ไม่พบโฟลเดอร์วิชาที่มี PDF — จบการทำงาน")
        return
    if not has_new and not VERIFY_NOTEBOOKS:
        print("ไม่มีไฟล์ใหม่ — จบการทำงาน")
        return

    ensure_chrome()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{DEBUG_PORT}")
        ctx = browser.contexts[0]

        for subject, files in prepared.items():
            print(f"\n== {subject} ==")
            page = None
            try:
                entry = record.setdefault(subject, {"url": "", "uploaded": []})

                # 1) ถ้ามี notebook ที่จำไว้ ตรวจว่ายังเปิดได้จริงหรือถูกลบไปแล้ว
                if entry.get("url"):
                    if VERIFY_NOTEBOOKS or new_files(entry, files):
                        page = open_existing_notebook(ctx, entry["url"])
                        if page is None:
                            print("  ⚠ notebook เดิมหาไม่เจอ/ถูกลบ → สร้างใหม่ แล้วอัปโหลดไฟล์ทั้งหมดอีกครั้ง")
                            entry["url"] = ""
                            entry["uploaded"] = []
                            save_record(record)

                new = new_files(entry, files)
                if entry.get("url") and not new:
                    print("  ไม่มีไฟล์ใหม่")
                    continue

                # 2) ยังไม่มี notebook (ใหม่ หรือถูกลบ) → สร้าง
                if not entry.get("url"):
                    page, url = create_notebook(ctx, subject)
                    entry["url"] = url
                    save_record(record)  # จำ notebook ไว้ทันที กันสร้างซ้ำถ้าอัปโหลดพลาด
                    new = list(files)

                # 3) อัปโหลด
                confirmed = upload_files(page, new)
                entry["uploaded"] = sorted(set(entry.get("uploaded", [])) | confirmed)
                save_record(record)
            except Exception as e:  # วิชาหนึ่งพลาด ไม่ให้กระทบวิชาอื่น
                if "closed" in str(e).lower():
                    print(f"  ⚠ แท็บ/Chrome ถูกปิดระหว่างทำงาน ({subject}) — อย่าปิดหน้าต่าง Chrome "
                          "จนกว่าสคริปต์จะจบ รอบหน้าจะตรวจและอัปโหลดต่อให้เอง")
                    if not browser.is_connected():
                        print("  Chrome ถูกปิดแล้ว หยุดการทำงานรอบนี้")
                        break
                else:
                    print(f"  ✗ ข้ามวิชา {subject}: {e}")
            finally:
                if page is not None:
                    try:
                        page.close()
                    except PWError:
                        pass

    print("\nเสร็จแล้ว (Chrome ยังเปิดอยู่ ปิดเองได้)")


if __name__ == "__main__":
    main()