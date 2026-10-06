# ==============================================================================
# sap_export.py — pulls the two SAP inputs of the reconciliation engine through
# the SAP GUI Scripting API (no keystrokes, no screen coordinates):
#
#   1. Supplier master data  — custom table viewer on LFA1  -> SUPPLIERS_FILE
#   2. Vendor line items     — FBL1N (company code, vendor ranges, posting dates,
#                              saved layout)                -> LINE_ITEMS_FILE
#
# Display / export only — nothing is ever saved in SAP.
#
# Lessons from production built into the code:
#   * The custom table export only accepts a target file that ALREADY exists
#     ("File does not exist"), so an empty workbook is created first and the
#     script waits for it to CHANGE, not merely to exist.
#   * SAP writes the file to a plain-ASCII temp folder; Python moves it to the
#     final (possibly non-ASCII) folder/name afterwards.
#   * Large tables reach the GUI in several packages — the export starts only
#     after the grid holds every row announced in the window title.
#   * Every step logs the status bar and the open SAP windows, so a failed run
#     can be diagnosed from the log alone.
# ==============================================================================

from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import time
from datetime import date
from pathlib import Path

import keyring
import win32com.client
import win32event
from dateutil.relativedelta import relativedelta

log = logging.getLogger("sap_export")

# ── Configuration (placeholders) ──────────────────────────────────────────────
SAP_CONNECTION = "YOUR_SAP_CONNECTION"
SAP_CLIENT = "YOUR_CLIENT"
SAP_USER = "YOUR_SAP_USER"
KEYRING_SERVICE = "YOUR_SAP_KEYRING"          # python -m keyring set YOUR_SAP_KEYRING YOUR_SAP_USER
SAP_MUTEX_NAME = "Global\\SAP_Automation_Lock"  # shared by every script that drives SAP

COMPANY_CODE = "YOUR_COMPANY_CODE"
VENDOR_RANGES = ["YOUR_VENDOR_RANGE_1*", "YOUR_VENDOR_RANGE_2*"]
FBL1N_LAYOUT = "/YOUR_LAYOUT"
TABLE_VIEWER_TCODE = "YOUR_TABLE_VIEWER"       # custom table display transaction

EXPORT_DIR = Path(os.environ["LOCALAPPDATA"]) / "Temp" / "sap_export"   # ASCII only
LOAD_DELAY = 30     # seconds to let a large grid settle before exporting
STEP_DELAY = 5      # seconds before each export / save click

MULTISEL = ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/"
            "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I")


def fmt_duration(seconds: float) -> str:
    m, s = divmod(int(round(seconds)), 60)
    return f"{m}m {s}s" if m else f"{s}s"


# ── SAP session ───────────────────────────────────────────────────────────────

def close_sap() -> None:
    for proc in ("saplogon.exe", "saplgpad.exe", "sapgui.exe", "nwbc.exe"):
        subprocess.call(["taskkill", "/F", "/T", "/IM", proc],
                        stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)


def sap_wait(session, timeout: float = 300) -> None:
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not session.Busy:
                return
        except Exception:
            return
        time.sleep(0.5)


def sap_login():
    password = keyring.get_password(KEYRING_SERVICE, SAP_USER)
    if not password:
        raise RuntimeError(f"No password in Credential Manager: python -m keyring set {KEYRING_SERVICE} {SAP_USER}")
    close_sap()
    time.sleep(2)
    subprocess.Popen('start "" saplogon.exe', shell=True)
    time.sleep(3)
    gui = None
    for _ in range(30):
        try:
            gui = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            time.sleep(1)
    if gui is None:
        raise RuntimeError("SAP GUI Scripting engine not found (timeout).")
    session = gui.GetScriptingEngine.OpenConnection(SAP_CONNECTION, True).Children(0)
    time.sleep(1)
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = SAP_CLIENT
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
    session.findById("wnd[0]").sendVKey(0)
    sap_wait(session)
    sbar = session.findById("wnd[0]/sbar")
    if sbar.MessageType in ("E", "A"):
        raise RuntimeError(f"SAP login: {sbar.Text}")
    session.findById("wnd[0]").maximize()
    return session


def sap_tcode(session, code: str) -> None:
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{code}"
    session.findById("wnd[0]").sendVKey(0)
    sap_wait(session)


def sap_state(session, step: str) -> None:
    """Log what SAP shows: status bar, open windows, and the path/file fields as SAP reads them back."""
    try:
        sbar = session.findById("wnd[0]/sbar")
        log.debug(f"  [{step}] wnd[0]='{session.findById('wnd[0]').Text}' | sbar {sbar.MessageType}: '{sbar.Text}'")
    except Exception as e:
        log.debug(f"  [{step}] SAP not responding: {e}")
        return
    for n in (1, 2):
        try:
            w = session.findById(f"wnd[{n}]")
        except Exception:
            continue
        extra = ""
        for fld in ("DY_PATH", "DY_FILENAME"):
            try:
                extra += f" | {fld}='{session.findById(f'wnd[{n}]/usr/ctxt{fld}').text}'"
            except Exception:
                pass
        log.debug(f"  [{step}] wnd[{n}]='{w.Text}'{extra}")


# ── Export helpers ────────────────────────────────────────────────────────────

def wait_grid_loaded(session, grid_id: str, title: str, timeout: float = 600) -> None:
    """Large tables arrive in packages: wait until RowCount reaches the 'Rows N' of the window title."""
    m = re.search(r"(?:Rows|Γραμμές)\s+([\d.,]+)", title)
    expected = int(re.sub(r"[.,]", "", m.group(1))) if m else None
    t = time.time()
    while time.time() - t < timeout:
        sap_wait(session, timeout)
        try:
            rows = session.findById(grid_id).RowCount
        except Exception:
            rows = None
        if expected is None or (rows is not None and rows >= expected):
            log.info(f"  grid loaded: {rows} / {expected} rows in {fmt_duration(time.time() - t)}")
            return
        time.sleep(1)
    log.warning(f"  grid did not reach {expected} rows in {fmt_duration(timeout)} — continuing")


def save_dialog(session, final_path: Path, sap_name: str, precreate: bool = False, timeout: float = 300) -> None:
    """Fill SAP's Directory/File dialog with an ASCII temp target, wait for a NEW file, move it to final_path.
    The old final file is replaced only after a successful export."""
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = EXPORT_DIR / sap_name
    path.unlink(missing_ok=True)
    if precreate:
        from openpyxl import Workbook
        Workbook().save(path)               # the custom export only overwrites existing files
    before = (path.stat().st_mtime, path.stat().st_size) if path.exists() else None

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(EXPORT_DIR)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sap_name
    sap_state(session, "dialog filled")
    time.sleep(STEP_DELAY)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sap_wait(session, timeout)
    sap_state(session, "after Continue")
    for _ in range(2):                      # dialog still open -> SAP wants a second confirmation
        try:
            session.findById("wnd[1]/usr/ctxtDY_FILENAME")
        except Exception:
            break
        time.sleep(STEP_DELAY)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        sap_wait(session, timeout)

    def written() -> bool:
        if not path.exists():
            return False
        st = path.stat()
        if before and (st.st_mtime, st.st_size) == before:
            return False
        time.sleep(2)
        return path.stat().st_size == st.st_size      # finished writing

    t = time.time()
    while not written():
        if time.time() - t > timeout:
            sap_state(session, "file not found")
            raise RuntimeError(f"SAP export did not produce {sap_name}")
        time.sleep(1)
    log.debug(f"  {sap_name}: {path.stat().st_size:,} bytes in {fmt_duration(time.time() - t)}")
    final_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(path), str(final_path))


def export_suppliers(session, target: Path) -> None:
    """Supplier master (LFA1) via the custom table viewer, display mode only."""
    sap_tcode(session, TABLE_VIEWER_TCODE)
    session.findById("wnd[0]/usr/ctxtTABNAME").text = "LFA1"
    session.findById("wnd[0]/usr/chkDISP").selected = True         # display only
    session.findById("wnd[0]/tbar[1]/btn[8]").press()               # Execute
    sap_wait(session, 600)
    grid = "wnd[0]/usr/shell/shellcont[1]/shell"
    wait_grid_loaded(session, grid, session.findById("wnd[0]").Text)
    time.sleep(LOAD_DELAY)
    t = time.time()
    session.findById(grid).pressToolbarButton("EXPR")              # export to Excel
    sap_wait(session, 600)
    log.info(f"  export dialog opened in {fmt_duration(time.time() - t)}")
    time.sleep(LOAD_DELAY)
    save_dialog(session, target, "LFA1.xlsx", precreate=True)
    log.info(f"Saved: {target.name}")


def export_line_items(session, target: Path) -> None:
    """FBL1N: vendor ranges (multiple selection), company code, all items,
    posting dates from 1 Jan of (today - 7 months) to today, saved layout."""
    year = (date.today() - relativedelta(months=7)).year
    date_from, date_to = f"01.01.{year}", date.today().strftime("%d.%m.%Y")
    sap_tcode(session, "FBL1N")
    session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").press()   # multiple selection
    sap_wait(session)
    for i, v in enumerate(VENDOR_RANGES):
        session.findById(f"{MULTISEL}[1,{i}]").text = v
    session.findById("wnd[1]/tbar[0]/btn[8]").press()                        # Copy
    sap_wait(session)
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = COMPANY_CODE
    session.findById("wnd[0]/usr/radX_AISEL").select()                        # all items
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = date_to
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = FBL1N_LAYOUT
    session.findById("wnd[0]/tbar[1]/btn[8]").press()                        # Execute
    sap_wait(session, 600)
    log.info(f"FBL1N {COMPANY_CODE} {date_from} - {date_to}: {session.findById('wnd[0]/sbar').Text}")
    time.sleep(STEP_DELAY)
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()        # List > Export > Spreadsheet
    sap_wait(session)
    time.sleep(STEP_DELAY)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()                        # format: XLSX
    sap_wait(session)
    save_dialog(session, target, "FBL1N.xlsx")
    log.info(f"Saved: {target.name}")


def run(suppliers_file: Path, line_items_file: Path, suppliers: bool = True) -> None:
    """One SAP login for both exports; SAP is always closed at the end."""
    t = time.time()
    mutex = win32event.CreateMutex(None, False, SAP_MUTEX_NAME)
    win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    session = None
    try:
        session = sap_login()
        if suppliers:
            export_suppliers(session, suppliers_file)
        export_line_items(session, line_items_file)
    except Exception:
        if session is not None:
            sap_state(session, "ERROR")
        raise
    finally:
        close_sap()
        win32event.ReleaseMutex(mutex)
    log.info(f"SAP closed — SAP time: {fmt_duration(time.time() - t)}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    log.setLevel(logging.DEBUG)
    out = Path(r"C:\Users\YOUR_USERNAME\Desktop\MATCHING REPORT")
    run(out / "SUPPLIERS.xlsx", out / "LINE_ITEMS.xlsx")
