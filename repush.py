# file: send_to_rep_system.py

import sys
import re
import time
import subprocess
import msvcrt
import win32api  # pyright: ignore[reportMissingModuleSource]

try:
    import win32com.client  # pywin32
    import win32clipboard # pyright: ignore[reportMissingModuleSource]
    import win32con # pyright: ignore[reportMissingModuleSource]
    from pywintypes import com_error # pyright: ignore[reportMissingModuleSource]
except Exception as ex:
    print("Missing dependency. Ensure you're on Windows and run: pip install pywin32")
    raise

# ---------- Helpers ----------

def normalize_numbers(raw_text: str) -> list:
    """
    Accepts mixed separators (comma, semicolon, space, newline).
    Keeps only non-empty numeric tokens;
    """
    tokens = re.split(r"[,\s;]+", raw_text.strip())
    tokens = [t for t in tokens if t]
    # Keep only digits
    tokens = [t for t in tokens if t.isdigit()]
    return tokens

def set_clipboard_text(text: str) -> None:
    """Set Windows clipboard text. Falls back to `clip` if needed."""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
    except Exception:
        # Fallback using Windows 'clip' command
        subprocess.run("clip", input=text, text=True, shell=True, check=True)

def get_sap_session():
    """
    Get the first active SAP GUI session from the first connection.
    Adjust indices if you use multiple connections or sessions.
    """
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
    except Exception:
        raise RuntimeError("SAP GUI is not running or Scripting is disabled on the client.")
    try:
        application = sapgui.GetScriptingEngine
        if application is None or application.Children.Count == 0:
            raise RuntimeError("No SAP GUI connections found. Log on and try again.")
        connection = application.Children(0)
        if connection.Children.Count == 0:
            raise RuntimeError("No active SAP sessions found in the first connection.")
        session = connection.Children(0)
        return session
    except com_error as ce:
        raise RuntimeError("Unable to access SAP Scripting engine (check server + client settings).") from ce

def _wait_until_ready(session, timeout=10.0, poll=0.1):
    """Wait until modal dialog (wnd[1]) is gone and allow a short settle."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            session.findById("wnd[1]")  # modal still present
            time.sleep(poll)
            continue
        except Exception:
            time.sleep(poll)  # short settle after modal closes
            return True
    return False

def read_numbers_interactive() -> list[str]:
    """
    Interactive input:
      - Press Enter to commit the current line and continue
      - Press Ctrl+Enter to finish input
      - Supports pasting multiple lines
    """
    print("Enter order numbers (one per line).")
    print("Press Ctrl+Enter when DONE.")
    print("-" * 60)

    lines: list[str] = []
    current = []

    def flush_line():
        nonlocal current
        lines.append("".join(current))
        current = []

    while True:
        ch = msvcrt.getwch()  # wide-char read, no echo by default

        if ch == "\r":  # Enter pressed
            # check modifier keys at the moment of Enter
            ctrl_down = win32api.GetKeyState(win32con.VK_CONTROL) < 0

            if ctrl_down:
                # finish: commit any partially typed content first (if non-empty)
                if current:
                    flush_line()
                print()  # move to next console line
                break
            else:
                flush_line()
                print()  # echo newline
                continue

        elif ch in ("\x08",):  # Backspace
            if current:
                current.pop()
                # move cursor back, erase last char visually
                print("\b \b", end="", flush=True)
            continue

        elif ch == "\x03":  # Ctrl+C
            raise KeyboardInterrupt

        else:
            # Regular character (including pasted content). Handle '\n' from paste explicitly.
            if ch == "\n":
                # Treat newline from paste as Enter (continue)
                flush_line()
                print()
            else:
                current.append(ch)
                print(ch, end="", flush=True)

    return normalize_numbers("\n".join(lines))

def is_no_rows_error(ex: Exception) -> bool:
    """
    Detect SAP GUI 'control not found' error that usually indicates
    no ALV grid / no eligible rows to process.
    """
    msg = str(ex)
    return (
        "The control could not be found by id" in msg
        or "findById" in msg and "not found" in msg
    )

# ---------- Core automation ----------

def run_flow(session, so_numbers: list[str]):
    # 1) Copy numbers to clipboard with newline separation (best for SAP multi-select paste)
    set_clipboard_text("\r\n".join(so_numbers))

    # 2) Go to transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZREP_INTEGR_MONI"
    session.findById("wnd[0]").sendVKey(0)  # Enter

    # 3) Set checkboxes as requested
    #    - Uncheck Billing Documents (chkP_BD)
    #    - Uncheck Customer/Contacts (chkP_CUS)
    #    - Check Last Attempt Only (chkPA_LAST)
    for elt in (
        ("wnd[0]/usr/chkP_BD", False),
        ("wnd[0]/usr/chkP_CUS", False),
        ("wnd[0]/usr/chkPA_LAST", True),
    ):
        obj_id, should_select = elt
        ctrl = session.findById(obj_id)
        ctrl.setFocus()
        ctrl.selected = bool(should_select)

    # 4) Open Multiple Selection for SO (VBELN)
    session.findById("wnd[0]/usr/btn%_SO_VBELN_%_APP_%-VALU_PUSH").press()

    # 5) In the popup:
    # btn[24] = "Paste from Clipboard"
    # btn[8]  = "OK"
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Ensure popup is closed and main window is ready
    _wait_until_ready(session)

    try:
        # Execute (F8)
        session.findById("wnd[0]").sendVKey(8)

        # Try to load ALV grid
        grid_id = "wnd[0]/usr/cntlCC1/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell"
        grid = None
        for _ in range(50):
            try:
                grid = session.findById(grid_id)
                break
            except Exception:
                time.sleep(0.1)

        if grid is None:
            raise RuntimeError("No results grid — no eligible rows to repush.")

        # Select row and send
        grid.selectedRows = "0"
        grid.pressToolbarButton("SEND_PO")

    except Exception as ex:
        if is_no_rows_error(ex):
            print("No eligible rows to repush.")
            return
        else:
            raise

def main():
    # --- replace the input-gathering block in main() ---
    print("-" * 60)

    if sys.stdin.isatty():
        so_numbers = read_numbers_interactive()
    else:
        # still support piped input or redirected files
        raw = sys.stdin.read()
        so_numbers = normalize_numbers(raw)

    if not so_numbers:
        print("No usable order numbers were provided. Exiting.")
        sys.exit(1)

    print(f"Parsed {len(so_numbers)} entries. Copying to clipboard and automating SAP...")

    try:
        session = get_sap_session()
        run_flow(session, so_numbers)
        print("Done. If there were qualifying rows, 'Send to Rep System' was triggered.")
    except Exception as ex:
        print(f"[ERROR] {ex}")
        sys.exit(2)

if __name__ == "__main__":
    main()