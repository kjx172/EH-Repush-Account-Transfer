# file: repush.py
import sys
import re
import time
import subprocess
import msvcrt
import win32api
import win32con
try:
    import win32com.client  # pywin32
    import win32clipboard # pyright: ignore[reportMissingModuleSource]
    from pywintypes import com_error # pyright: ignore[reportMissingModuleSource]
except Exception:
    print("Missing dependency. Ensure you're on Windows and run: pip install pywin32")
    raise

# ---------- HELPERS ----------

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
    """Put text onto the Windows clipboard; fallback to 'clip'."""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
    except Exception:
        subprocess.run("clip", input=text, text=True, shell=True, check=True)

def get_sap_session():
    """Return the first session of the first connection (adjust indices if needed)."""
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
    except Exception:
        raise RuntimeError("SAP GUI is not running or scripting is disabled.")
    try:
        app = sapgui.GetScriptingEngine
        if app is None or app.Children.Count == 0:
            raise RuntimeError("No SAP GUI connections found.")
        conn = app.Children(0)
        if conn.Children.Count == 0:
            raise RuntimeError("No active SAP sessions found in the first connection.")
        return conn.Children(0)
    except com_error as ce:
        raise RuntimeError("Unable to access SAP Scripting engine (server/client).") from ce

# ---------- UTILS ----------

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
    current: list[str] = []

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

DOC_FLOW = {
    # Orders: uncheck Billing + Customer; check Last Attempt; open SO_VBELN
    "order": {
        "checkboxes": [
            ("wnd[0]/usr/chkP_BD",   False),  # Billing docs OFF
            ("wnd[0]/usr/chkP_CUS",  False),  # Customer/contacts OFF
            ("wnd[0]/usr/chkPA_LAST", True),  # Last Attempt ON
        ],
        "multi_select_btn": "wnd[0]/usr/btn%_SO_VBELN_%_APP_%-VALU_PUSH",
    },
    # Invoices: uncheck Sales + Customer; check Last Attempt; open SO_BD
    "invoice": {
        "checkboxes": [
            ("wnd[0]/usr/chkP_SD",   False),  # Sales OFF
            ("wnd[0]/usr/chkP_CUS",  False),  # Customer/contacts OFF
            ("wnd[0]/usr/chkPA_LAST", True),  # Last Attempt ON
        ],
        "multi_select_btn": "wnd[0]/usr/btn%_SO_BD_%_APP_%-VALU_PUSH",
    },
}

def open_tx_and_apply_criteria(session, doc_type: str, numbers: list[str]):
    """Go to ZREP_INTEGR_MONI, set relevant checkboxes, and paste numbers."""
    cfg = DOC_FLOW[doc_type]

    # Put numbers on clipboard (newline-separated works best in SAP multi-select)
    set_clipboard_text("\r\n".join(numbers))

    # Transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZREP_INTEGR_MONI"
    session.findById("wnd[0]").sendVKey(0)  # Enter

    # Checkboxes
    for obj_id, selected in cfg["checkboxes"]:
        ctrl = session.findById(obj_id)
        ctrl.setFocus()
        ctrl.selected = bool(selected)

    # Multiple selection popup for the correct field
    session.findById(cfg["multi_select_btn"]).press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()  # Paste from clipboard
    session.findById("wnd[1]/tbar[0]/btn[8]").press()   # OK

def execute_and_finish(session):
    """Execute and send to rep system (Select ALL rows)."""
    # Execute (F8)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    try:
        # Use your new flow
        session.findById("wnd[0]").resizeWorkingPane(155, 43, False)

        grid_id = (
            "wnd[0]/usr/cntlCC1/shellcont/shell/"
            "shellcont[1]/shell/shellcont[0]/shell"
        )
        grid = session.findById(grid_id)

        # Clear active cell and select all rows
        grid.setCurrentCell(-1, "")
        grid.selectAll()

        # Send to Rep System
        grid.pressToolbarButton("SEND_PO")

    except Exception:
        # Grid wasn't found after Execute → no results/no screen change
        print("⚠️  No grid detected after Execute — No successful repushes")

# ---------- MAIN ----------
def run_flow(doc_type: str):
    """
    Continuously collect numbers interactively and execute the flow for each batch.
    Stops when the user submits an empty batch (i.e., immediately Ctrl+Enter) or answers 'n' at the prompt.
    """
    if doc_type not in DOC_FLOW:
        raise ValueError(f"Unsupported doc_type: {doc_type}. Use 'order' or 'invoice'.")

    while True:
        print(f"\n=== Repush {doc_type} Batch ===")
        try:
            numbers = read_numbers_interactive()
        except KeyboardInterrupt:
            print("\nCancelled by user (Ctrl+C). Exiting.")
            break

        if not numbers:
            print("No numbers entered. Exiting batch loop.")
            break

        session = get_sap_session()
        
        # Run the shared flow for this batch
        open_tx_and_apply_criteria(session, doc_type, numbers)
        execute_and_finish(session)
        print(f"\n=== Repush {doc_type} Batch Complete ===")

        return