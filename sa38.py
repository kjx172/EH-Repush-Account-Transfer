# file: sa38.py
# Purpose: Launch SA38 and run ZCUSTOMER_UPD_VALUES (Mass Account Transfers entry point)
# Requires: Windows, SAP GUI for Windows with Scripting enabled, pywin32

import sys
import time
import re
try:
    import win32com.client  # pywin32
    from pywintypes import com_error # pyright: ignore[reportMissingModuleSource]
except Exception:
    print("Missing dependency. Install pywin32: pip install pywin32")
    raise

# ---------- HELPERS ----------

def get_sap_session():
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

def _set_clipboard_text(text: str):
        try:
            import subprocess
            try:
                import win32clipboard # pyright: ignore[reportMissingModuleSource]
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(text)
                win32clipboard.CloseClipboard()
            except Exception:
                # Fallback via Windows 'clip'
                subprocess.run("clip", input=text, text=True, shell=True, check=True)
        except Exception as _e:
            raise RuntimeError(f"Failed setting clipboard text: {_e}")

# ---------- UTILS ----------
def open_tx(session):
    # Go to SA38
    session.findById("wnd[0]/tbar[0]/okcd").text = "sa38"
    session.findById("wnd[0]").sendVKey(0)  # Enter

    # Enter program name and execute
    prog = session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM")
    prog.text = "ZCUSTOMER_UPD_VALUES"
    try:
        prog.caretPosition = len(prog.text)
    except Exception:
        pass

    # Execute (F8)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

def collect_sap_ve_sets():
    """
    Collect multiple sets of SAP account numbers paired with a VE number.
    Input rules:
      - SAP numbers: numeric only; non-numeric entries are ignored.
      - VE number: must be numeric; user is re-prompted until valid.
      - 'v' moves to VE entry; 'd' finishes all sets.
    """

    sets = []

    print("SAP Numbers & VE # pairings (press 'd' when done with ALL sets):")

    while True:
        sap_numbers = []
        print("\n⚠️  Enter SAP numbers: (enter v to proceed to VE#, d to complete)")

        # Collect SAP numbers until 'v' or 'd'
        while True:
            line = input().strip()

            if line.lower() == "d":   # done with all sets
                print("✅  SAP & VE# sets saved successfully")
                return sets

            if line.lower() == "v":   # move on to VE entry
                break

            # numeric check for SAP numbers
            if line.isnumeric():
                sap_numbers.append(line)
            else:
                if line:   # ignore empty; warn only on non-empty invalid
                    print(f"❌  '{line}' ignored (SAP numbers must be numeric).")

        # ---- Collect VE number (must be numeric) ----
        ve = ""
        while True:
            print("⚠️  Enter VE # (numeric only):")
            ve = input().strip()

            if ve.isnumeric():
                break

            print(f"❌  '{ve}' is not numeric. Please enter VE # again.")

        # Store set
        sets.append({
            "sap_numbers": sap_numbers,
            "ve_number": ve
        })

        print(f"✅  {len(sap_numbers)} Accounts will be assigned to VE # {ve}")
        print("---")

def loop_enter_account_transfer(account_ve_sets, session):
    """
    Loop through SAP account → VE sets and execute the mass transfer flow for each set.

    Expected structure per set:
      {
        "sap_numbers": ["0046189407", "0046189625", ...],
        "ve_number": "46213203"
      }

    Side effects per set:
      - Opens multi-select for KUNNR, pastes SAP numbers from clipboard
      - Sets VE field, executes
      - Selects all rows in result grid, presses CREATE_BATCH, confirms, then Back
    """
    # Constant IDs
    btn_kunnr_multisel = (
        "wnd[0]/usr/tabsTABSTRIP_TABSTRIPS/tabpTAB1/"
        "ssub%_SUBSCREEN_TABSTRIPS:ZCUSTOMER_UPD_VALUES:1100/btn%_S_KUNNR_%_APP_%-VALU_PUSH"
    )
    fld_ve = (
        "wnd[0]/usr/tabsTABSTRIP_TABSTRIPS/tabpTAB1/"
        "ssub%_SUBSCREEN_TABSTRIPS:ZCUSTOMER_UPD_VALUES:1100/ctxtP_VE"
    )
    grid_id = "wnd[0]/usr/cntlCC1/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell"

    # Stores batch input name
    messages = []

    # loop
    total = len(account_ve_sets)
    for idx, item in enumerate(account_ve_sets, start=1):
        sap_numbers = (item.get("sap_numbers") or [])
        ve_number = str(item.get("ve_number") or "").strip()

        # 1) Open multi-select for KUNNR and paste numbers from clipboard
        if not sap_numbers:
            print(f"[{idx}/{total}] Skipped: no SAP numbers in this set.")
            continue

        # Put numbers on clipboard (newline separation works best for SAP multi-select)
        _set_clipboard_text("\r\n".join(sap_numbers))

        session.findById(btn_kunnr_multisel).press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()    # Clear selections
        #time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[24]").press()  # Paste from Clipboard
        session.findById("wnd[1]/tbar[0]/btn[8]").press()   # OK

        # 2) Enter VE number
        if not ve_number:
            raise ValueError(f"Missing VE number for set #{idx}.")

        ve = session.findById(fld_ve)
        ve.text = ve_number
        try:
            ve.setFocus()
        except Exception:
            pass
        try:
            ve.caretPosition = len(ve_number)
        except Exception:
            pass

        # 3) Execute
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 4) In result grid: clear current cell, select all, CREATE_BATCH
        try:
            grid = session.findById(grid_id)
            try:
                grid.setCurrentCell(-1, "")
            except Exception:
                pass
            # Some environments expose selectAll as a method; others as an action.
            try:
                grid.selectAll()
            except Exception:
                # If method invocation fails, try without parens (rare)
                try:
                    grid.selectAll
                except Exception:
                    pass
            grid.pressToolbarButton("CREATE_BATCH")
            #_wait_ready(session)
        except Exception:
            # No grid / different layout / empty result — not fatal for the loop
            pass

        
        # 5) Capture message text
        message_text = None
        try:
            # Read the message from the popup text field
            message_text = session.findById("wnd[1]/usr/txtMESSTXT1").Text
            messages.append(message_text)
        except Exception:
            # No message field found; leave as None
            pass

        # 6) Confirm popup and go Back
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()  # e.g., "Continue/OK" on confirmation
        except Exception:
            pass

        try:
            session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Back
        except Exception:
            pass

        print(f"[{idx}/{total}] {len(sap_numbers)} account(s) queued for VE #{ve_number}")

    print("Completed Batch Inputs")
    return messages

def batch_input_monitoring(messages: list[str], session) -> list[str]:
    """
    Navigate to SM35 and return only the substrings enclosed in double quotes
    from the provided `batch_names` lines.

    Example:
        Input line:  'Batch input with name "ZCUP-011123" created.'
        Output item: 'ZCUP-011123'
    """
    # --- 1) Enter Batch Input Monitoring (SM35) ---
    try:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Back
    except Exception:
        pass
    try:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Back again
    except Exception:
        pass

    session.findById("wnd[0]/tbar[0]/okcd").text = "sm35"
    session.findById("wnd[0]").sendVKey(0)  # Enter

    # --- 2) Extract quoted names from input lines ---
    batch_names: list[str] = []
    for message in messages or []:
        # Find all segments inside double quotes
        # e.g., '... "ABC" ... "DEF"' -> ['ABC', 'DEF']
        matches = re.findall(r'"([^"]+)"', str(message))
        if matches:
            batch_names.extend(m.strip() for m in matches if m.strip())

    return batch_names

def execute_batches(batch_names, session):
    """
    Execute each batch input name in SM35.
    Follows the SAP GUI flow you provided and handles missing popup buttons
    by allowing user interaction to continue.
    """

    # ID constants for readability
    fld_batch = "wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-MAPN"
    tbl_row_sel = (
        "wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/"
        "ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI"
    )
    cell_groupid = (
        tbl_row_sel + "/txtITAB_APQI-GROUPID[0,0]"
    )

    for bname in batch_names:
        # 1) Enter "batchname" into SM35 selection box
        session.findById(fld_batch).text = f"{bname}"
        try:
            session.findById(fld_batch).caretPosition = len(bname) + 1
        except Exception:
            pass

        session.findById("wnd[0]").sendVKey(0)  # ENTER

        # 2) Select first row in the APQI table
        try:
            table = session.findById(tbl_row_sel)
            table.getAbsoluteRow(0).selected = True

            cell = session.findById(cell_groupid)
            cell.setFocus()
            try:
                cell.caretPosition = 0
            except Exception:
                pass
        except Exception:
            print(f"⚠️  No table row found for batch '{bname}'. Skipping.")
            continue

        # 3) Press Execute
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 4) Select "ERROR only" radio button in popup
        try:
            rad_err = session.findById("wnd[1]/usr/radD0300-ERROR")
            rad_err.select()
            rad_err.setFocus()
        except Exception:
            print(f"⚠️  Could not find error/radio button popup for '{bname}'.")
            # allow user to decide whether to control SAP manually
            input("Press ENTER to continue to next batch...")
            continue

        # 5) Press OK buttons twice
        for _ in range(2):
            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:
                print(f"⚠️  Popup close button missing for '{bname}'.")
                input("Press ENTER to continue (in case of real SAP error)...")
                break

        print(f"✅  Executed batch '{bname}'")

# ---------- MAIN ----------
def run_flow():
    session = get_sap_session()

    # 1. Go to SA38 & ZCUSTOMER_UPD_VALUES
    open_tx(session)

    # 2. Get Account # & VE # sets
    account_ve_sets = collect_sap_ve_sets()
    #account_ve_sets = [{'sap_numbers': ['0046189407', '0046189625', '0046192356', '0046193493', '0046194163', '0046194253', '0046197208'], 've_number': '46213203'}, {'sap_numbers': ['0046224213', '0046192759', '0046189339', '0046189354', '0046189389'], 've_number': '46214310'}]

    # 3. Enter into sa38
    messages = loop_enter_account_transfer(account_ve_sets, session)

    # 4. Go to batch input monitoring & filter messages
    batch_names = batch_input_monitoring(messages, session)

    # 5. Loop through batch names and execute the batches.
    execute_batches(batch_names, session)