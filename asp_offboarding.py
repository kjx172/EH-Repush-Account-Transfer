# file: asp_offboarding.py

import sys

try:
    import win32com.client  # pywin32
    from pywintypes import com_error  # type: ignore
except Exception:
    print("Missing dependency. Install pywin32: pip install pywin32")
    raise

# --------- HELPERS ---------

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

# --------- UTILS ---------

def _safe_get_text(ctrl):
    """Return .text or .Text from an SAP field without raising."""
    try:
        return str(ctrl.text)
    except Exception:
        try:
            return str(ctrl.Text)
        except Exception:
            return ""

def _safe_set_text(ctrl, value: str):
    try:
        ctrl.text = value
    except Exception:
        try:
            ctrl.Text = value
        except Exception as e:
            raise

def _open_pa30(session):
    # Resize (best-effort; don't fail flow if not supported on environment)
    try:
        session.findById("wnd[0]").resizeWorkingPane(155, 43, False)
    except Exception:
        pass
    # Go to PA30
    session.findById("wnd[0]/tbar[0]/okcd").text = "pa30"
    session.findById("wnd[0]").sendVKey(0)  # Enter

def _enter_personnel_number(session, pernr: str):
    fld = session.findById("wnd[0]/usr/ctxtRP50G-PERNR")
    _safe_set_text(fld, pernr)
    try:
        fld.setFocus()
    except Exception:
        pass
    try:
        fld.caretPosition = len(pernr)
    except Exception:
        pass
    session.findById("wnd[0]").sendVKey(0)

def _select_menu_row_and_create(session):
    # Select the 2nd row in the PA30 IT menu, focus the 2nd row text, then press the toolbar create/display button
    try:
        tbl_path = (
            "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/"
            "ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/"
            "tblSAPMP50ATC_MENU"
        )
        session.findById(f"{tbl_path}").getAbsoluteRow(1).selected = True
        cell = session.findById(
            f"{tbl_path}/txtGV_ITEXT[0,1]"
        )
        try:
            cell.setFocus()
        except Exception:
            pass
        try:
            cell.caretPosition = 0
        except Exception:
            pass
    except Exception:
        # Proceed even if specific table/cell is not present in the user layout
        pass
    # Press changw button
    session.findById("wnd[0]/tbar[1]/btn[6]").press()

def _prefix_first_name_with_zzz(session):
    # Read current first name (P0002-VORNA), set to ZZZ_<current>, avoiding double prefix
    fld = session.findById("wnd[0]/usr/txtP0002-VORNA")
    current = _safe_get_text(fld).strip()
    new_value = current if current.upper().startswith("ZZZ_") else f"ZZZ_{current}"
    _safe_set_text(fld, new_value)
    try:
        fld.setFocus()
    except Exception:
        pass
    try:
        fld.caretPosition = len(new_value)
    except Exception:
        pass
    # Save (toolbar 0, button 11 per user's flow)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

def _close_services_tab_for_pernr(session, pernr: str):
    """Implements flow for ZVZSERVTABN."""
    import datetime

    try:
        session.findById("wnd[0]").resizeWorkingPane(155, 43, False)
    except Exception:
        pass

    try:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
    except Exception:
        pass

    session.findById("wnd[0]/tbar[0]/okcd").text = "ZVZSERVTABN"
    session.findById("wnd[0]").sendVKey(0)

    try:
        session.findById("wnd[0]/usr/ctxtS_IWERK-LOW").text = "4610"
    except Exception:
        pass

    pernr_field = session.findById("wnd[0]/usr/ctxtS_PERNR-LOW")
    _safe_set_text(pernr_field, pernr)
    try:
        pernr_field.setFocus()
    except Exception:
        pass
    try:
        pernr_field.caretPosition = 8
    except Exception:
        pass

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    today = datetime.datetime.now().strftime("%m/%d/%Y")
    grid = session.findById("wnd[0]/usr/cntlALV_CONTAINER/shellcont/shell")
    try:
        grid.modifyCell(0, "DATBI", today)
    except Exception:
        try:
            grid.setCurrentCell(0, "DATBI")
        except Exception:
            pass
    try:
        grid.currentCellColumn = "DATBI"
    except Exception:
        pass

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    try:
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    except Exception:
        pass

# --------- MAIN ---------

def run_flow(personnel_number: str | None = None):
    """
    Execute ASP Offboarding flow in PA30 for a given personnel number.

    If `personnel_number` is None, the user will be prompted to enter it.
    """
    session = get_sap_session()

    if not personnel_number:
        pernr = input("Enter ASP VE Number: ").strip()
    else:
        pernr = str(personnel_number).strip()

    if not pernr:
        print("❌ No personnel number provided. Aborting.")
        return

    if not pernr.isdigit():
        print(f"⚠️  Personnel number '{pernr}' contains non-digits")
        return

    _open_pa30(session)
    _enter_personnel_number(session, pernr)
    _select_menu_row_and_create(session)
    _prefix_first_name_with_zzz(session)
    _close_services_tab_for_pernr(session, pernr)

    print(f"✅ ASP Offboarding completed for PERNR {pernr}")

