# ise_ose.py
from config import AppConfig
from get_ticket import sanitize_ticket_input, get_ticket_core_fields
import time
CFG = AppConfig.load()

try:
    import win32com.client  # type: ignore
except Exception as e:
    win32com = None

def _yes_no(prompt: str, default_yes: bool = True) -> bool:
    suffix = "[Y/n]" if default_yes else "[y/N]"
    while True:
        ans = input(f"{prompt} {suffix}: ").strip().lower()
        if not ans:
            return default_yes
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("Please answer y/yes or n/no.")

class SapGuiError(Exception):
    pass

class SapGui:
    # ----------------------------
    # HELPERS
    # ----------------------------
    def __init__(self):
        if win32com is None:
            raise SapGuiError(
                "pywin32 not available. Run on Windows with SAP GUI installed and 'pip install pywin32'."
            )
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        self.application = self.SapGuiAuto.GetScriptingEngine
        # Reuse the first active session; if none, raise
        if self.application.Connections.Count == 0:
            raise SapGuiError("No active SAP GUI connection found. Log in to a system first.")
        self.connection = self.application.Connections(0)
        if self.connection.Children.Count == 0:
            raise SapGuiError("No active SAP GUI session found. Open a session and try again.")
        self.session = self.connection.Children(0)
        self._wait_control("wnd[0]", timeout=20)
        self._wait_control("wnd[0]/tbar[0]/okcd", timeout=20)

    def _wait_control(self, id_: str, timeout: float = 10.0, interval: float = 0.1) -> None:
        end = time.time() + timeout
        while time.time() < end:
            try:
                self.session.findById(id_)
                return
            except Exception:
                time.sleep(interval)
        raise SapGuiError(f"Timeout waiting for control: {id_}")

    def _exec(self) -> None:
        self.session.findById("wnd[0]").sendVKey(0)

    # Transform VE number so that its first digit is replaced with '5'
    def _to_vendor_number(self, ve_number: str) -> str:
        if not ve_number:
            raise SapGuiError("VE number is empty.")
        
        s = str(ve_number)
        # Replace the first character with '5'
        return "5" + s[1:] if len(s) > 1 else "5"
    
    # ----------------------------
    # UTILS
    # ----------------------------
    # start transaction
    def start_tx(self, code: str) -> None:
        self._wait_control("wnd[0]/tbar[0]/okcd", timeout=20)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{code.upper()}"
        self._exec()

    # ZREP_VENDORS
    def run_report(self, progname: str) -> None:
        # assumes we're on SA38 initial screen
        self.session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = progname
        # place caret for parity with your recording
        try:
            self.session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = len(progname)
        except Exception:
            pass
        self._exec()
        # Press Execute
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

    #Zendesk rep company -> table key via config
    def resolve_rep_table_key(self, rep_company: str | None) -> str:
        mapping = CFG.sap.get("rep_table_key_map", {})

        map_key = (rep_company or "").strip().lower()
        table_key = mapping.get(map_key, mapping.get("otherrep", "Other"))

        return table_key

    # Open rep company on table
    def filter_rep_company_and_open(self, table_key: str) -> None:
        shell_id = "wnd[0]/usr/cntlC_REP_COMP/shellcont/shell"
        self._wait_control(shell_id, timeout=20)

        # NEW: clear current row and select the RCKEY column before filtering
        grid = self.session.findById(shell_id)
        try:
            grid.currentCellRow = -1             # ensure no row context
        except Exception:
            pass
        grid.selectColumn("RCKEY")               # focus the Rep Company key column

        # Open filter dialog
        grid.pressToolbarButton("&MB_FILTER")

        # Enter table key into LOW field and accept
        low_id_primary = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%_%DYN001-LOW"
        low_id_fallback = "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"
        try:
            #self._wait_control(low_id_primary, timeout=10)
            self.session.findById(low_id_primary).text = table_key
            try:
                self.session.findById(low_id_primary).caretPosition = len(table_key)
            except Exception:
                pass
        except Exception:
            #self._wait_control(low_id_fallback, timeout=10)
            self.session.findById(low_id_fallback).text = table_key
            try:
                self.session.findById(low_id_fallback).caretPosition = len(table_key)
            except Exception:
                pass

        self.session.findById("wnd[1]").sendVKey(0)  # OK/Enter

        # Back on the table: ensure focus and double‑click the (filtered) current cell
        self._wait_control(shell_id, timeout=10)
        self.session.findById(shell_id).doubleClickCurrentCell()

    # Delete E # in 3rd table
    def delete_Enumber(self, enumber: str) -> None:
        if not enumber:
            raise SapGuiError("E-number is required")

        main_grid_id = "wnd[0]/usr/cntlC_REP_USER/shellcont/shell"

        # Ensure main window is responsive and grid is present
        self._wait_control("wnd[0]", timeout=20)
        self._wait_control(main_grid_id, timeout=20)

        # 1) Resize working pane
        try:
            self.session.findById("wnd[0]").resizeWorkingPane(155, 43, False)
        except Exception:
            pass

        grid = self.session.findById(main_grid_id)

        # 2) Open Filter on the C_REP_USER grid
        try:
            grid.pressToolbarButton("&MB_FILTER")
        except Exception as e:
            raise SapGuiError(f"Could not open filter dialog on C_REP_USER grid: {e}")

        # ---- Filter dialog (wnd[1]) : pick field, move, open single-value entry ----
        # Select a field from left list (CONTAINER1_FILT) row 1 (0-based) = portal user id
        left_list = "wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell"
        right_list = "wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER2_FILT/shellcont/shell"

        self._wait_control(left_list, timeout=10)
        try:
            self.session.findById(left_list).currentCellRow = 1
            self.session.findById(left_list).selectedRows = "1"
        except Exception:
            # Fallback: try row 0 if row 1 is not available
            try:
                self.session.findById(left_list).currentCellRow = 0
                self.session.findById(left_list).selectedRows = "0"
            except Exception as e:
                raise SapGuiError(f"Unable to select a field in the left filter list: {e}")

        # Move selected field to the right with "single arrow" button
        try:
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press()
        except Exception as e:
            raise SapGuiError(f"Could not move field to the right list in filter dialog: {e}")

        # Select first item in right list and open "Multiple/Single" value dialog (btn600_BUTTON)
        self._wait_control(right_list, timeout=10)
        try:
            self.session.findById(right_list).selectedRows = "0"
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press()
        except Exception as e:
            raise SapGuiError(f"Could not open value entry for the selected filter field: {e}")

        # ---- Value entry popup (wnd[2]) : put E-number in LOW and confirm ----
        # Two common LOW field IDs (theme-dependent)
        low_primary = "wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%_DYN001-LOW"
        low_fallback = "wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"

        try:
            self._wait_control(low_primary, timeout=6)
            self.session.findById(low_primary).text = enumber
            try:
                self.session.findById(low_primary).caretPosition = len(enumber)
            except Exception:
                pass
        except Exception:
            # Use fallback ID if primary not found
            self._wait_control(low_fallback, timeout=6)
            self.session.findById(low_fallback).text = enumber
            try:
                self.session.findById(low_fallback).caretPosition = len(enumber)
            except Exception:
                pass

        # Confirm value entry (OK/Enter)
        try:
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        except Exception:
            self.session.findById("wnd[2]").sendVKey(0)

        # ---- Back on main grid: select filtered row 0 and delete ----
        self._wait_control(main_grid_id, timeout=15)
        grid = self.session.findById(main_grid_id)

        # Clear current cell context (some grids require this before selecting rows)
        try:
            grid.currentCellColumn = ""
        except Exception:
            pass

        # Select the first (filtered) row
        try:
            grid.selectedRows = "0"
        except Exception as e:
            raise SapGuiError(f"Could not select filtered row to delete: {e}")

        # Press Delete on the toolbar
        try:
            grid.pressToolbarButton("DEL")
        except Exception:
            # Some themes place 'Delete' on the application toolbar; fall back to menu key DEL
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[14]").press()  # heuristic; may vary
            except Exception as e:
                raise SapGuiError(f"Could not trigger delete for selected row: {e}")

        # Confirm deletion (Yes)
        try:
            self._wait_control("wnd[1]/usr/btnBUTTON_1", timeout=8)
            self.session.findById("wnd[1]/usr/btnBUTTON_1").press()
        except Exception:
            # Fallback to ENTER if the button ID differs
            try:
                self.session.findById("wnd[1]").sendVKey(0)
            except Exception as e:
                raise SapGuiError(f"Could not confirm deletion: {e}")

    # If OSE, enter 5VE#
    def insert_VEnumber(self, ve_number: str) -> None:
        grid_id = "wnd[0]/usr/cntlC_REP_VEND/shellcont/shell"
        self._wait_control(grid_id, timeout=20)
        grid = self.session.findById(grid_id)

        # Insert a new row
        grid.pressToolbarButton("INSERT")

        # Position and sort on LIFNR (some systems may not expose SORT; tolerate it)
        try:
            grid.setCurrentCell(-1, "LIFNR")
        except Exception:
            pass
        try:
            grid.selectColumn("LIFNR")
        except Exception:
            pass
        try:
            grid.pressToolbarButton("&SORT_ASC")
        except Exception:
            # Sorting button may not exist depending on grid config; safe to ignore
            pass

        # Transform the VE number (first digit -> '5') and write to first visible row
        lifnr_value = self._to_vendor_number(ve_number)
        self._wait_control(grid_id, timeout=10)
        grid.modifyCell(0, "LIFNR", lifnr_value)
        try:
            grid.currentCellColumn = "LIFNR"
        except Exception:
            pass

        # Commit and save
        try:
            grid.pressEnter()
        except Exception:
            # Fallback: main window Enter
            self.session.findById("wnd[0]").sendVKey(0)

        # Save (toolbar button 11)
        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

    #Zendesk rep company -> rep integ key via config
    def resolve_rep_integ_key(self, rep_company: str | None) -> str:
        mapping = CFG.sap.get("rep_integration_map", {})

        map_key = (rep_company or "").strip().lower()
        integ_key = mapping.get(map_key, mapping.get("otherrep", "Other"))

        return integ_key

    # If OSE, rep integration table
    def open_integr_comp_and_set(self, ve_number: str, rep_integ: str, vkorg: str = "4600") -> None:
        # Back twice
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        except Exception:
            # If not possible to go back twice, continue with best effort
            pass

        # Start transaction ZREP_INTEGR_COMP
        try:
            # literal OKCode
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZREP_INTEGR_COMP"
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception:
            # Fallback to start_tx helper
            try:
                self.start_tx("ZREP_INTEGR_COMP")
            except Exception as e:
                raise SapGuiError(f"Unable to start ZREP_INTEGR_COMP: {e}")

        # Press New/Insert
        try:
            self._wait_control("wnd[0]/tbar[1]/btn[5]", timeout=10)
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
        except Exception as e:
            raise SapGuiError(f"Could not press 'New/Insert' on ZREP_INTEGR_COMP: {e}")

        #  Write VKORG and SREP in the table control
        tbl_id = "wnd[0]/usr/tblSAPLZREP_MVIEWSTCTRL_ZREP_INTEGR_COMP"
        self._wait_control(tbl_id, timeout=15)

        # VKORG at column 0, row 0 (ctxt)
        try:
            self.session.findById(f"{tbl_id}/ctxtZREP_INTEGR_COMP-VKORG[0,0]").text = str(vkorg)
        except Exception:
            # Some themes expose txt instead of ctxt
            self.session.findById(f"{tbl_id}/txtZREP_INTEGR_COMP-VKORG[0,0]").text = str(vkorg)

        # SREP (vendor) at column 1, row 0; VE #
        try:
            self.session.findById(f"{tbl_id}/txtZREP_INTEGR_COMP-SREP[1,0]").text = ve_number
        except Exception:
            # Fallback to ctxt if column is configured as a combo/context field
            self.session.findById(f"{tbl_id}/ctxtZREP_INTEGR_COMP-SREP[1,0]").text = ve_number

        # Focus COMP [2,0]
        comp_id_ctxt = f"{tbl_id}/ctxtZREP_INTEGR_COMP-COMP[2,0]"
        comp_id_txt  = f"{tbl_id}/txtZREP_INTEGR_COMP-COMP[2,0]"
        comp_id = comp_id_ctxt
        try:
            self.session.findById(comp_id_ctxt).setFocus()
            self.session.findById(comp_id_ctxt).caretPosition = 0
        except Exception:
            comp_id = comp_id_txt
            self.session.findById(comp_id_txt).setFocus()
            self.session.findById(comp_id_txt).caretPosition = 0

        # F4 (value help)
        self.session.findById("wnd[0]").sendVKey(4)

        # Wait for popup and press "Selection options" (tbar[0]/btn[17])
        self._wait_control("wnd[1]", timeout=10)
        try:
            self.session.findById("wnd[1]/tbar[0]/btn[17]").press()
        except Exception:
            # Some themes: button might not exist; fallback to ENTER to open selection screen
            try:
                self.session.findById("wnd[1]").sendVKey(0)
            except Exception:
                pass

        # Enter VKORG and Company on the selection screen (tab 1)
        # VKORG field [0,24] often ctxt; Company field [1,24] often txt
        sel_vkorg = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[0,24]"
        sel_comp_txt = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]"
        sel_comp_ctxt = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]"

        self._wait_control(sel_vkorg, timeout=10)
        self.session.findById(sel_vkorg).text = str(vkorg)

        if not rep_integ:
            raise SapGuiError("Company value is required for COMP search (e.g., 'CAROTEK_US').")
        try:
            self.session.findById(sel_comp_txt).text = rep_integ
            fld = self.session.findById(sel_comp_txt)
        except Exception:
            self.session.findById(sel_comp_ctxt).text = rep_integ
            fld = self.session.findById(sel_comp_ctxt)

        try:
            fld.setFocus()
            fld.caretPosition = len(rep_integ)
        except Exception:
            pass

        # Execute selection (OK)
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Focus first result (label) then confirm
        try:
            self._wait_control("wnd[1]/usr/lbl[7,3]", timeout=10)
            self.session.findById("wnd[1]/usr/lbl[7,3]").setFocus()
            try:
                self.session.findById("wnd[1]/usr/lbl[7,3]").caretPosition = 2
            except Exception:
                pass
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            # If result grid differs, try simple ENTER to accept current selection
            self.session.findById("wnd[1]").sendVKey(0)

        # Save
        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

    
# ----------------------------
# MAIN
# ----------------------------
def run(ticket_id: int | None = None, program: str = "ZREP_VENDORS") -> None:
    if ticket_id is None:
        raw = input("Enter Zendesk ticket #: ").strip()
        ticket_id = sanitize_ticket_input(raw)

    # Pull E-number and Rep company from Zendesk
    ticket = get_ticket_core_fields(ticket_id)
    if isinstance(ticket, tuple) and len(ticket) == 2:  # backward-compat
        ticket = ticket[0]

    rep_company_name = ticket.get("rep_company")
    e_number = ticket.get("e_number")
    pernr = ticket.get("ve_number")
    rep_type = ticket.get("onboarding_rep_flag")

    # Brief summary
    print("\n[ISE/OSE Offboarding]")
    print("\n⚠️  Ticket Data Retrieved: ⚠️")
    print(f" Rep Company Key:              {rep_company_name}")
    print(f" E Number:                     {e_number}")
    print(f" Rep Type:                     {rep_type}")
    print(f" VE Number (OSE's only):       {pernr}")

    if not _yes_no("\nProceed with table assignment removal using the details above?"):
        print("Cancelled by user.")
        return

    sap = SapGui()

    # 1. Open SA38
    sap.start_tx("SA38")

    # 2. Run ZREP_VENDORS
    sap.run_report(program)

    # 3. Get table key and filter reps
    table_key = sap.resolve_rep_table_key(rep_company_name)
    sap.filter_rep_company_and_open(table_key)

    # 4. Delete VE number
    sap.delete_Enumber(e_number)

    '''
    # 4. Insert E#
    sap.insert_Enumber(e_number)

    # 5. Insert VE# if OSE
    if rep_type == "newose":
        sap.insert_VEnumber(pernr)

    # 6. Rep integration table
    if rep_type == "newose":
        rep_integ_key = sap.resolve_rep_integ_key(rep_company_name)
        sap.open_integr_comp_and_set(pernr, rep_integ_key)
    '''

    print(f"\n✅ PCA table removal complete")
