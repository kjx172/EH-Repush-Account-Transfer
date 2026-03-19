# mk02_open_vendor.py
from config import AppConfig
from get_ticket import sanitize_ticket_input, get_ticket_core_fields
import time

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None


class SapGuiError(Exception):
    pass


class SapGui:
    """Minimal SAP GUI wrapper mirroring patterns used in ise_ose.py."""

    def __init__(self) -> None:
        if win32com is None:
            raise SapGuiError(
                "pywin32 not available. Run on Windows with SAP GUI installed and 'pip install pywin32'."
            )
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        self.application = self.SapGuiAuto.GetScriptingEngine
        if self.application.Connections.Count == 0:
            raise SapGuiError("No active SAP GUI connection found. Log in to a system first.")
        self.connection = self.application.Connections(0)
        if self.connection.Children.Count == 0:
            raise SapGuiError("No active SAP GUI session found. Open a session and try again.")
        self.session = self.connection.Children(0)
        self._wait_control("wnd[0]", timeout=20)
        self._wait_control("wnd[0]/tbar[0]/okcd", timeout=20)

    # ----------------------------
    # HELPERS
    # ----------------------------
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

    def _to_vendor_number(self, ve_number: str) -> str:
        """Replace the first digit of VE number with '5' to form vendor (LIFNR)."""
        if not ve_number:
            raise SapGuiError("VE number is empty.")
        s = str(ve_number)
        return "5" + s[1:] if len(s) > 1 else "5"

    def start_tx(self, code: str) -> None:
        self._wait_control("wnd[0]/tbar[0]/okcd", timeout=20)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{code.upper()}"
        self._exec()

    # ----------------------------
    # FLOW: MK02 vendor 'OPEN'
    # ----------------------------
    def mark_vendor_open(self, ve_number: str) -> None:
        """Implements the exact sequence provided for MK02 using the VE -> vendor transform."""
        # 0) Resize (best effort)
        try:
            self.session.findById("wnd[0]").resizeWorkingPane(155, 43, False)
        except Exception:
            pass

        # 1) MK02
        self.start_tx("mk02")

        # 2) Checkboxes D0110 and D0120
        try:
            self.session.findById("wnd[0]/usr/chkRF02K-D0110").selected = True
        except Exception:
            pass
        try:
            self.session.findById("wnd[0]/usr/chkRF02K-D0120").selected = True
        except Exception:
            pass

        # 3) LIFNR = 5 + (VE w/o first digit), EKORG empty
        lifnr = self._to_vendor_number(ve_number)
        self.session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = lifnr
        try:
            self.session.findById("wnd[0]/usr/ctxtRF02K-EKORG").text = ""
        except Exception:
            pass

        # 4) Focus D0120 and ENTER
        try:
            self.session.findById("wnd[0]/usr/chkRF02K-D0120").setFocus()
        except Exception:
            pass
        self._exec()

        # 5) Set NAME1 and SORT1 to OPEN
        self.session.findById(
            "wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1"
        ).text = "OPEN"
        sort1 = self.session.findById(
            "wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1"
        )
        sort1.text = "OPEN"
        try:
            sort1.setFocus()
        except Exception:
            pass
        try:
            sort1.caretPosition = 4
        except Exception:
            pass

        # 6) Save
        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()


# ----------------------------
# MAIN ENTRY
# ----------------------------

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

def run(ticket_id: int | None = None) -> None:
    """Read VE number from Zendesk ticket and run the MK02 'OPEN' flow."""
    if ticket_id is None:
        raw = input("Enter Zendesk ticket #: ").strip()
        ticket_id = sanitize_ticket_input(raw)

    ticket = get_ticket_core_fields(ticket_id)
    if isinstance(ticket, tuple) and len(ticket) == 2:  # backward-compat
        ticket = ticket[0]

    ve_number = ticket.get("ve_number")

    print("\n[MK02 Vendor OPEN]")
    print("\n⚠️ Ticket Data Retrieved: ⚠️")
    print(f" VE Number: {ve_number}")
    if not ve_number:
        raise SapGuiError("Ticket is missing 've_number'.")

    if not _yes_no("Proceed with MK02 update using the VE number above?"):
        print("Cancelled by user.")
        return

    sap = SapGui()
    sap.mark_vendor_open(str(ve_number))
    print("\n✅ MK02 update complete")


if __name__ == "__main__":
    try:
        run()
    except SystemExit:
        raise
    except Exception as e:
        print(f"\n[ERROR] MK02 Vendor OPEN failed: {e}\n")