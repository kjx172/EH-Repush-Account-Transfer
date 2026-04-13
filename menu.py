# menu.py
import sys
from typing import Callable
import os
from repush import run_flow as cmd_repush
from acct_transfer import run_flow as cmd_account_transfer
from asp_offboarding import run_flow as cmd_asp_offboarding
from ise_ose import run as cmd_ise_ose_offboarding
from mk02_open_vendor import run as cmd_vendor_record


# Menu options
MENU = {
    "1": ("ASP Table Assignment", cmd_asp_offboarding),
    "2": ("ISE/OSE Table assignment", cmd_ise_ose_offboarding),
    "3": ("OSE Vendor Record", cmd_vendor_record),
    "4": ("Mass Account Transfer", cmd_account_transfer),
    "5": ("Repush Orders",   lambda: cmd_repush("order")),
    "6": ("Repush Invoices", lambda: cmd_repush("invoice")),
    "q": ("Quit",            lambda: sys.exit(0)),
}


def main() -> None:
    while True:
        print('\n==== Offboarding SAP Utilities ==== v1.0.2' )
        for key, (label, _) in MENU.items():
            print(f" {key}) {label}")
        choice = input('Select an option: ').strip().lower()
        action = MENU.get(choice)
        if not action:
            print('❌ Invalid selection. Try again.')
            continue
        label, func = action
        try:
            func()
        except SystemExit:
            raise
        except Exception as e:
            print(f"\n[ERROR] {label} failed: {e}\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ Unhandled error:")
        print(e)
        input("\nPress ENTER to exit...")
