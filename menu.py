# menu.py
import sys
from typing import Callable
import os
from repush import run_flow as cmd_repush
from sa38 import run_flow as cmd_account_transfer
from asp_offboarding import run_flow as cmd_asp_offboarding


# Menu options
MENU = {
    "1": ("ASP Offboarding", cmd_asp_offboarding),
    "2": ("Mass Account Transfer", cmd_account_transfer),
    "3": ("Repush Orders",   lambda: cmd_repush("order")),
    "4": ("Repush Invoices", lambda: cmd_repush("invoice")),
    "q": ("Quit",            lambda: sys.exit(0)),
}


def main() -> None:
    while True:
        print('\n==== Additional SAP Utilities ====' )
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
