# zendesk_sap_cli/config.py
from __future__ import annotations
from pathlib import Path
from typing import Any, Dict
import yaml
import sys
import os


def _resource_path(relative_path: str) -> Path:
    """
    Get the absolute path to a resource.
    Handles both PyInstaller (frozen) and normal execution.

    Example: _resource_path("config.yaml")
    """
    if hasattr(sys, "_MEIPASS"):
        # Running inside a PyInstaller bundle
        base_path = Path(sys._MEIPASS)
    else:
        # Running normally (script mode)
        base_path = Path(__file__).parent

    return base_path / relative_path


class AppConfig(dict):

    @classmethod
    def load(cls, path: str | Path | None = None) -> "AppConfig":
        # If user supplied a path → use it
        if path:
            cfg_path = Path(path)

        else:
            # Otherwise load config.yaml next to EXE or script
            cfg_path = _resource_path("config.yaml")

        with open(cfg_path, "r", encoding="utf-8") as f:
            data: Dict[str, Any] = yaml.safe_load(f) or {}

        return cls(data)

    # -----------------------------
    # ZENDESK FIELD ACCESSORS
    # -----------------------------
    @property
    def employee_name_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("employee_name_field_id", "")).strip()

    @property
    def start_date_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("start_date_field_id", "")).strip()

    @property
    def employee_region_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("employee_region_field_id", "")).strip()

    @property
    def i_number_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("i_number_field_id", "")).strip()

    @property
    def e_number_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("e_number_field_id", "")).strip()

    @property
    def email_internal_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("email_internal_field_id", "")).strip()

    @property
    def email_rep_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("email_rep_field_id", "")).strip()

    @property
    def company_address_rep_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("company_address_rep_field_id", "")).strip()

    @property
    def onboarding_internal_flag_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("onboarding_internal_flag_field_id", "")).strip()

    @property
    def onboarding_rep_flag_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("onboarding_rep_flag_field_id", "")).strip()

    @property
    def rep_company_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("rep_company_field_id", "")).strip()
    
    @property
    def ve_number_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("ve_number_field_id", "")).strip()
    
    @property
    def salesforce_alias_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("salesforce_alias_field_id", "")).strip()
    
    @property
    def phone_field_id(self) -> str:
        return str(self.get("zendesk", {}).get("phone_field_id", "")).strip()

    # -----------------------------
    # SAP DEFAULTS
    # -----------------------------
    @property
    def sap(self) -> dict:
        return self.get("sap_defaults", {})
