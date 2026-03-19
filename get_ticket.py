# get_ticket.py
import os
import requests
from dotenv import load_dotenv
from config import AppConfig


# ----------------------------
# CFG & ENV data
# ----------------------------
load_dotenv()
CFG = AppConfig.load()
EMPLOYEE_NAME_FIELD_ID = CFG.employee_name_field_id
START_DATE_FIELD_ID = CFG.start_date_field_id
EMPLOYEE_REGION_FIELD_ID = CFG.employee_region_field_id
I_NUMBER_FIELD_ID = CFG.i_number_field_id
E_NUMBER_FIELD_ID = CFG.e_number_field_id
EMAIL_INTERNAL_FIELD_ID = CFG.email_internal_field_id
EMAIL_REP_FIELD_ID = CFG.email_rep_field_id
COMPANY_ADDR_REP_FIELD_ID = CFG.company_address_rep_field_id
VE_NUMBER_FIELD_ID = CFG.ve_number_field_id
SALESFORCE_ALIAS_FIELD_ID = CFG.salesforce_alias_field_id
ONBOARDING_INTERNAL_FLAG = CFG.onboarding_internal_flag_field_id
ONBOARDING_REP_FLAG = CFG.onboarding_rep_flag_field_id
REP_COMPANY_FIELD_ID = CFG.rep_company_field_id
PHONE_FIELD_ID = CFG.phone_field_id

# ----------------------------
# ZENDESK API INFO
# ----------------------------

class ZendeskAuthError(Exception):
    pass

class ZendeskClient:
    def __init__(self):
        self.subdomain = os.getenv('ZENDESK_SUBDOMAIN')
        self.email = os.getenv('ZENDESK_EMAIL')
        self.token = os.getenv('ZENDESK_API_TOKEN')
        missing = [k for k,v in {
            'ZENDESK_SUBDOMAIN': self.subdomain,
            'ZENDESK_EMAIL': self.email,
            'ZENDESK_API_TOKEN': self.token,
        }.items() if not v]
        if missing:
            raise ZendeskAuthError(f"Missing environment variables: {', '.join(missing)}")
        self.base = f"https://{self.subdomain}.zendesk.com/api/v2"

    def fetch_ticket(self, ticket_id: int) -> dict:
        url = f"{self.base}/tickets/{ticket_id}.json"
        auth = (f"{self.email}/token", self.token)
        resp = requests.get(url, auth=auth, timeout=30)
        resp.raise_for_status()
        return resp.json()["ticket"]

    @staticmethod
    def flatten_custom_fields(ticket: dict) -> dict:
        flat = {}
        for cf in ticket.get('custom_fields', []):
            key = cf.get('key') or str(cf.get('id'))
            flat[key] = cf.get('value')
        return flat

# ----------------------------
# UTILS
# ----------------------------
def sanitize_ticket_input(raw: str) -> int:
    """Accepts '#num' or 'num', strips leading '#' and validates numeric."""
    s = raw.strip()
    if s.startswith('#'):
        s = s[1:]
    if not s.isdigit():
        raise ValueError('Ticket ID must be numeric (optionally prefixed with #).')
    return int(s)

# Cache for ticket field option lookups
_FIELD_DEF_CACHE = {}

def _resolve_dropdown_display(client: "ZendeskClient", field_id: str, tag: str) -> str:
    """
    For a dropdown field, convert the stored tag (value) to its display 'name'.
    If resolution fails, return the original tag.
    """
    if not tag:
        return tag
    
    # load + cache field definition
    fld = _FIELD_DEF_CACHE.get(field_id)
    if fld is None:
        url = f"{client.base}/ticket_fields/{field_id}.json"
        auth = (f"{client.email}/token", client.token)
        resp = requests.get(url, auth=auth, timeout=30)
        resp.raise_for_status()
        fld = resp.json().get("ticket_field", {})
        _FIELD_DEF_CACHE[field_id] = fld
    try:
        for opt in fld.get("custom_field_options", []):
            if opt.get("value") == tag:
                # prefer 'name' (display), fall back to 'raw_name' or the tag
                return opt.get("name") or opt.get("raw_name") or tag
    except Exception:
        pass
    return tag

def get_ticket_core_fields(ticket_id: int) -> dict:
    # Initializ client and get provided ticket
    client = ZendeskClient()
    t = client.fetch_ticket(ticket_id)

    result = {
        "employee_name": None,
        "start_date": None,
        "employee_region": None,
        "i_number": None,
        "e_number": None,
        "email_internal": None,
        "email_rep": None,
        "company_address_rep": None,
        "is_internal": False,
        "is_rep": False,
        "rep_company": None,
        "onboarding_rep_flag": None,
        "ve_number": None,
        "salesforce_alias": None,
        "phone_number": None
    }

    # Use ticket ID to fill results with relevant data
    for cf in t.get("custom_fields", []):
        cid = str(cf.get("id"))
        val = cf.get("value")

        if cid == EMPLOYEE_NAME_FIELD_ID:
            result["employee_name"] = val
        elif cid == START_DATE_FIELD_ID:
            result["start_date"] = val
        elif cid == EMPLOYEE_REGION_FIELD_ID:
            result["employee_region"] = val
        elif cid == I_NUMBER_FIELD_ID:
            result["i_number"] = val
        elif cid == E_NUMBER_FIELD_ID:
            result["e_number"] = val
        elif cid == EMAIL_INTERNAL_FIELD_ID:
            result["email_internal"] = val
        elif cid == EMAIL_REP_FIELD_ID:
            result["email_rep"] = val
        elif cid == COMPANY_ADDR_REP_FIELD_ID:
            result["company_address_rep"] = val
        elif cid == ONBOARDING_INTERNAL_FLAG and val:
            result["is_internal"] = True
        elif cid == ONBOARDING_REP_FLAG:
            result["onboarding_rep_flag"] = val
            if val:
                result["is_rep"] = True
        elif cid == REP_COMPANY_FIELD_ID:
            result["rep_company"] = val
        elif cid == VE_NUMBER_FIELD_ID:
            result["ve_number"] = val
        elif cid == SALESFORCE_ALIAS_FIELD_ID:
            result["salesforce_alias"] = _resolve_dropdown_display(client, SALESFORCE_ALIAS_FIELD_ID, val)
        elif cid == PHONE_FIELD_ID:
            result["phone_number"] = val

    return result