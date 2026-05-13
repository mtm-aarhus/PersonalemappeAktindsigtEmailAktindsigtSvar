import requests
from requests_ntlm import HttpNtlmAuth
import xml.etree.ElementTree as ET
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import json

def create_ntlm_session(username: str, password: str) -> requests.Session:
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    return session

def get_site_digest(site_url: str, session: requests.Session) -> str:
    """Henter et FormDigestValue for det angivne web/scope."""
    endpoint = f"{site_url}/_api/contextinfo"
    r = session.post(endpoint, headers={"Accept": "application/json; odata=verbose"})
    r.raise_for_status()
    digest = r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
    print(f"Got digest for {site_url}: {digest[:20]}...")
    return digest

def search_sharepoint_user(root_api_url: str, session: requests.Session, digest: str, email: str):
    """Søger efter en bruger i PeoplePicker (kræver root-level digest)."""
    endpoint = f"{root_api_url}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser"

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
    }

    payload = {
        "queryParams": {
            "QueryString": email,
            "MaximumEntitySuggestions": 50,
            "AllowEmailAddresses": False,
            "AllowOnlyEmailAddresses": False,
            "PrincipalType": 1,
            "PrincipalSource": 15,
            "SharePointGroupID": 0
        }
    }

    r = session.post(endpoint, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    results = json.loads(r.json()["d"]["ClientPeoplePickerSearchUser"])
    for entity in results:
        entity_email = entity.get("EntityData", {}).get("Email")
        if entity_email and entity_email.lower() == email.lower():
            return entity
    return None

def get_list_and_id(api_url, aktid, session):
    """Henter liste og itemid til opdatering af bruger i go."""
    endpoint = f"{api_url}/cases/AKT50/{aktid}/_goapi/Administration/ModernConfiguration"
    
    payload = {
        "providerTypes": ["ModernCase", "MoveDocument", "Insight", "SearchSystem", "UserSettings"]
    }
    
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    r = session.post(endpoint, headers=headers, json=payload)
    r.raise_for_status()
    data = r.json()
    caselist = data.get("ModernCase").get("ItemServerUrl").split('/')[-2]
    itemid = data.get("ModernCase").get("ListItemID")
    return(caselist, itemid)

def update_case_field(api_url: str, session: requests.Session, digest: str, form_values: list, listnumber, item_id):
    """Opdaterer felt(er) i sagslisten."""
    endpoint = (
        f"{api_url}/aktindsigt/_api/web/GetList(@a1)/items(@a2)/ValidateUpdateListItem()"
        f"?@a1='%2Faktindsigt%2FLists%2F{listnumber}'&@a2='{item_id}'"
    )

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "X-Sp-Requestresources": f"listUrl=%2Faktindsigt%2FLists%2F{listnumber}"
    }

    payload = {
        "formValues": form_values,
        "bNewDocumentUpdate": False,
        "checkInComment": None
    }

    r = session.post(endpoint, headers=headers, data=json.dumps(payload))
    r.raise_for_status()
    return r.json()

def update_case_owner(api_url: str, username: str, password: str, case_id: str, email_anmoder: str, ):
    """Opdaterer sagens CaseOwner-felt korrekt med to forskellige digests."""
    session = create_ntlm_session(username, password)

    # 1️⃣ Root-digest (bruges til PeoplePicker)
    root_digest = get_site_digest(api_url, session)

    # 2️⃣ Aktindsigt-digest (bruges til feltopdatering)
    akt_digest = get_site_digest(f"{api_url}/aktindsigt", session)
    
    listnumber, item_id = get_list_and_id(api_url, case_id, session)
    # verify_case_item(api_url, session, listnumber, item_id)

    # Find bruger
    caseowner_entity = search_sharepoint_user(api_url, session, root_digest, email_anmoder)
    if not caseowner_entity:
        return False

    form_values = [
    {
        "FieldName": "SupplerendeSagsbehandlere",
        "FieldValue": json.dumps([caseowner_entity]),
        "HasException": False,
        "ErrorMessage": None
    }
]

    # Opdater felt
    result = update_case_field(api_url, session, akt_digest, form_values, listnumber, item_id)
    return result

def verify_case_item(api_url, session, listnumber, item_id):
    endpoint = (
        f"{api_url}/aktindsigt/_api/web/GetList(@a1)/items(@a2)"
        f"?@a1='%2Faktindsigt%2FLists%2F{listnumber}'&@a2='{item_id}'"
    )
    headers = {"Accept": "application/json;odata=verbose"}
    r = session.get(endpoint, headers=headers)
    r.raise_for_status()


def close_case(case_id, session, go_api_url):
    url = f"{go_api_url}/_goapi/Cases/CloseCase"

    payload = json.dumps({
    "CaseId": case_id
    })
    headers = {
    'Content-Type': 'application/json'
    }

    response = session.post( url, headers=headers, data=payload)

    return response.text

def get_case_metadata(gourl, sagsnummer, session):
    url = f"{gourl}/_goapi/Cases/Metadata/{sagsnummer}"

    session.headers.update({"Content-Type": "application/json"})

    response = session.get( url)

    return response.text

import json
import re

def get_case_documents(session, GOAPI_URL, SagsURL, SagsID):

    Akt = SagsURL.split("/")[1]
    encoded_sags_id = SagsID.replace("-", "%2D")
    ListURL = f"%27%2Fcases%2F{Akt}%2F{encoded_sags_id}%2FDokumenter%27"

    ViewId = None
    ikke_journaliseret_id = None
    journaliseret_id = None
    view_ids_to_use = []
    all_rows = []

    response = session.get(f"{GOAPI_URL}/{SagsURL}/_goapi/Administration/GetLeftMenuCounter")
    response.raise_for_status()
    
    ViewsIDArray = json.loads(response.text)

    for item in ViewsIDArray:
        if item["ViewName"] == "UdenMapper.aspx":
            ViewId = item["ViewId"]
            break

        elif item["ViewName"] == "Ikkejournaliseret.aspx":
            ikke_journaliseret_id = item["ViewId"]
            if ikke_journaliseret_id is None:
                LinkURL = item["LinkUrl"]
                response = session.get(f'{GOAPI_URL}{LinkURL}')
                response.raise_for_status()

                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', response.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                context_info = json.loads(match.group(1))
                ikke_journaliseret_id = context_info.get("viewId", "").strip("{}")

        elif item["ViewName"] == "Journaliseret.aspx":
            journaliseret_id = item["ViewId"]
            if journaliseret_id is None:
                LinkURL = item["LinkUrl"]
                response = session.get(f'{GOAPI_URL}{LinkURL}')
                response.raise_for_status()

                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', response.text, re.DOTALL)
                if not match:
                    raise ValueError("Kunne ikke finde _spPageContextInfo i HTML")

                context_info = json.loads(match.group(1))
                journaliseret_id = context_info.get("viewId", "").strip("{}")

    if ViewId is None:
        view_ids_to_use = [vid for vid in [ikke_journaliseret_id, journaliseret_id] if vid]

    views = [ViewId] if ViewId else view_ids_to_use

    if not views:
        raise ValueError("Ingen gyldige ViewId fundet")

    for current_view_id in views:
        firstrun = True
        MorePages = True
        NextHref = None

        while MorePages:
            url = f"{GOAPI_URL}/{SagsURL}/_api/web/GetList(@listUrl)/RenderListDataAsStream"

            if firstrun:
                full_url = f"{url}?@listUrl={ListURL}&View={current_view_id}"
            else:
                full_url = f"{url}?@listUrl={ListURL}{NextHref.replace('?', '&')}"

            response = session.post(full_url, timeout=500)
            response.raise_for_status()

            dokumentliste_json = response.json()
            dokumentliste_rows = dokumentliste_json.get("Row", [])
            all_rows.extend(dokumentliste_rows)

            NextHref = dokumentliste_json.get("NextHref")
            MorePages = bool(NextHref)
            firstrun = False

    return all_rows