from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import smtplib
from email.message import EmailMessage
import json
from  datetime import datetime 
import html
import re
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
from GoBrugerstyring import *
import json
import xml.etree.ElementTree as ET
import re

def text_to_html(body: str) -> str:
    """Konverterer plain text med linjeskift til HTML med <br> og klikbare links."""
    if not body:
        return ""

    # Escapér HTML-tegn (<, >, & osv.)
    safe = html.escape(body)

    # Gør links klikbare
    safe = re.sub(
        r"(https?://[^\s]+)",
        r'<a href="\1" target="_blank">\1</a>',
        safe
    )

    # Erstat linjeskift med <br>
    html_body = safe.replace("\n", "<br>\n")

    return html_body
def finaliser_dokumenter(go_api_url: str, doc_ids: list, session: requests.Session, orchestrator_connection: OrchestratorConnection):
    """Gør dokumenter endelige inden sagen lukkes."""
    url = f"{go_api_url}/_goapi/Documents/Finalize/ByDocumentId"
    response = session.post(url, data=json.dumps({"DocumentIds": doc_ids}), headers={"Content-Type": "application/json"})
    orchestrator_connection.log_info(f"Finaliser response status: {response.status_code}")
    orchestrator_connection.log_info(f"Finaliser response body: {response.text}")
    response.raise_for_status()
    orchestrator_connection.log_info(f"Endeliggjorde {len(doc_ids)} dokumenter.")
    
def journaliser_sag(go_api_url: str, case_id: str, session: requests.Session, orchestrator_connection: OrchestratorConnection):
    response = session.get(f"{go_api_url}/_goapi/Cases/Metadata/{case_id}/False")
    response.raise_for_status()
    metadata_str = response.json().get("Metadata")
    xdoc = ET.fromstring(metadata_str)
    relative_case_url = xdoc.attrib.get("ows_CaseUrl")

    # Find ViewId - med fallback til HTML-parsing hvis ViewId er None
    response = session.get(f"{go_api_url}/{relative_case_url}/_goapi/Administration/GetLeftMenuCounter")
    response.raise_for_status()
    
    ikke_journaliseret_id = None
    journaliseret_id = None

    for item in response.json():
        if item.get("ViewName") == "Ikkejournaliseret.aspx":
            ikke_journaliseret_id = item.get("ViewId")
            if ikke_journaliseret_id is None:
                link_url = item.get("LinkUrl")
                html_response = session.get(f"{go_api_url}{link_url}")
                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', html_response.text, re.DOTALL)
                if match:
                    context_info = json.loads(match.group(1))
                    ikke_journaliseret_id = context_info.get("viewId", "").strip("{}")
        elif item.get("ViewName") == "Journaliseret.aspx":
            journaliseret_id = item.get("ViewId")
            if journaliseret_id is None:
                link_url = item.get("LinkUrl")
                html_response = session.get(f"{go_api_url}{link_url}")
                match = re.search(r'_spPageContextInfo\s*=\s*({.*?});', html_response.text, re.DOTALL)
                if match:
                    context_info = json.loads(match.group(1))
                    journaliseret_id = context_info.get("viewId", "").strip("{}")

    view_ids = [vid for vid in [ikke_journaliseret_id, journaliseret_id] if vid]
    if not view_ids:
        orchestrator_connection.log_info("Ingen ikke-journaliserede dokumenter fundet.")
        return

    # Hent dokumenter med paginering
    Akt = relative_case_url.split("/")[1]
    encoded_sags_id = relative_case_url.rsplit("/")[-1].replace("-", "%2D")
    list_url = f"%27%2Fcases%2F{Akt}%2F{encoded_sags_id}%2FDokumenter%27"

    doc_ids = []
    for view_id in view_ids:
        firstrun = True
        more_pages = True
        next_href = None
        base_url = f"{go_api_url}/{relative_case_url}/_api/web/GetList(@listUrl)/RenderListDataAsStream"

        while more_pages:
            if firstrun:
                url = f"{base_url}?@listUrl={list_url}&View={view_id}"
            else:
                url = f"{base_url}?@listUrl={list_url}{next_href.replace('?', '&')}"

            response = session.post(url, timeout=500)
            response.raise_for_status()
            data = response.json()

            doc_ids.extend(str(row.get("DocID")) for row in data.get("Row", []) if row.get("DocID"))

            next_href = data.get("NextHref")
            more_pages = bool(next_href)
            firstrun = False

    orchestrator_connection.log_info(f"Fandt {len(doc_ids)} ikke-journaliserede dokumenter.")

    if doc_ids:
        finaliser_dokumenter(go_api_url, doc_ids, session, orchestrator_connection)
        
        url = f"{go_api_url}/_goapi/Documents/MarkMultipleAsCaseRecord/ByDocumentId"
        response = session.post(url, data=json.dumps({"DocumentIds": doc_ids}), headers={"Content-Type": "application/json"})
        response.raise_for_status()
        orchestrator_connection.log_info(f"Journaliserede {len(doc_ids)} dokumenter.")

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    specific_content = json.loads(queue_element.data)
    caseid = specific_content.get("caseid")
    udleveringsmappeid = specific_content.get("udleveringsmappeid").split('/')[-1]
    IndsenderMail = specific_content.get("to")
    SagsbehandlerMail = specific_content.get("from")
    UdviklerMail = orchestrator_connection.get_constant('balas')
    body = specific_content.get('body')
    subject = specific_content.get('subject')
    go_api_url = orchestrator_connection.get_constant("GOApiURL").value
    go_api_login = orchestrator_connection.get_credential("GOAktApiUser")
    go_username = go_api_login.username
    go_password = go_api_login.password


    # SMTP Configuration (from your provided details)
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25

    msg= EmailMessage()
    msg['To'] = IndsenderMail
    msg['From'] = SagsbehandlerMail
    msg['Subject'] = subject
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(text_to_html(body), subtype='html')
    msg['Bcc'] = UdviklerMail

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send success email: {e}")
        raise e
    
    try:
        session = create_ntlm_session(go_username, go_password)
        update_case_owner(go_api_url, go_username, go_password, udleveringsmappeid, IndsenderMail)
        journaliser_sag(go_api_url, udleveringsmappeid, session, orchestrator_connection)
        close_case(go_api_url=go_api_url, case_id=udleveringsmappeid, session=session)
        result = close_case(go_api_url=go_api_url, case_id=udleveringsmappeid, session=session)
        orchestrator_connection.log_info(f"close_case response: {result}")
    except Exception as e:
        orchestrator_connection.log_error(f'Process failed: {e}')
        raise e

