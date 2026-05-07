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

def journaliser_sag(go_api_url: str, case_id: str, session: requests.Session, orchestrator_connection: OrchestratorConnection):
    """Henter alle dokumenter på sagen og journaliserer dem inden lukning."""
    # Hent sagens relative URL via metadata
    response = session.get(f"{go_api_url}/_goapi/Cases/Metadata/{case_id}/False")
    response.raise_for_status()
    metadata_str = response.json().get("Metadata")
    xdoc = ET.fromstring(metadata_str)
    relative_case_url = xdoc.attrib.get("ows_CaseUrl")

    # Hent dokumenter på sagen
    list_url = f"'/{relative_case_url}/Dokumenter'"
    url = f"{go_api_url}/{relative_case_url}/_api/web/GetList(@listUrl)/RenderListDataAsStream?@listUrl={list_url}"
    payload = json.dumps({"parameters": {"__metadata": {"type": "SP.RenderListDataParameters"}, "ViewXml": "<View><RowLimit>500</RowLimit></View>"}})
    response = session.post(url, headers={"content-type": "application/json;odata=verbose"}, data=payload)
    response.raise_for_status()

    doc_ids = [str(row.get("DocID")) for row in response.json().get("Row", []) if row.get("DocID")]

    if doc_ids:
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
    except Exception as e:
        orchestrator_connection.log_error(f'Process failed: {e}')
        raise e

