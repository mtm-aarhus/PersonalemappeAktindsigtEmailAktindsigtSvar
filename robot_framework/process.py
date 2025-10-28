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

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    specific_content = json.loads(queue_element.data)
    caseid = specific_content.get("caseid")
    IndsenderMail = specific_content.get("to")
    SagsbehandlerMail = specific_content.get("from")
    UdviklerMail = orchestrator_connection.get_constant('balas')
    body = specific_content.get('body')
    subject = specific_content.get('subject')


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
