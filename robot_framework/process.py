from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import smtplib
from email.message import EmailMessage
import json
from  datetime import datetime 
import pyodbc
def insert_email_text(cur, mailtekst, caseid):
    # 1) cases
    cur.execute("""
        INSERT INTO dbo.cases (
           EmailtekstUdlevering 
        )
        VALUES (?)
    """, (mailtekst)
    )

    # 3) caselogs — inkl. UTC timestamp
    utc_now = datetime.now()
    cur.execute("""
        INSERT INTO dbo.caselogs (case_aktid, message, field, action, [user], [timestamp])
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        caseid,
        "Mail sendt med udleveringslink",
        "status",
        "modtaget",
        "System",
        utc_now
    ))

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
    msg.add_alternative(body, subtype='html')
    msg['Bcc'] = UdviklerMail

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send success email: {e}")
        raise e
    #----------------- Here the case details are sent to the database
    sql_server = orchestrator_connection.get_constant("SqlServer").value  
    conn_string = f"DRIVER={{SQL Server}};SERVER={sql_server};DATABASE=AKTINDSIGTERPERSONALEMAPPER;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    conn.autocommit = False
    cur = conn.cursor()
    insert_email_text(cur, body, caseid)
