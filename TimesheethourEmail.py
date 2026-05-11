import os
import requests
import base64
from dotenv import load_dotenv

load_dotenv()

GRAPH_TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
GRAPH_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
GRAPH_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")

GRAPH_SENDER = os.getenv(
    "GRAPH_SENDER",
    "hr@marlugroupwa.com.au"
)

GRAPH_TO = os.getenv(
    "GRAPH_TO",
    "planning@marlugroupwa.com.au"
)


# ======================================================
# GET GRAPH TOKEN
# ======================================================

def get_graph_access_token():

    token_url = (
        f"https://login.microsoftonline.com/"
        f"{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    )

    data = {
        "client_id": GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    response = requests.post(
        token_url,
        data=data,
        timeout=30
    )

    response.raise_for_status()

    return response.json()["access_token"]


# ======================================================
# SEND EMAIL WITH MEMORY ATTACHMENT
# ======================================================

def send_email_with_attachment(
    subject,
    body,
    attachment_bytes,
    filename
):

    token = get_graph_access_token()

    # memory bytes -> base64
    attachment_content = base64.b64encode(
        attachment_bytes
    ).decode("utf-8")

    url = (
        f"https://graph.microsoft.com/v1.0/users/"
        f"{GRAPH_SENDER}/sendMail"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body,
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": GRAPH_TO
                    }
                }
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentBytes": attachment_content,
                }
            ],
        },
        "saveToSentItems": True,
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=60
    )

    if not response.ok:

        print("❌ EMAIL SEND ERROR")
        print(response.text)

    response.raise_for_status()

    print(f"✅ Email sent to {GRAPH_TO}")