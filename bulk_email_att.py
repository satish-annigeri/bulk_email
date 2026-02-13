import os
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from mako.template import Template


load_dotenv()
login_id = os.environ.get("LOGIN_ID", "")
sender_name = os.environ.get("SENDER_NAME", "")
pwd = os.environ.get("APP_PASSWORD", "")


def render_message(tpl_txt: str, data: dict) -> bytes | str:
    tpl = Template(tpl_txt)
    return tpl.render(**data)


def build_message(
    sender_name: str,
    login_id: str,
    recipient_name: str,
    recipient_email: str,
    subject: str,
    body: str,
    pdf_path: str,
    pdf_bytes: bytes,
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, login_id))
    msg["To"] = formataddr((recipient_name, recipient_email))
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(
        pdf_bytes, maintype="application", subtype="pdf", filename=pdf_path
    )
    return msg


def send_message(server: smtplib.SMTP, msg: EmailMessage):
    try:
        server.send_message(msg)
        print("Message sent")
    except Exception as e:
        print(f"Error sending message: {e}")


if __name__ == "__main__":
    pdf_path = "80G_website_approval.pdf"
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    tpl_txt = """Dear ${name},
Please find attached the PDF document you requested. Best regards, ${sender_name}"""
    data = {
        "name": "Satish Annigeri",
        "sender_name": sender_name,
    }

    # with smtplib.SMTP("smtp.gmail.com", 587) as server:
    #     server.starttls()
    #     server.login(login_id, pwd)

    #     body = render_message(tpl_txt, data)
    #     msg = build_message(
    #         sender_name,
    #         login_id,
    #         "Satish Annigeri",
    #         "satish.annigeri@outlook.com",
    #         "Test Email with Attachment",
    #         "This is a test email with an attachment.",
    #         pdf_path,
    #         pdf_bytes,
    #     )
    #     send_message(server, msg)
