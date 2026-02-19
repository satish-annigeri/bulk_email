import time
import os
from os.path import splitext, isfile
import tomllib
from email.message import EmailMessage
from email.utils import formataddr
import smtplib

from dotenv import load_dotenv
import pandas as pd
from mako.template import Template
from markdown import markdown


def read_config(fname_toml: str) -> dict:
    if not isfile(fname_toml):
        raise FileNotFoundError(f"File not found: {fname_toml}")
    with open(fname_toml, "rb") as f:
        config = tomllib.load(f)
    if not isfile(config["recipients"]):
        raise FileNotFoundError(f"Recipients file not found: {config['recipients']}")
    if not isfile(config["template"]):
        raise FileNotFoundError(f"Template file not found: {config['template']}")
    if not isfile(config["attachment"]):
        print(f"Warning: PDF attachment file not found: {config['attachment']}")
    return config


def read_data_file(fname: str) -> pd.DataFrame:
    if not isfile(fname):
        raise FileNotFoundError(f"File not found: {fname}")
    ext = splitext(fname)[1].lower()
    if ext in [".xlsx", ".xls"]:
        print(f"Reading Excel file {fname}")
        df = pd.read_excel(fname)
    elif ext == ".csv":
        print(f"Reading CSV file {fname}")
        df = pd.read_csv(fname)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    return df


def clean_data(
    df: pd.DataFrame,
    cols: list[str] | None = None,
    cols_na: str | list[str] | None = None,
    cols_dup: str | list[str] | None = None,
    cols_sort: str | list[str] | None = None,
) -> pd.DataFrame:
    if cols:  # Rename column names
        df.columns = cols
    if cols_na:  # Drop rows with one or more null values
        df = df.dropna(subset=cols_na)
    if cols_dup:  # Drop rows with duplicate values in the specified columns, keep the first occurence
        df = df.drop_duplicates(subset=cols_dup)
    if cols_sort:  # Sort the DataFrame by the specified columns
        df = df.sort_values(by=cols_sort)
    return df


def mangle_name(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    def clean_name(name: str) -> str:
        name = name.strip()
        hon = name.split(".", 1)
        if len(hon) > 1 and hon[0] in ["Prof", "Dr", "Mr", "Ms", "Er"]:
            name = f"{hon[0]}. {hon[1]}"
        name = " ".join(
            w.capitalize() if len(w) != 2 else w.upper() for w in name.split()
        )
        return name

    df[col_name] = df[col_name].apply(clean_name)
    return df


def read_template(tpl_fname: str) -> tuple[Template, str]:
    if not isfile(tpl_fname):
        raise FileNotFoundError(f"File not found: {tpl_fname}")

    tpl_type = splitext(tpl_fname)[1].lower()

    return Template(filename=tpl_fname), tpl_type


def tpl_render(tpl: Template, tpl_type: str, **kwargs) -> str:
    if tpl_type not in [".html", ".htm", ".md"]:
        raise ValueError(f"Unsupported template file type: {tpl_type}")

    txt = tpl.render(**kwargs)
    if isinstance(txt, bytes):
        txt = txt.decode("utf-8")
    else:
        txt = str(txt)

    if tpl_type == ".md":
        return markdown(txt)
    else:
        return txt


def read_recipients_data(
    recipients_fname: str, cols_dup=["email", "name"], cols_sort=["name"]
) -> pd.DataFrame:
    print(recipients_fname)
    df = read_data_file(recipients_fname)
    df = clean_data(df, cols_dup=cols_dup, cols_sort=cols_sort)
    df = mangle_name(df, "name")
    # print(df)
    return df


def build_message(
    sender_name: str,
    sender_email: str,
    recipient_name: str,
    recipient_email: str,
    subject: str,
    body: str,
    pdf_fname: str = "",
    pdf_bytes: bytes = b"",  # PDF file as bytes to be attached to the message
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender_email))
    msg["To"] = formataddr((recipient_name, recipient_email))
    msg["Subject"] = subject
    msg.set_content(body, subtype="html")
    if pdf_fname and pdf_bytes:
        msg.add_attachment(
            pdf_bytes, maintype="application", subtype="pdf", filename=pdf_fname
        )
    return msg


def send_bulk_emails(
    tpl_fname: str,
    df: pd.DataFrame,
    start: int = 1,
    count: int = -1,
    login_id: str = "",
    pwd: str = "",
    sender_name: str = "",
    subject: str = "",
    pdf_fname: str = "",
    delay: float = 0.5,
    dry_run: bool = True,
) -> None:
    tpl, tpl_type = read_template(tpl_fname)

    i = 0
    if i < 0:
        return
    elif count < 0:
        count = len(df)
    last = min(start + count - 1, len(df))
    if dry_run:
        for _, row in df.iterrows():
            i += 1
            if i >= start and i <= last:
                html = tpl_render(tpl, tpl_type, name=row["name"], mode=True)
                start_portion, _, _ = html.partition("</p>")
                print(f"{row['name']} <{row['email']}> {start_portion}</p>...")
        return
    else:
        if isfile(pdf_fname):
            with open(pdf_fname, "rb") as f:
                pdf_bytes = f.read()
            print(f"PDF attachment {pdf_fname} loaded successfully.")
        else:
            print(f"Attachment {pdf_fname} not found. Continuing without attachment.")
            pdf_bytes = b""
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(login_id, pwd)
            for _, row in df.iterrows():
                i += 1
                if i >= start and i <= last:
                    html = tpl_render(tpl, tpl_type, name=row["name"], mode=row["mode"])
                    msg = build_message(
                        sender_name=sender_name,
                        sender_email=login_id,
                        recipient_name=row["name"],
                        recipient_email=row["email"],
                        subject=subject,
                        body=html,
                        pdf_fname=pdf_fname,
                        pdf_bytes=pdf_bytes,
                    )
                    try:
                        server.send_message(msg)
                        print(f"{row['name']} <{row['email']}> {row['mode']}")
                        time.sleep(delay)
                    except Exception as e:
                        print(f"Error sending message: {e}")
            print("All emails sent successfully.")
            # Here you would add the actual email sending logic using an email library like smtplib or a service API


if __name__ == "__main__":
    config = read_config("config.toml")
    print(config)
    # fname = config["recipients"]
    # tpl_fname = config["template"]

    load_dotenv()
    config["login_id"] = os.getenv("LOGIN_ID")
    config["sender_name"] = os.getenv("SENDER_NAME")
    config["password"] = os.getenv("APP_PASSWORD")
    print(
        f"Login ID: {config['login_id']}, Sender Name: {config['sender_name']}, Password: {'*' * len(config['password']) if config['password'] else None}"
    )

    df = read_recipients_data(config["recipients"], cols_dup=[], cols_sort=[])
    df["mode"] = df["att_mode"].apply(lambda x: x.strip().lower().startswith("online"))
    df = df[["email", "name", "mode"]].copy()
    print(df)
    if config["login_id"] and config["password"] and config["sender_name"]:
        send_bulk_emails(
            config["template"],
            df,
            4,
            1,
            login_id=config["login_id"],
            pwd=config["password"],
            sender_name=config["sender_name"],
            subject=config["subject"],
            pdf_fname=config["attachment"],
            dry_run=False,
        )
