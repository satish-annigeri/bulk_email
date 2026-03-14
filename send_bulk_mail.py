import marimo

__generated_with = "0.20.4"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo

    import os
    import time
    from pathlib import Path
    from markdown import markdown
    from dotenv import load_dotenv
    from mako.template import Template
    from pprint import pprint
    import pandas as pd
    import smtplib

    import bulk_mail_utils as utils

    return (
        Path,
        Template,
        load_dotenv,
        markdown,
        mo,
        os,
        pd,
        smtplib,
        time,
        utils,
    )


@app.cell
def _(load_dotenv, os):
    load_dotenv()
    login_id = os.getenv("LOGIN_ID", "")
    sender_name = os.getenv("SENDER_NAME", "")
    pwd = os.getenv("APP_PASSWORD", "")
    print(f"Login ID: {sender_name} <{login_id}> {pwd}")
    return login_id, pwd, sender_name


@app.cell
def _():
    # config = utils.read_config("config.toml")
    # pprint(config, indent=4)
    # tpl_fname = config["template"]
    # recipients_fname = config["recipients"]
    # pdf_fname = config["attachment"]
    # print(f"Template: {tpl_fname} Recipients data: {recipients_fname}")
    # print(f"PDF attachment file: {pdf_fname}")
    return


@app.cell
def _(mo):
    fname_tpl = mo.ui.file_browser(filetypes=[".md", ".html"], multiple=False)
    fname_tpl
    return (fname_tpl,)


@app.cell
def _(Template, markdown):
    def tpl_render(tpl_txt, md_fmt: bool, **kwargs):
        tpl = Template(tpl_txt)
        html = tpl.render(**kwargs)
        if md_fmt:
            html = markdown(html)
        return html

    return (tpl_render,)


@app.cell
def _(mo):
    subject = mo.ui.text(label="Subject:", max_length=120, value="PLSG Session 1 - Saturday March 7, 2026 from 4:00 pm to 7:00 pm")
    test_name = mo.ui.text(label="Sender's name", value="Satish Annigeri")
    test_email = mo.ui.text(label="Sender's email", value="satish.annigeri@outlook.com")
    test_send = mo.ui.checkbox(label="Send test email")
    mo.vstack([mo.hstack([test_name, test_email], align="start"), mo.hstack([subject], align="start")])
    return subject, test_email, test_name, test_send


@app.cell
def _(Path, fname_tpl, mo, test_email, test_name, tpl_render):
    _fname = fname_tpl.path()
    _fpath = Path(_fname)
    md_fmt = _fpath.suffix == ".md"
    tpl_txt = _fpath.read_text(encoding="utf-8")

    _html = tpl_render(tpl_txt, md_fmt, name=test_name.value, email=test_email.value)
    mo.md(_html)
    return md_fmt, tpl_txt


@app.cell
def _(test_send):
    test_send
    return


@app.cell
def _(
    login_id,
    md_fmt,
    pwd,
    sender_name,
    smtplib,
    subject,
    test_email,
    test_name,
    test_send,
    tpl_render,
    tpl_txt,
    utils,
):
    if test_send.value and subject.value and test_name.value and test_email.value:
        print("Sending test email")
        print(f"Subject: {subject.value}")

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(login_id, pwd)
            try:
                _html = tpl_render(tpl_txt, md_fmt, name=test_name.value, email=test_email.value)
                msg = utils.build_message(
                    sender_name=sender_name,
                    sender_email=login_id,
                    recipient_name=test_name.value,
                    recipient_email=test_email.value,
                    subject=subject.value, body=_html
                )
                server.send_message(msg)
                print(f"{test_name.value} <{test_email.value}>")
            except Exception as e:
                print(f"Error sending message: {e}")
    else:
        print("Dry run")
        print(f"Subject: {subject.value}")
        print(f"{test_name.value} <{test_email.value}>")
    return


@app.cell
def _(mo):
    fname_recipients = mo.ui.file_browser(filetypes=[".xlsx", ".xls", ".csv"], multiple=False)
    fname_recipients
    return (fname_recipients,)


@app.cell
def _(Path, fname_recipients, mo, pd):
    _fname = fname_recipients.path()
    _fpath = Path(_fname)
    print(_fpath)
    if _fpath.suffix in [".xlsx", ".xls"]:
        _df = pd.read_excel(_fname)
    elif _fpath.suffix in [".csv"]:
        _df = pd.read_csv(_fname)
    _df = _df[["name", "email"]]
    df = _df.copy()
    mo.ui.table(df)
    return (df,)


@app.cell
def _(mo):
    send_bulk = mo.ui.checkbox(label="Send bulk email")
    send1 = mo.ui.number(label="First:", start=1)
    count = mo.ui.number(label="Count:", value=1, start=-1)
    mo.hstack([send_bulk, send1, count])
    return count, send1, send_bulk


@app.cell
def _(
    count,
    df,
    login_id,
    md_fmt,
    pwd,
    send1,
    send_bulk,
    sender_name,
    smtplib,
    subject,
    time,
    tpl_render,
    tpl_txt,
    utils,
):
    i1 = int(send1.value)
    _count = int(count.value)
    if _count > 0:
        _send2 = i1 +  _count - 1
    else:
        _send2 = i1 + len(df) - 1
    i2 = min(_send2, len(df))
    r = range(i1-1, i2)
    # print(r)

    if send_bulk.value:
        print(subject.value)

        with smtplib.SMTP("smtp.gmail.com", 587) as _server:
            _server.starttls()
            _server.login(login_id, pwd)
            print("Connected to SMTP server")
            for _i in range(i1 - 1, i2):
                _name = df.iloc[_i]["name"]
                _email = df.iloc[_i]["email"]
                try:
                    # _name = test_name.value
                    # _email = test_email.value
                    print(f"{_i+1:4d}: {_name} <{_email}>", end="...")
                    _html = tpl_render(tpl_txt, md_fmt, name=_name, email=_email)
                    _msg = utils.build_message(
                        sender_name=sender_name,
                        sender_email=login_id,
                        recipient_name=_name,
                        recipient_email=_email,
                        subject=subject.value, body=_html
                    )
                    _server.send_message(_msg)
                    time.sleep(0.5)
                    print()
                except Exception as e:
                    print(f"Error sending message: {e}")
            print("Disconnected from SMTP server")
    else:
        print("Dry run. Emails are not being sent")
        print(f"Subject: {subject.value}")
        for _i in range(i1 - 1, i2):
            # print(_i+1)
            _name = df.iloc[_i]["name"]
            _email = df.iloc[_i]["email"]
            print(f"{_i+1:4d}: {_name} <{_email}>")
    return


if __name__ == "__main__":
    app.run()
