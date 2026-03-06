import marimo

__generated_with = "0.20.2"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo

    import os
    from pathlib import Path
    from markdown import markdown
    from dotenv import load_dotenv
    from mako.template import Template
    from pprint import pprint
    import pandas as pd

    import bulk_mail_utils as utils

    return Path, Template, load_dotenv, markdown, mo, os, pd, pprint, utils


@app.cell
def _(load_dotenv, os):
    load_dotenv()
    login_id = os.getenv("LOGIN_ID", "")
    sender_name = os.getenv("SENDER_NAME", "")
    pwd = os.getenv("APP_PASSWORD", "")
    print(f"Login ID: {sender_name} <{login_id}> {pwd}")
    return login_id, pwd, sender_name


@app.cell
def _(pprint, utils):
    config = utils.read_config("config.toml")
    pprint(config, indent=4)
    tpl_fname = config["template"]
    recipients_fname = config["recipients"]
    pdf_fname = config["attachment"]
    print(f"Template: {tpl_fname} Recipients data: {recipients_fname}")
    print(f"PDF attachment file: {pdf_fname}")
    return config, pdf_fname, recipients_fname, tpl_fname


@app.cell
def _(mo):
    fname_tpl = mo.ui.file_browser(filetypes=[".md", ".html"])
    fname_tpl
    return (fname_tpl,)


@app.cell
def _(Path, Template, fname_tpl, markdown, mo):
    _fname = fname_tpl.path()
    _fpath = Path(_fname)
    _tpl_txt = _fpath.read_text(encoding="utf-8")
    if _fpath.suffix == ".md":
        _tpl_txt = markdown(_tpl_txt)
    _tpl = Template(_tpl_txt)
    _recipient = {"name": "Satish Annigeri", "email": "satish.annigeri@gmail.com", "mode": True}
    _html = _tpl.render(**_recipient)
    mo.Html(_html)
    mo.md(f"Email will be sent to: **<{_recipient['email']}>**")
    return


@app.cell
def _(mo):
    fname_recipients = mo.ui.file_browser(filetypes=[".xlsx", ".xls", ".csv"])
    fname_recipients
    return (fname_recipients,)


@app.cell
def _(Path, fname_recipients, pd):
    print(fname_recipients.path())
    _fname = fname_recipients.path()
    _fpath = Path(_fname)
    if _fpath.suffix in [".xlsx", ".xls"]:
        _df = pd.read_excel(_fname)
    elif _fpath.suffix in [".csv"]:
        _df = pd.read_csv(_fname)
    print(_fpath.suffix)
    _df = _df[["name", "email"]]
    _df
    return


@app.cell
def _(mo, tpl_fname):
    with open(tpl_fname, "rb") as fp:
        _txt = fp.read()
    mo.ui.code_editor(_txt.decode("utf-8"), language="html")
    return


@app.cell
def _(mo, tpl_fname, utils):
    tpl, tpl_type = utils.read_template(tpl_fname)
    print(tpl.source)
    txt = utils.tpl_render(tpl, tpl_type, **{"name": "Satish Annigeri", "mode": False})
    mo.md(txt)
    return


@app.cell
def _(mo, pd, recipients_fname, utils):
    df = utils.read_recipients_data(recipients_fname, cols_dup=[], cols_sort=[])
    test_df = pd.DataFrame(
        {
            "email": ["satish.annigeri@outlook.com"],
            "name": ["Satish Annigeri"],
        }
    )
    df = pd.concat([test_df, df], ignore_index=True)[["email", "name"]]
    # df["mode"] = df["att_mode"].str.lower().str.startswith("online")
    mo.ui.table(df)
    return (df,)


@app.cell
def _(config, df, login_id, pdf_fname, pwd, sender_name, tpl_fname, utils):
    utils.send_bulk_emails(
        tpl_fname,
        df,
        1,
        1,
        login_id=login_id,
        pwd=pwd,
        sender_name=sender_name,
        subject=config["subject"],
        pdf_fname=pdf_fname,
        dry_run=True,
    )
    return


if __name__ == "__main__":
    app.run()
