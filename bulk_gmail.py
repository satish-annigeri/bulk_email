import marimo

__generated_with = "0.19.11"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo

    import os
    from dotenv import load_dotenv
    from mako.template import Template
    from pprint import pprint
    import pandas as pd

    import bulk_mail_utils as utils

    return load_dotenv, mo, os, pd, pprint, utils


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
    pdf_fname = config['attachment']
    print(f"Template: {tpl_fname} Recipients data: {recipients_fname}")
    print(f"PDF attachment file: {pdf_fname}")
    return config, pdf_fname, recipients_fname, tpl_fname


@app.cell
def _(mo, tpl_fname):
    with open(tpl_fname, "rb") as fp:
        _txt = fp.read()
    mo.ui.code_editor(_txt.decode("utf-8"), language="html")
    return


@app.cell
def _(mo, tpl_fname, utils):
    tpl,tpl_type = utils.read_template(tpl_fname)
    # print(tpl_type)
    txt = tpl.render(**{"name": "Satish Annigeri", "mode": False})
    mo.Html(txt)
    return


@app.cell
def _(mo, pd, recipients_fname, utils):
    df = utils.read_recipients_data(recipients_fname, cols_dup=[], cols_sort=[])
    test_df = pd.DataFrame({
        "email": ["satish.annigeri@outlook.com", "satish.annigeri@outlook.com"],
        "name": ["Satish Inperson", "Satish Online"],
        "att_mode": ["In person", "Online"]
    })
    df = pd.concat([test_df, df], ignore_index=True)
    df["mode"] = df["att_mode"].str.lower().str.startswith("online")
    mo.ui.table(df)
    return (df,)


@app.cell
def _(config, df, login_id, pdf_fname, pwd, sender_name, tpl_fname, utils):
    utils.send_bulk_emails(tpl_fname, df, 1, 1, login_id=login_id, pwd=pwd, sender_name=sender_name, subject=config["subject"], pdf_fname=pdf_fname, dry_run=True)
    return


if __name__ == "__main__":
    app.run()
