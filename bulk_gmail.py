import marimo

__generated_with = "0.19.11"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo

    import os
    from dotenv import load_dotenv
    from mako.template import Template

    import bulk_mail_utils as utils

    return load_dotenv, mo, os, utils


@app.cell
def _(load_dotenv, os):
    load_dotenv()
    login_id = os.getenv("LOGIN_ID", "")
    sender_name = os.getenv("SENDER_NAME", "")
    pwd = os.getenv("APP_PASSWORD", "")
    print(f"Login ID: {sender_name} <{login_id}> {pwd}")
    return


@app.cell
def _(utils):
    config = utils.read_config("config.toml")
    tpl_fname = config["template"]
    recipients_fname = config["recipients"]
    print(f"Template: {tpl_fname} Recipients data: {recipients_fname}")
    return recipients_fname, tpl_fname


@app.cell
def _(mo, tpl_fname, utils):
    tpl,tpl_type = utils.read_template(tpl_fname)
    # print(tpl_type)
    txt = tpl.render(**{"name": "Satish Annigeri", "mode": False})
    mo.Html(txt)
    return


@app.cell
def _(recipients_fname, tpl_fname, utils):
    df = utils.read_recipients_data(recipients_fname)
    utils.send_bulk_emails(tpl_fname, df, 1, 1)
    return


if __name__ == "__main__":
    app.run()
