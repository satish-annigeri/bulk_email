import os
from os.path import splitext, isfile
import tomllib

from dotenv import load_dotenv
import pandas as pd
from mako.template import Template
from markdown import markdown


def read_config(fname_toml: str) -> dict:
    if not isfile(fname_toml):
        raise FileNotFoundError(f"File not found: {fname_toml}")
    with open(fname_toml, "rb") as f:
        config = tomllib.load(f)
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


def read_recipients_data(recipients_fname: str) -> pd.DataFrame:
    df = read_data_file(recipients_fname)
    df = clean_data(df, cols_dup=["email", "name"], cols_sort=["name"])
    df = mangle_name(df, "name")
    return df


def send_bulk_emails(
    tpl_fname: str,
    df: pd.DataFrame,
    start: int = 1,
    count: int = -1,
    dry_run: bool = True,
) -> None:
    tpl, tpl_type = read_template(tpl_fname)

    # df = read_data_file(config["recipients"])
    # df = clean_data(df, cols_dup=["email", "name"], cols_sort=["name"])
    # df = mangle_name(df, "name")

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
        for _, row in df.iterrows():
            html = tpl_render(tpl, tpl_type, name=row["name"], mode=False)
            start_portion, _, _ = html.partition("</p>")
            print(
                f"{row['name']} <{row['email']}> with content: {start_portion}</p>..."
            )
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

    df = read_recipients_data(config["recipients"])
    send_bulk_emails(config["template"], df, 2, 1)
