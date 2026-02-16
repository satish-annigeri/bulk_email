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


def tpl_render(tpl_fname: str, **kwargs) -> str:
    ext = splitext(tpl_fname)[1].lower()
    if ext not in [".html", ".htm", ".md"]:
        raise ValueError(f"Unsupported template file type: {ext}")

    tpl = Template(filename=tpl_fname)
    txt = tpl.render(**kwargs)
    if isinstance(txt, bytes):
        txt = txt.decode("utf-8")
    else:
        txt = str(txt)

    if ext == ".md":
        return markdown(txt)
    else:
        return txt


def send_bulk_emails(config: dict[str, str], dry_run: bool = True) -> None:
    df = read_data_file(config["recipients"])
    df = clean_data(df, cols_dup=["email", "name"], cols_sort=["name"])
    df = mangle_name(df, "name")
    if dry_run:
        for _, row in df.iterrows():
            html = tpl_render(config["template"], name=row["name"], mode=True)
            start, _, _ = html.partition("</p>")
            print(f"{row['name']} <{row['email']}> {start}</p>...")
        return
    else:
        for _, row in df.iterrows():
            html = tpl_render(config["template"], name=row["name"], mode=False)
            start, _, _ = html.partition("</p>")
            print(f"{row['name']} <{row['email']}> with content: {start}</p>...")
            # Here you would add the actual email sending logic using an email library like smtplib or a service API


if __name__ == "__main__":
    config = read_config("config.toml")
    print(config)
    fname = config["recipients"]
    tpl_fname = config["template"]

    load_dotenv()
    config["login_id"] = os.getenv("LOGIN_ID")
    config["sender_name"] = os.getenv("SENDER_NAME")
    config["password"] = os.getenv("APP_PASSWORD")
    print(
        f"Login ID: {config['login_id']}, Sender Name: {config['sender_name']}, Password: {'*' * len(config['password']) if config['password'] else None}"
    )

    send_bulk_emails(config)
    # df = read_data_file(fname)
    # print(len(df))
    # df = clean_data(df, cols_dup=["email", "name"], cols_sort=["name"])
    # print(len(df))
    # df = mangle_name(df, "name")
    # print(df[["email", "name"]])
    # html = tpl_render(tpl_fname, name="Satish Annigeri", mode=True)
    # print(html)
    # html = tpl_render(tpl_fname, name="Satish Annigeri", mode=False)
    # print(html)
