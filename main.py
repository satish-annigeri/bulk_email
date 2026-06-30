from typing import Annotated
import typer
from pathlib import Path
import pandas as pd


def main(tpl: str, recipients: str):
    print(f"Template file: {tpl}")
    print(f"Recipients list file: {recipients}")

    p = Path(tpl)
    if p.is_file():
        print(f"Template file '{tpl}' exists.")
        txt = p.read_text()
        print(f"Template content:\n{txt}")  # Print first
        print()
    else:
        raise FileNotFoundError(f"Template file '{tpl}' does not exist.")

    p = Path(recipients)
    if p.is_file():
        print(f"Recipients list file '{recipients}' exists.")
        df = pd.read_excel(recipients)
        print(f"Column names: {df.columns}")
        print(f"Number of rows: {len(df)}")
        print(df.head())
    else:
        raise FileNotFoundError(f"Recipients list file '{recipients}' does not exist.")


if __name__ == "__main__":
    typer.run(main)
