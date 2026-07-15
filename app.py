import pandas as pd
import streamlit as st
from mako.template import Template
from mako.lookup import TemplateLookup

tpl_lookup = TemplateLookup(directories=["template"])


def tab_template(tab_tpl):
    with tab_tpl:
        tpl_file = st.file_uploader(
            "Choose a template file (*.md, *.html)", type=["md", "html"]
        )
        if tpl_file is not None:
            st.markdown("### Template Preview")
            tpl_content = tpl_file.getvalue().decode("utf-8").replace("\r\n", "\n")
            tpl_raw = Template(tpl_content)
            content = tpl_raw.render(name="Satish")
            st.write(tpl_content)
            st.markdown("### Filled Template")
            st.markdown(content, unsafe_allow_html=True)


def tab_recipients(tab_recip):
    with tab_recip:
        recipients_file = st.file_uploader(
            "Choose a recipients list file (*.xlsx)", type=["xlsx"]
        )
        if recipients_file is not None:
            st.markdown("### Recipients List Preview")
            df = pd.read_excel(recipients_file)
            st.write(df.head())


def tab_send(tab3):
    with tab3:
        st.markdown("### Send Emails")
        st.write("This feature is under development.")


st.title("Bulk Email Sender")
tab1, tab2, tab3 = st.tabs(["Template", "Recipients", "Send"])
tab_template(tab1)
tab_recipients(tab2)
tab_send(tab3)
