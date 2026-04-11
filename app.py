import streamlit as st

from config import STYLES
from ui_pages.merge_page import render_merge_page
from ui_pages.split_page import render_split_page
from ui_pages.validasi_page import render_validasi_page

st.set_page_config(page_title="Dashboard Validasi Data Sentra Antasena", layout="wide")


def main():
    st.markdown(STYLES, unsafe_allow_html=True)

    st.sidebar.title("Menu Utama")
    menu_options = ["Validasi Data", "Split Workbook", "Merge Workbook"]
    selected_menu = st.sidebar.radio("Pilih Menu:", menu_options, index=0)

    if selected_menu == "Validasi Data":
        render_validasi_page()
    elif selected_menu == "Split Workbook":
        render_split_page()
    elif selected_menu == "Merge Workbook":
        render_merge_page()

    st.markdown(STYLES, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

