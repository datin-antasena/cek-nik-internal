from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials


@st.cache_data(ttl=3600)
def ambil_data_salur_gspread():
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_key(st.secrets["SPREADSHEET_ID"]).worksheet("BNBA")

        kolom_nik = sheet.col_values(4)
        set_nik_salur = {str(nik).strip() for nik in kolom_nik[1:] if nik}
        waktu_update = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%d %b %Y, %H:%M:%S WIB")

        return set_nik_salur, waktu_update

    except Exception as e:
        return set(), f"Gagal mengambil data: {e}"

