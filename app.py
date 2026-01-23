import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Cek Validitas Data NIK", layout="wide")

st.title("🔍 Alat Validasi NIK Excel")
st.markdown("Upload file Excel, pilih kolom, dan sistem akan mengecek: Panjang 16, Format Angka, Akhiran '00', dan Duplikasi.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        st.subheader("Preview Data Asli")
        st.dataframe(df.head(), use_container_width=True)
        
        columns = df.columns.tolist()
        target_col = st.selectbox("Pilih Kolom NIK yang akan dicek:", columns)
        
        if st.button("Proses Cek Data"):
            df_result = df.copy()
            # Pastikan format string
            df_result[target_col] = df_result[target_col].astype(str).replace('nan', '')
            
            # Hitung urutan kemunculan untuk logika Ganda
            df_result['__temp_count'] = df_result.groupby(target_col).cumcount() + 1
            
            def check_nik(row):
                val = row[target_col]
                count = row['__temp_count']
                val = val.replace('.0', '').strip()
                
                # 1. Cek Panjang
                if len(val) != 16:
                    return "NIK TIDAK 16 DIGIT"
                # 2. Cek Angka (ISERR)
                elif not val.isdigit():
                    return "BUKAN ANGKA (ADA HURUF/SIMBOL)"
                # 3. Cek Akhiran 00
                elif val.endswith("00"):
                    return "NIK TERKONVENSI"
                # 4. Cek Unik
                elif count == 1:
                    return "UNIK"
                # 5. Sisanya Ganda
                else:
                    return f"GANDA {count}"

            df_result['STATUS_CEK'] = df_result.apply(check_nik, axis=1)
            df_result.drop(columns=['__temp_count'], inplace=True)
            
            st.success("Pengecekan Selesai!")
            st.dataframe(df_result, use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False)
                
            st.download_button(
                label="📥 Download Hasil",
                data=buffer.getvalue(),
                file_name=f"Hasil_{uploaded_file.name}",
                mime="application/vnd.ms-excel"
            )
            
    except Exception as e:
        st.error(f"Error: {e}")