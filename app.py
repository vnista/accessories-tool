import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Honda Accessories Matcher", layout="centered")

@st.cache_data
def load_master():
    return pd.read_excel("Master_Accessories_Merged.xlsx.xlsx")

def merge_files(df_en: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    # Seleziono solo le colonne che mi servono dal master
    master_small = df_master[["PARTNUMBER", "DESCRIPTION", "REMARK", "MASTER IMAGE"]].copy()
    master_small = master_small.rename(
        columns={
            "DESCRIPTION": "DESCRIPTION_IT",
            "REMARK": "REMARK_IT",
            "MASTER IMAGE": "MASTER_IMAGE",
        }
    )

    # Merge left: tutte le righe EN, aggiungo dati IT dove trovo il partnumber
    merged = df_en.merge(master_small, on="PARTNUMBER", how="left")

    # Flag NOT_FOUND
    merged["NOT_FOUND"] = merged["DESCRIPTION_IT"].isna()

    return merged

def to_excel_download(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="ACCESSORIES")
    return output.getvalue()

def main():
    st.title("Honda Accessories – EN → IT matcher")

    st.markdown(
        "Carica il **master EN** (Excel con colonna `PARTNUMBER`) e il tool "
        "aggiungerà descrizione IT, remark IT e master image dal database."
    )

    df_master = load_master()
    st.success(f"Database IT caricato: {len(df_master)} righe")

    uploaded_file = st.file_uploader(
        "Carica il file master EN (.xlsx)", type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            df_en = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Errore nella lettura del file: {e}")
            return

        if "PARTNUMBER" not in df_en.columns:
            st.error("Il file caricato deve avere una colonna chiamata 'PARTNUMBER'.")
            return

        st.write("Anteprima file EN:")
        st.dataframe(df_en.head())

        if st.button("Esegui match"):
            result = merge_files(df_en, df_master)

            st.subheader("Risultato (anteprima)")
            st.dataframe(result.head())

            total = len(result)
            not_found = result["NOT_FOUND"].sum()
            st.info(f"Righe totali: {total} – Non trovate nel master: {not_found}")

            excel_bytes = to_excel_download(result)
            st.download_button(
                label="Scarica risultato in Excel",
                data=excel_bytes,
                file_name="Master_EN_IT_with_images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if __name__ == "__main__":
    main()
