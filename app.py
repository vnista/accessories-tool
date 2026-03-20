import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Honda Accessories – EN → IT", layout="centered")

# === 1. Caricamento master IT (dal repo GitHub) ===
@st.cache_data
def load_master():
    # Il file deve essere nella stessa cartella di app.py nel repo
    return pd.read_excel("Master_Accessories_Merged.xlsx")


# === 2. Logica di merge con sovrascrittura colonne EN ===
def merge_files_overwrite(df_en: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    # Seleziono solo le colonne utili dal master
    master_small = df_master[["PARTNUMBER", "DESCRIPTION", "REMARK", "MASTER IMAGE"]].copy()
    master_small = master_small.rename(
        columns={
            "DESCRIPTION": "DESCRIPTION_IT",
            "REMARK": "REMARK_IT",
            "MASTER IMAGE": "MASTER_IMAGE_IT",
        }
    )

    # Merge LEFT: mantengo tutte le righe e la struttura del file EN
    merged = df_en.merge(master_small, on="PARTNUMBER", how="left")

    # Sovrascrivo le colonne del file EN con l’italiano dove disponibile
    if "DESCRIPTION" in merged.columns:
        merged["DESCRIPTION"] = merged["DESCRIPTION_IT"].fillna(merged["DESCRIPTION"])
    if "REMARK" in merged.columns:
        merged["REMARK"] = merged["REMARK_IT"].fillna(merged["REMARK"])
    if "MASTER IMAGE" in merged.columns:
        merged["MASTER IMAGE"] = merged["MASTER_IMAGE_IT"].fillna(merged["MASTER IMAGE"])

    # Flag per vedere quali codici non sono presenti nel master IT
    merged["NOT_FOUND"] = merged["DESCRIPTION_IT"].isna()

    # Mantengo solo le colonne originali + NOT_FOUND
    cols_originali = list(df_en.columns)
    cols_output = cols_originali + ["NOT_FOUND"]
    result = merged[cols_output]

    return result


# === 3. Utility per creare il file Excel da scaricare ===
def to_excel_download(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="ACCESSORIES")
    return output.getvalue()


# === 4. App Streamlit ===
def main():
    st.title("Honda Accessories – Conversione EN → IT")

    st.markdown(
        "Carica il **file master accessori in inglese** (Excel con colonna `PARTNUMBER`).\n\n"
        "L’app userà il database interno in italiano (`Master_Accessories_Merged.xlsx`) "
        "per **sostituire le descrizioni, i remark e le eventuali master image**, "
        "mantenendo **esattamente la stessa struttura** del file originale."
    )

    # Carico il master IT una volta sola
    try:
        df_master = load_master()
        st.success(f"Database IT caricato: {len(df_master)} righe dal master.")
    except FileNotFoundError:
        st.error(
            "Impossibile trovare il file "
            "`Master_Accessories_Merged.xlsx`.\n\n"
            "Verifica che sia nel **repo GitHub**, nella **stessa cartella di app.py**, "
            "e che il nome sia esattamente identico (maiuscole, estensione)."
        )
        return

    # Upload del file EN
    uploaded_file = st.file_uploader(
        "Carica il file master accessori in inglese (.xlsx / .xls)",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            df_en = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Errore nella lettura del file caricato: {e}")
            return

        if "PARTNUMBER" not in df_en.columns:
            st.error("Il file caricato deve avere una colonna chiamata **'PARTNUMBER'**.")
            return

        st.subheader("Anteprima file caricato")
        st.dataframe(df_en.head())

        if st.button("Esegui conversione EN → IT"):
            result = merge_files_overwrite(df_en, df_master)

            st.subheader("Anteprima risultato (solo IT, stessa struttura)")
            st.dataframe(result.head())

            total = len(result)
            not_found = result["NOT_FOUND"].sum()
            st.info(f"Righe totali: {total} – Codici NON trovati nel master IT: {not_found}")

            excel_bytes = to_excel_download(result)
            st.download_button(
                label="Scarica file Excel convertito in italiano",
                data=excel_bytes,
                file_name="Master_accessories_IT_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
