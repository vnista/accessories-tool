import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Honda Accessories – EN → IT", layout="centered")

# === 1. Caricamento master IT (dal repo GitHub) ===
@st.cache_data
def load_master():
    # Il file deve essere nella stessa cartella di app.py nel repo
    return pd.read_excel("Master_Accessories_Merged.xlsx")


# === 2. Logica di merge con sovrascrittura colonne EN (incl. GROUP) ===
def merge_files_overwrite(df_en: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    # Seleziono le colonne utili dal master, inclusa GROUP
    master_small = df_master[["PARTNUMBER", "DESCRIPTION", "REMARK", "GROUP", "MASTER IMAGE"]].copy()
    master_small = master_small.rename(
        columns={
            "DESCRIPTION": "DESCRIPTION_IT",
            "REMARK": "REMARK_IT",
            "GROUP": "GROUP_IT",
            "MASTER IMAGE": "MASTER_IMAGE_IT",
        }
    )

    # Merge LEFT: mantengo tutte le righe del file EN
    merged = df_en.merge(master_small, on="PARTNUMBER", how="left")

    # Sovrascrivo le colonne del file EN con l’italiano dove disponibile
    if "DESCRIPTION" in merged.columns:
        merged["DESCRIPTION"] = merged["DESCRIPTION_IT"].fillna(merged["DESCRIPTION"])
    if "REMARK" in merged.columns:
        merged["REMARK"] = merged["REMARK_IT"].fillna(merged["REMARK"])
    if "GROUP" in merged.columns:
        merged["GROUP"] = merged["GROUP_IT"].fillna(merged["GROUP"])
    if "MASTER IMAGE" in merged.columns:
        merged["MASTER IMAGE"] = merged["MASTER_IMAGE_IT"].fillna(merged["MASTER IMAGE"])

    # Flag per capire quali codici non sono nel master
    merged["NOT_FOUND"] = merged["DESCRIPTION_IT"].isna()

    # Colonne finali richieste
    cols_finali = ["PARTNUMBER", "DESCRIPTION", "REMARK", "GROUP", "MASTER IMAGE", "NOT_FOUND"]
    cols_finali = [c for c in cols_finali if c in merged.columns]

    result = merged[cols_finali]
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
        "Carica il **file master accessori in inglese**.\n\n"
        "- Il file ha intestazioni a partire dalla **riga 10** (le prime 9 righe sono da ignorare).\n"
        "- L’app userà `Master_Accessories_Merged.xlsx` per sostituire in italiano "
        "`DESCRIPTION`, `REMARK`, `GROUP` e `MASTER IMAGE`.\n"
        "- L’output conterrà solo: `PARTNUMBER`, `DESCRIPTION`, `REMARK`, `GROUP`, `MASTER IMAGE` "
        "più una colonna di servizio `NOT_FOUND`."
    )

    # Carico il master IT
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
            # header=9 → considera la riga 10 come riga intestazioni (0‑based)
            df_en = pd.read_excel(uploaded_file, header=9)
        except Exception as e:
            st.error(f"Errore nella lettura del file caricato: {e}")
            return

        if "PARTNUMBER" not in df_en.columns:
            st.error("Il file (a partire dalla riga 10) deve avere una colonna chiamata **'PARTNUMBER'**.")
            return

        st.subheader("Anteprima file caricato (da riga 10 in poi)")
        st.dataframe(df_en.head())

        if st.button("Esegui conversione EN → IT"):
            result = merge_files_overwrite(df_en, df_master)

            st.subheader("Anteprima risultato (solo colonne richieste)")
            st.dataframe(result.head())

            total = len(result)
            not_found = result["NOT_FOUND"].sum() if "NOT_FOUND" in result.columns else 0
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
