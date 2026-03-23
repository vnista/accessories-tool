import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Honda Accessories – Tool accessori", layout="centered")


# === 1. Caricamento master IT (dal repo GitHub) ===
@st.cache_data
def load_master():
    return pd.read_excel("Master_Accessories_Merged.xlsx")


# === 2a. Funzione conversione EN → IT ===
def merge_files_overwrite(df_en: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    master_small = df_master[["PARTNUMBER", "DESCRIPTION", "REMARK", "GROUP", "MASTER IMAGE"]].copy()
    master_small = master_small.rename(
        columns={
            "DESCRIPTION": "DESCRIPTION_IT",
            "REMARK": "REMARK_IT",
            "GROUP": "GROUP_IT",
            "MASTER IMAGE": "MASTER_IMAGE_IT",
        }
    )

    merged = df_en.merge(master_small, on="PARTNUMBER", how="left")

    if "DESCRIPTION" in merged.columns:
        merged["DESCRIPTION"] = merged["DESCRIPTION_IT"].fillna(merged["DESCRIPTION"])
    if "REMARK" in merged.columns:
        merged["REMARK"] = merged["REMARK_IT"].fillna(merged["REMARK"])
    if "GROUP" in merged.columns:
        merged["GROUP"] = merged["GROUP_IT"].fillna(merged["GROUP"])
    if "MASTER IMAGE" in merged.columns:
        merged["MASTER IMAGE"] = merged["MASTER_IMAGE_IT"].fillna(merged["MASTER IMAGE"])

    merged["NOT_FOUND"] = merged["DESCRIPTION_IT"].isna()

    cols_finali = ["PARTNUMBER", "DESCRIPTION", "REMARK", "GROUP", "MASTER IMAGE", "NOT_FOUND"]
    cols_finali = [c for c in cols_finali if c in merged.columns]

    return merged[cols_finali]


# === 2b. Funzione pulizia file già in italiano ===
def process_italian_file(df_it: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    mandatory = "Listino comprensivo di IVA, montaggio escluso"

    # Aggiunge la frase obbligatoria in fondo a ogni REMARK
    if "REMARK" in df_it.columns:
        df_it["REMARK"] = df_it["REMARK"].fillna("").astype(str).str.strip()

        def add_mandatory(text):
            if text == "":
                return mandatory
            if mandatory.lower() in text.lower():
                return text
            return f"{text} {mandatory}"

        df_it["REMARK"] = df_it["REMARK"].apply(add_mandatory)

    # Traduce GROUP usando il master IT (fix duplicati PARTNUMBER)
    if "GROUP" in df_it.columns and "PARTNUMBER" in df_it.columns:
        gmap = (
            df_master[["PARTNUMBER", "GROUP"]]
            .dropna(subset=["PARTNUMBER"])
            .drop_duplicates(subset=["PARTNUMBER"], keep="first")  # ← fix chiave
            .set_index("PARTNUMBER")["GROUP"]
        )
        df_it["GROUP"] = df_it["PARTNUMBER"].map(gmap).fillna(df_it["GROUP"])

    return df_it

# === 3. Utility per creare il file Excel da scaricare ===
def to_excel_download(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="ACCESSORIES")
    return output.getvalue()


# === 4. App Streamlit ===
def main():
    st.title("Honda Accessories – Tool accessori")

    mode = st.radio(
        "Seleziona funzione:",
        ["1) Conversione EN → IT", "2) Pulizia file IT (REMARK + GROUP)"],
        index=0,
    )

    # Carico il master IT una volta sola
    try:
        df_master = load_master()
        st.success(f"Database IT caricato: {len(df_master)} righe dal master.")
    except FileNotFoundError:
        st.error(
            "Impossibile trovare il file `Master_Accessories_Merged.xlsx`.\n\n"
            "Verifica che sia nel **repo GitHub**, nella **stessa cartella di app.py**, "
            "e che il nome sia esattamente identico (maiuscole, estensione)."
        )
        return

    # ───────────────────────────────────────────
    # MODALITÀ 1: Conversione EN → IT
    # ───────────────────────────────────────────
    if mode.startswith("1"):
        st.markdown(
            "### Conversione EN → IT\n"
            "Carica il **file master accessori in inglese**.\n\n"
            "- Le intestazioni devono essere alla **riga 10** (le prime 9 righe vengono ignorate).\n"
            "- L'app sostituirà con l'italiano le colonne `DESCRIPTION`, `REMARK`, `GROUP` e `MASTER IMAGE`.\n"
            "- L'output conterrà solo: `PARTNUMBER`, `DESCRIPTION`, `REMARK`, `GROUP`, `MASTER IMAGE` "
            "+ colonna di servizio `NOT_FOUND`."
        )

        uploaded_file = st.file_uploader(
            "Carica il file master accessori in inglese (.xlsx / .xls)",
            type=["xlsx", "xls"],
            key="upload_en",
        )

        if uploaded_file is not None:
            try:
                df_en = pd.read_excel(uploaded_file, header=9)
            except Exception as e:
                st.error(f"Errore nella lettura del file: {e}")
                return

            if "PARTNUMBER" not in df_en.columns:
                st.error(
                    "Il file (a partire dalla riga 10) deve avere una colonna chiamata **'PARTNUMBER'**."
                )
                return

            st.subheader("Anteprima file EN caricato")
            st.dataframe(df_en.head())

            if st.button("Esegui conversione EN → IT"):
                result = merge_files_overwrite(df_en, df_master)

                st.subheader("Anteprima risultato")
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

    # ───────────────────────────────────────────
    # MODALITÀ 2: Pulizia file già in italiano
    # ───────────────────────────────────────────
    else:
        st.markdown(
            "### Pulizia file IT (REMARK + GROUP)\n"
            "Carica un file **già localizzato in italiano**.\n\n"
            "- Alla colonna `REMARK` verrà aggiunta in fondo la frase obbligatoria: "
            "`\"Listino comprensivo di IVA, montaggio escluso\"`.\n"
            "- La colonna `GROUP` verrà riscritta con le voci italiane del database "
            "(es. `PACKS` → `PACCHETTI`, `INTERIOR` → `INTERNI`, ecc.).\n"
            "- L'output avrà la **stessa struttura** del file caricato, con queste due colonne aggiornate."
        )

        uploaded_file_it = st.file_uploader(
            "Carica il file accessori IT da aggiornare (.xlsx / .xls)",
            type=["xlsx", "xls"],
            key="upload_it",
        )

        if uploaded_file_it is not None:
            try:
                df_it = pd.read_excel(uploaded_file_it, header=9)
            except Exception as e:
                st.error(f"Errore nella lettura del file: {e}")
                return

            if "PARTNUMBER" not in df_it.columns:
                st.error("Il file deve avere una colonna **'PARTNUMBER'**.")
                return

            st.subheader("Anteprima file IT caricato")
            st.dataframe(df_it.head())

            if st.button("Esegui pulizia file IT"):
                result_it = process_italian_file(df_it, df_master)

                st.subheader("Anteprima risultato (REMARK + GROUP aggiornati)")
                st.dataframe(result_it.head())

                excel_bytes_it = to_excel_download(result_it)
                st.download_button(
                    label="Scarica file IT aggiornato",
                    data=excel_bytes_it,
                    file_name="Master_accessories_IT_cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


if __name__ == "__main__":
    main()
