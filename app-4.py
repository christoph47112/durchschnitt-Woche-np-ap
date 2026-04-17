import streamlit as st
import pandas as pd
from io import BytesIO
import math

st.set_page_config(page_title="Durchschnittliche Abverkaufsmengen", layout="wide")
st.title("Berechnung der Ø Abverkaufsmengen pro Woche von Werbeartikeln")

# ──────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────

def fix_columns(df):
    """Benennt unnamed / None Spalten in 'Name' um (zweite Spalte)."""
    cols = list(df.columns)
    if cols[1] is None or str(cols[1]).startswith("Unnamed"):
        cols[1] = "Name"
        df.columns = cols
    return df

def detect_format(df):
    cols = set(df.columns)
    if {"Menge Aktion", "Aktionsumsatz", "Umsatz ohne Aktion", "Umsatz Menge ohne Aktion"}.issubset(cols):
        return "neu_mit_umsatz"
    elif {"Menge Aktion", "Umsatz Menge ohne Aktion"}.issubset(cols):
        return "neu_ohne_umsatz"
    elif {"Artikel", "Woche", "Menge", "Name"}.issubset(cols):
        return "alt"
    return "unbekannt"

def convert_original_file(uploaded_file):
    df_original = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    df_original.columns = df_original.iloc[1]
    df_original = df_original[2:]
    df_transformed = df_original[
        [df_original.columns[0], df_original.columns[1],
         "Woche", "VerkaufsME | Wochentag", "Gesamtergebnis"]
    ].copy()
    df_transformed.columns = ["Artikel", "Name", "Woche", "VerkaufsME", "Menge"]
    df_transformed["Woche"] = pd.to_numeric(df_transformed["Woche"], errors='coerce')
    df_transformed["Menge"] = pd.to_numeric(df_transformed["Menge"], errors='coerce')
    return df_transformed

def prepare_df(df, fmt):
    df = df.copy()
    df["Woche"] = pd.to_numeric(df["Woche"], errors='coerce')
    if fmt in ("neu_mit_umsatz", "neu_ohne_umsatz"):
        df["Menge Aktion"] = pd.to_numeric(df["Menge Aktion"], errors='coerce')
        df["Umsatz Menge ohne Aktion"] = pd.to_numeric(df["Umsatz Menge ohne Aktion"], errors='coerce')
    if fmt == "neu_mit_umsatz":
        df["Aktionsumsatz"] = pd.to_numeric(df["Aktionsumsatz"], errors='coerce')
        df["Umsatz ohne Aktion"] = pd.to_numeric(df["Umsatz ohne Aktion"], errors='coerce')
    return df

def apply_rounding(series, option):
    if option == 'Aufrunden':
        return series.apply(lambda x: math.ceil(x) if pd.notna(x) else x)
    elif option == 'Abrunden':
        return series.apply(lambda x: math.floor(x) if pd.notna(x) else x)
    elif option == 'Kaufmännisch runden':
        return series.apply(lambda x: round(x) if pd.notna(x) else x)
    return series.apply(lambda x: round(x, 2) if pd.notna(x) else x)

def to_excel_bytes(df):
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return output

# ──────────────────────────────────────────────
# Upload
# ──────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Bitte laden Sie Ihre Abverkaufsdatei hoch (Excel)",
    type=["xlsx"]
)

if uploaded_file:
    data = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Blatt auswählen", data.sheet_names)
    df_raw = data.parse(sheet_name)

    # WICHTIG: Zuerst Spalten fixen, dann Format erkennen
    df_raw = fix_columns(df_raw)
    fmt = detect_format(df_raw)

    if fmt == "unbekannt":
        st.warning("⚠️ Unbekanntes Format – die App versucht es umzuwandeln.")
        df = convert_original_file(uploaded_file)
        fmt = "alt"
    else:
        df = prepare_df(df_raw, fmt)

    if fmt == "neu_mit_umsatz":
        st.success("✅ Neues Format erkannt (Menge + Umsatz für Normal & Aktion)")
    elif fmt == "neu_ohne_umsatz":
        st.success("✅ Neues Format erkannt (nur Mengen, kein Umsatz in €)")
    else:
        st.info("📄 Altes Format erkannt (nur Normalpreise)")

    # ──────────────────────────────────────────────
    # Sidebar
    # ──────────────────────────────────────────────

    st.sidebar.title("🔍 Artikel-Filter")
    artikel_filter = st.sidebar.text_input("Nach Artikelnummer filtern (optional)")
    artikel_name_filter = st.sidebar.text_input("Nach Artikelname filtern (optional)")

    st.sidebar.title("⚙️ Rundung (Mengen)")
    round_option = st.sidebar.selectbox(
        "Rundungsoption:",
        ['Nicht runden', 'Aufrunden', 'Abrunden', 'Kaufmännisch runden'],
        index=0
    )

    if artikel_filter:
        df = df[df['Artikel'].astype(str).str.contains(artikel_filter, case=False, na=False)]
    if artikel_name_filter:
        df = df[df['Name'].str.contains(artikel_name_filter, case=False, na=False)]

    # ──────────────────────────────────────────────
    # TABS
    # ──────────────────────────────────────────────

    if fmt in ("neu_mit_umsatz", "neu_ohne_umsatz"):
        tab1, tab2 = st.tabs(["📦 Normalpreis", "🏷️ Aktionspreis"])

        # ── Tab 1: Normalpreis ──
        with tab1:
            st.subheader("Ø Abverkauf pro Woche – Normalpreis")
            st.caption("Nur Wochen ohne Aktion (Menge Aktion = leer)")

            df_normal = df[df["Menge Aktion"].isna()].copy()
            df_normal = df_normal[df_normal["Umsatz Menge ohne Aktion"].notna()]

            if df_normal.empty:
                st.warning("Keine Normalpreis-Daten gefunden.")
            else:
                agg_dict = {'Umsatz Menge ohne Aktion': 'mean', 'Woche': 'count'}
                if fmt == "neu_mit_umsatz":
                    agg_dict['Umsatz ohne Aktion'] = 'mean'

                result_normal = df_normal.groupby(
                    ['Artikel', 'Name'], sort=False
                ).agg(agg_dict).reset_index()

                result_normal['Ø Menge/Woche'] = apply_rounding(
                    result_normal['Umsatz Menge ohne Aktion'], round_option
                )

                if fmt == "neu_mit_umsatz":
                    result_normal['Ø Umsatz/Woche (€)'] = result_normal['Umsatz ohne Aktion'].round(2)
                    result_normal = result_normal[['Artikel', 'Name', 'Ø Menge/Woche', 'Ø Umsatz/Woche (€)', 'Woche']]
                else:
                    result_normal = result_normal[['Artikel', 'Name', 'Ø Menge/Woche', 'Woche']]

                result_normal.rename(columns={'Woche': 'Anzahl Wochen'}, inplace=True)

                st.dataframe(result_normal, use_container_width=True)
                st.info(f"✅ {len(result_normal)} Artikel | {df_normal.shape[0]} Normalpreis-Wochen")

                st.download_button(
                    label="📥 Normalpreis-Ergebnisse herunterladen",
                    data=to_excel_bytes(result_normal),
                    file_name="durchschnitt_normalpreis.xlsx"
                )

        # ── Tab 2: Aktionspreis ──
        with tab2:
            st.subheader("Ø Abverkauf pro Woche – Aktionspreis")
            st.caption("Nur Aktionswochen (Menge Aktion ist gefüllt)")

            df_aktion = df[df["Menge Aktion"].notna()].copy()

            if df_aktion.empty:
                st.warning("Keine Aktionsdaten gefunden.")
            else:
                agg_dict = {'Menge Aktion': 'mean', 'Woche': 'count'}
                if fmt == "neu_mit_umsatz":
                    agg_dict['Aktionsumsatz'] = 'mean'

                result_aktion = df_aktion.groupby(
                    ['Artikel', 'Name'], sort=False
                ).agg(agg_dict).reset_index()

                result_aktion['Ø Menge/Woche (Aktion)'] = apply_rounding(
                    result_aktion['Menge Aktion'], round_option
                )

                if fmt == "neu_mit_umsatz":
                    result_aktion['Ø Umsatz/Woche (€)'] = result_aktion['Aktionsumsatz'].round(2)
                    result_aktion = result_aktion[['Artikel', 'Name', 'Ø Menge/Woche (Aktion)', 'Ø Umsatz/Woche (€)', 'Woche']]
                else:
                    result_aktion = result_aktion[['Artikel', 'Name', 'Ø Menge/Woche (Aktion)', 'Woche']]

                result_aktion.rename(columns={'Woche': 'Anzahl Aktionswochen'}, inplace=True)

                st.dataframe(result_aktion, use_container_width=True)
                st.info(f"✅ {len(result_aktion)} Artikel in Aktionen | {df_aktion.shape[0]} Aktionswochen")

                st.download_button(
                    label="📥 Aktionspreis-Ergebnisse herunterladen",
                    data=to_excel_bytes(result_aktion),
                    file_name="durchschnitt_aktionspreis.xlsx"
                )

    else:
        # Altes Format
        st.subheader("Ø Abverkaufsmenge pro Woche – Normalpreis")

        result = df.groupby(['Artikel', 'Name'], sort=False).agg(
            {'Menge': 'mean'}
        ).reset_index()
        result.rename(columns={'Menge': 'Ø Menge/Woche'}, inplace=True)
        result['Ø Menge/Woche'] = apply_rounding(result['Ø Menge/Woche'], round_option)

        st.dataframe(result, use_container_width=True)
        st.info("✅ Verarbeitung abgeschlossen.")

        st.download_button(
            label="📥 Ergebnisse herunterladen",
            data=to_excel_bytes(result),
            file_name="durchschnittliche_abverkaeufe.xlsx"
        )

# ──────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────

st.markdown("---")
st.markdown("⚠️ **Hinweis:** Diese Anwendung speichert keine Daten und hat keinen Zugriff auf Ihre Dateien.")
st.markdown("🌟 **Erstellt von Christoph R. Kaiser mit Hilfe von Künstlicher Intelligenz.**")
