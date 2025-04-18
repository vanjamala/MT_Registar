import streamlit as st
import pandas as pd
from io import BytesIO
import re
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("🎈 Usporedba MasterTeam Registar")
st.write(
    "Usporedi MasterTeam i Registar evidenciju sati rada."
)
# Upload files
uploaded_masterteam = st.file_uploader("Učitajte MasterTeam evidenciju", type=["xls", "xlsx"])
uploaded_registar = st.file_uploader("Učitajte Registar datoteku", type=["xls", "xlsx"])

# Process files if uploaded and inputs are valid
if uploaded_masterteam and uploaded_registar and st.button('Spoji podatke i pripremi izvještaj'):
    df = pd.read_excel(uploaded_masterteam, header=0, engine='openpyxl', dtype={'Oib': str})
    df_r = pd.read_excel(uploaded_registar, header=11, engine='xlrd', dtype={'OIB': str})
    def transform_registar(df_r):
        # Split zaposlenik by comma
        df_r[['Zaposlenik', 'OIB']] = df_r.iloc[:, 0].str.split(',', expand=True)
        df_r["Element plaće"] = df_r["Element plaće"].replace('', np.nan).fillna("Ukupno")
        df_r[["Zaposlenik", "Ustrojstvena jedinica", "Radno mjesto", "OIB"]] = df_r[["Zaposlenik", "Ustrojstvena jedinica", "Radno mjesto", "OIB"]].fillna(method="ffill")
        # Step 1: Create a new column "Element plaće MT" by copying "Element plaće"
        df_r["Element plaće MT"] = df_r["Element plaće"]
        # Step 2: Apply the transformations to the new column "Element plaće MT"
        df_r["Element plaće MT"] = df_r.apply(lambda row: (
            "R" if str(row["Element plaće MT"]).startswith("101010") else
            "B" if str(row["Element plaće MT"]).startswith("13005") else
            "H" if str(row["Element plaće MT"]).startswith("23") or str(row["Element plaće MT"]).startswith("226") else
            "RK" if str(row["Element plaće MT"]).startswith("191012") else
            "G" if str(row["Element plaće MT"]).startswith("10750") else
            "D" if str(row["Element plaće MT"]).startswith("101040") or str(row["Element plaće MT"]).startswith("101796") else
            "SD" if str(row["Element plaće MT"]).startswith("101930") else
            "P8" if str(row["Element plaće MT"]).startswith("101020") else
        row["Element plaće MT"]  # Default: keep the original value
        ), axis=1)
        return df_r
    # Apply transformation
    df_r = transform_registar(df_r)

    # Step 3: Aggregate by 'Oib', 'Šifra tipa prisustva', and 'Prezime i Ime'
    df_r = (
        df_r.groupby(['OIB', 'Element plaće MT', 'Zaposlenik'], as_index=False)
        .agg(
            Broj_sati =('Broj sati', 'sum')  # Sum the Sati za plaću
        )
    )
    df_r = df_r[~df_r['Element plaće MT'].str.startswith('1112')]
    # Transform MasterTeam file
    # Step 1: Concatenate 'Prezime' and 'Ime' to create 'Prezime i Ime'
    df['Prezime i Ime'] = df['Prezime'] + ' ' + df['Ime']

    # Step 2: Convert 'Datum' to datetime format (if it's not already)
    df['Datum'] = pd.to_datetime(df['Datum'])
    df = df[~df['Dan'].isin(['Subota', 'Nedjelja'])]
    #df = df[~df['Šifra tipa prisustva'].isin(['O', 'H', 'B'])]
    df = df[~df['Šifra tipa prisustva'].isin(['O'])]
    df.loc[(df['Blagdan (neradno)'] == 8) & (df['Šifra tipa prisustva'] == 'R'), 'Šifra tipa prisustva'] = 'P8'



    # Step 3: Aggregate by 'Oib', 'Šifra tipa prisustva', and 'Prezime i Ime'
    df_MT = (
        df.groupby(['Oib', 'Šifra tipa prisustva', 'Prezime i Ime'], as_index=False)
        .agg(
            Ukupni_sati_za_placu =('Sati za plaću', 'sum'),  # Sum the Sati za plaću
            min_date=('Datum', 'min'),                      # Min date in the period
            max_date=('Datum', 'max')                       # Max date in the period
        )
    )

    # Step 4: Create the 'Period' column as the range between min and max date
    df_MT['Period'] = df_MT['min_date'].astype(str) + ' do ' + df_MT['max_date'].astype(str)

    #Merge files and create the mismatch report

    # Step 1: Ensure that both Oib and OIB columns are treated as strings
    df_MT['Oib'] = df_MT['Oib'].astype(str).str.strip()  # Convert to string and strip any extra spaces
    df_r['OIB'] = df_r['OIB'].astype(str).str.strip()    # Convert to string and strip any extra spaces
    # Drop rows where 'Element plaće' is 'Ukupno' in df_r
    df_r = df_r[df_r['Element plaće MT'] != 'Ukupno']

    # Step 1: Merge the DataFrames based on Oib & Šifra tipa prisustva (from df_MT)
    # and OIB & Element plaće (from df_r), with suffixes for identification
    df_merged = pd.merge(
        df_MT, df_r,
        left_on=['Oib', 'Šifra tipa prisustva'],  # Columns in df_MT
        right_on=['OIB', 'Element plaće MT'],  # Columns in df_r
        how='outer',  # Use 'outer' to keep both matches and mismatches
        suffixes=('_MasterTeam', '_Registar')
    )

    # Step 2: Add a "mismatch" flag (True where there is no match for 'Šifra tipa prisustva')
    df_merged['mismatch'] = df_merged['Šifra tipa prisustva'].isna()
    # Step 3: Compare the values (Ukupni_sati_za_placu and Broj sati)
    df_merged['sati_match'] = df_merged['Ukupni_sati_za_placu'] == df_merged['Broj_sati']
    # Rename specific columns manually
    df_merged = df_merged.rename(columns={
        'Oib': 'Oib_MasterTeam',
        'Šifra tipa prisustva': 'Šifra tipa prisustva_MasterTeam',
        'Prezime i Ime': 'Prezime i Ime_MasterTeam',
        'Ukupni_sati_za_placu': 'Ukupni_sati_za_placu_MasterTeam',
        'Period': 'Period_MasterTeam',

        'OIB': 'OIB_Registar',
        'Element plaće MT': 'Element plaće MT_Registar',
        'Zaposlenik': 'Zaposlenik_Registar',
        'Broj_sati': 'Broj_sati_Registar'
    })

    # Step 4: Filter for only matching rows (where no mismatch and hours match)
    df_matches = df_merged[(df_merged['mismatch'] == False) & (df_merged['sati_match'] == True)]
    # Step 5: Filter for mismatches (either mismatch in Šifra tipa prisustva or hours)
    df_mismatches = df_merged[(df_merged['mismatch'] == True) | (df_merged['sati_match'] == False)]
    df_mismatches = df_mismatches.drop(columns=['min_date', 'max_date', 'mismatch', 'sati_match'])
    # Report
    st.write(df_mismatches)

    # Allow downloading the filtered report 2
    output_mismatches = BytesIO()
    with pd.ExcelWriter(output_mismatches, engine='xlsxwriter') as writer:
        df_mismatches.to_excel(writer, index=False, sheet_name="Usporedba MasterTeam Registar")
    output_mismatches.seek(0)

    st.download_button(
        label="Preuzmi usporedbu evidencija",
        data=output_mismatches,
        file_name="Usporedba MasterTeam Registar.xlsx",
        mime="application/vnd.ms-excel"
    )