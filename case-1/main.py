#!/usr/bin/env python

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import argparse

def fixed_table(location):
    raw_data = pd.read_excel(location, sheet_name='Current - ii01', skiprows=4)
    
    # Kelompok Bank
    main_table_bank = pd.DataFrame(columns=['Aktiva Rupiah', 'Valuta Asing', 'Jumlah Aset'])

    # Kelompok Dati II (Rp dan Valas)
    main_table_dati = pd.DataFrame(columns=['Jumlah Aset'])

    for i in range(5, len(raw_data.columns)):
        ext_table_bank = pd.DataFrame({
            'Aktiva Rupiah': raw_data.iloc[3:7, i].values,
            'Valuta Asing': raw_data.iloc[9:13, i].values,
            'Jumlah Aset': raw_data.iloc[15:19, i].values
        })
        main_table_bank = pd.concat([main_table_bank, ext_table_bank], ignore_index=True)

        ext_table_dati = pd.DataFrame({
            'Jumlah Aset': raw_data.iloc[21:60, i].values
        })
        main_table_dati = pd.concat([main_table_dati, ext_table_dati], ignore_index=True)
    
    # Tanggal Kelompok Bank
    tanggal_bank = []
    start_date_bank = datetime(2023, 1, 1)
    for bulan_bank in range(13):
        tahun_bulan = start_date_bank.strftime('%Y-%m')
        tanggal_bank.extend([tahun_bulan] * 4)
        start_date_bank = (start_date_bank.replace(day=1) + timedelta(days=32)).replace(day=1)

    # Tanggal Kelompok Dati II
    tanggal_dati = []
    start_date_dati = datetime(2023, 1, 1)
    for bulan_dati in range(13):
        tahun_bulan = start_date_dati.strftime('%Y-%m')
        tanggal_dati.extend([tahun_bulan] * 39)
        start_date_dati = (start_date_dati.replace(day=1) + timedelta(days=32)).replace(day=1)
    
    # Tipe Komponen Kelompok Bank
    tipe_komponen_bank = ['Bank Pemerintah', 'Bank Swasta Nasional', 'Bank Asing dan Bank Campuran', 'Bank Perkreditan Rakyat']
    tipe_komponen_bank = 13 * tipe_komponen_bank

    # Tipe Komponen Kelompok Dati II
    tipe_komponen_dati = [vals for vals in raw_data.iloc[21:60, 2].values]
    tipe_komponen_dati = 13 * tipe_komponen_dati

    main_table_bank['Tanggal'] = tanggal_bank
    main_table_bank['Tipe Komponen'] = tipe_komponen_bank

    main_table_dati['Tanggal'] = tanggal_dati
    main_table_dati['Tipe Komponen'] = tipe_komponen_dati

    columns_bank = ['Tanggal', 'Tipe Komponen', 'Aktiva Rupiah', 'Valuta Asing', 'Jumlah Aset']
    main_table_bank = main_table_bank[columns_bank]

    columns_dati = ['Tanggal', 'Tipe Komponen', 'Jumlah Aset']
    main_table_dati = main_table_dati[columns_dati]

    # Export DataFrames
    with pd.ExcelWriter('./dataset/Fixed Dataset.xlsx') as writer:
        main_table_bank.to_excel(writer, sheet_name='Kelompok Bank', index=False)
        main_table_dati.to_excel(writer, sheet_name='Kelompok Dati II', index=False)

    return f"Data has been exported to 'Fixed Dataset.xlsx'"

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process an Excel file and generate a fixed dataset.")
    parser.add_argument("location", help="Path to the Excel file")
    args = parser.parse_args()

    result = fixed_table(args.location)
    print(result)