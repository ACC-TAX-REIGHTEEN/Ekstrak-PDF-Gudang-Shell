import pandas as pd
import openpyxl
import glob
import os
from datetime import datetime

def format_date_source(date_val):
    try:
        if isinstance(date_val, str):
            dt_obj = pd.to_datetime(date_val)
            return dt_obj.strftime('%d/%m/%Y')
        elif isinstance(date_val, datetime):
            return date_val.strftime('%d/%m/%Y')
        return None
    except:
        return None

def format_date_target(date_val):
    try:
        if isinstance(date_val, datetime):
            return date_val.strftime('%d/%m/%Y')
        elif isinstance(date_val, str):
            return date_val 
        return str(date_val)
    except:
        return str(date_val)

print("--> Memulai proses...")

tolerance = 0.0
conf_filename = "lookup.conf"

if os.path.exists(conf_filename):
    try:
        with open(conf_filename, "r") as f:
            content = f.read().strip()
            tolerance = float(content)
        print(f"--> Konfigurasi ditemukan: Toleransi selisih DPP diatur sebesar ±{tolerance} rupiah.")
    except Exception as e:
        print(f"--> Gagal membaca {conf_filename}, default ke toleransi = 0 (Exact Match). Error: {e}")
else:
    print(f"--> File {conf_filename} tidak ditemukan. Default ke toleransi = 0 (Exact Match).")

source_files = glob.glob("data_export_*.xlsx")
if not source_files:
    print("--> Error: File data_export_.....xlsx tidak ditemukan.")
    exit()

source_file = source_files[0]
print(f"--> File sumber ditemukan: {source_file}")

try:
    df_source = pd.read_excel(source_file, sheet_name="data", dtype={'Nomor Faktur Pajak': str})
except Exception as e:
    print(f"--> Error saat membaca file sumber: {e}")
    exit()

lookup_data = {}

print("--> Membuat indeks data dari file sumber (Sistem Antrean)...")
for index, row in df_source.iterrows():
    raw_date = row.get('Tanggal Faktur Pajak')
    dpp_val = row.get('Harga Jual/Penggantian/DPP')
    no_fp = row.get('Nomor Faktur Pajak')
    
    clean_date = format_date_source(raw_date)
    
    if clean_date and pd.notna(dpp_val):
        dpp_float = float(dpp_val)
        
        if clean_date not in lookup_data:
            lookup_data[clean_date] = {}
            
        if dpp_float not in lookup_data[clean_date]:
            lookup_data[clean_date][dpp_float] = []
            
        lookup_data[clean_date][dpp_float].append(no_fp)

target_files = glob.glob("Laporan SHELL*.xlsx")
if not target_files:
    print("--> Error: File Laporan SHELL........xlsx tidak ditemukan.")
    exit()

target_filename = target_files[0]
print(f"--> Membuka file target: {target_filename}")
wb = openpyxl.load_workbook(target_filename)

for sheet in wb.worksheets:
    print(f"--> Memproses Sheet: {sheet.title}")
    
    col_tanggal = None
    col_dpp = None
    col_no_fp = None
    match_count = 0
    
    for row in sheet.iter_rows():
        row_values = [str(cell.value).strip() if cell.value is not None else "" for cell in row]
        
        if "Tanggal" in row_values and "DPP" in row_values and "No FP" in row_values:
            col_tanggal = row_values.index("Tanggal")
            col_dpp = row_values.index("DPP")
            col_no_fp = row_values.index("No FP")
            continue
            
        if col_tanggal is not None and col_dpp is not None and col_no_fp is not None:
            cell_tanggal = row[col_tanggal]
            cell_dpp = row[col_dpp]
            cell_no_fp = row[col_no_fp]
            
            val_tanggal = cell_tanggal.value
            val_dpp = cell_dpp.value
            
            if val_tanggal is not None and val_dpp is not None:
                target_date_str = format_date_target(val_tanggal)
                try:
                    target_dpp_float = float(val_dpp)
                except:
                    continue
                    
                found_no_fp = None
                
                if target_date_str in lookup_data:
                    dpp_dict = lookup_data[target_date_str]
                    
                    if target_dpp_float in dpp_dict and len(dpp_dict[target_dpp_float]) > 0:
                        found_no_fp = dpp_dict[target_dpp_float].pop(0)
                        if not dpp_dict[target_dpp_float]:
                            del dpp_dict[target_dpp_float]
                    
                    elif tolerance > 0:
                        best_match_dpp = None
                        min_diff = float('inf')
                        
                        for source_dpp in dpp_dict.keys():
                            diff = abs(source_dpp - target_dpp_float)
                            if diff <= tolerance and len(dpp_dict[source_dpp]) > 0:
                                if diff < min_diff:
                                    min_diff = diff
                                    best_match_dpp = source_dpp
                        
                        if best_match_dpp is not None:
                            found_no_fp = dpp_dict[best_match_dpp].pop(0)
                            if not dpp_dict[best_match_dpp]:
                                del dpp_dict[best_match_dpp]
                                
                if found_no_fp:
                    cell_no_fp.value = found_no_fp
                    match_count += 1

    print(f"--> Selesai memproses {sheet.title}. Ditemukan {match_count} kecocokan.")

print("--> Menyimpan file...")
wb.save(target_filename)
print("--> Selesai. Data berhasil disimpan dengan metode ticketing + fallback tolerance.")