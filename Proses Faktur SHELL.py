import os
import glob
import shutil
import sys
import subprocess
import random

def print_msg(msg):
    print(f"--> {msg}")

dapur_folder = "Dapur"
required_files = [
    "1_AmbilLampiranGmail.py",
	"2_EkstrakPdfShell.py",
	"3_HelperDeleteDuplicate.py",
	"4_XlookupData.py",
	"5_KnifeToOperationFinalData.py",
	"6_CombineDeleteDuplicateAndSort.py",
	"7_HelperDeleteDuplicateFinalData.py",
	"8_XlookupData.py",
	"9_CopyDataToTemplate.py",
	"10_HelperDeleteTemplateData.py",
	"11_HelperMergedSUM.py",
	"A_JustXlookupData.py",
	"credentials.json",
	"CTX.xlsx",
	"gmail.conf",
	"gudang.conf",
	"lookup.conf",
	"__init__.py",
	"TEMPLATE.xlsx",
	"token.json"
]

if not os.path.exists(dapur_folder):
    print_msg(f"Folder {dapur_folder} tidak ditemukan.")
    input("--> Tekan enter untuk keluar")
    sys.exit()

missing_files = []
for f in required_files:
    if not os.path.exists(os.path.join(dapur_folder, f)):
        missing_files.append(f)

if missing_files:
    print_msg(f"File berikut tidak ditemukan di dalam folder {dapur_folder}: {', '.join(missing_files)}")
    input("--> Tekan enter untuk keluar")
    sys.exit()

root_data_export = glob.glob("data_export*.xlsx")
root_laporan_shell = glob.glob("Laporan SHELL*.xlsx")

if root_data_export:
    shutil.copy(root_data_export[0], os.path.join(dapur_folder, os.path.basename(root_data_export[0])))
    print_msg("File data_export ditemukan dan disalin ke Dapur.")

if root_laporan_shell:
    shutil.move(root_laporan_shell[0], os.path.join(dapur_folder, os.path.basename(root_laporan_shell[0])))
    print_msg("File Laporan SHELL ditemukan dan dipindahkan ke Dapur.")

if not root_data_export and not root_laporan_shell:
    print_msg("Data Laporan Shell dan data_export tersebut tidak ada, akan di lanjutkan dengan metode template. Tekan enter untuk melanjutkan")
    input()

dapur_data_export = glob.glob(os.path.join(dapur_folder, "data_export*.xlsx"))

if not dapur_data_export:
    rand_id = random.randint(10000, 99999)
    ctx_path = os.path.join(dapur_folder, "CTX.xlsx")
    new_export_name = os.path.join(dapur_folder, f"data_export_{rand_id}.xlsx")
    shutil.copy(ctx_path, new_export_name)
    print_msg(f"File data_export tidak ditemukan. Membuat dummy dari CTX: {os.path.basename(new_export_name)}")

print_msg("Pilih Menu Proses:")
print_msg("1. Ambil data dari Gmail dan proses")
print_msg("2. Hanya proses pengecekan nomor Faktur Pajak")
pilihan = input("--> Masukkan pilihan (1/2): ")

if pilihan == "1":
    scripts = [
        "1_AmbilLampiranGmail.py",
		"2_EkstrakPdfShell.py",
		"3_HelperDeleteDuplicate.py",
		"4_XlookupData.py",
		"5_KnifeToOperationFinalData.py",
		"6_CombineDeleteDuplicateAndSort.py",
		"7_HelperDeleteDuplicateFinalData.py",
		"8_XlookupData.py",
		"9_CopyDataToTemplate.py",
		"10_HelperDeleteTemplateData.py",
		"11_HelperMergedSUM.py"
    ]
    
    for script in scripts:
        print_msg(f"Menjalankan {script}...")
        try:
            subprocess.run([sys.executable, script], cwd=dapur_folder, check=True)
        except subprocess.CalledProcessError:
            print_msg(f"Gagal menjalankan {script}.")
            input("--> Tekan enter untuk keluar")
            sys.exit()
            
    result_files = glob.glob(os.path.join(dapur_folder, "TEMPLATE_temp.xlsx")) + glob.glob(os.path.join(dapur_folder, "Laporan SHELL*.xlsx"))
    
    if result_files:
        latest_file = max(result_files, key=os.path.getctime)
        dest_name = "Laporan SHELL BARU.xlsx"
        shutil.copy(latest_file, dest_name)
        print_msg(f"File hasil disalin menjadi: {dest_name}")
    else:
        print_msg("File hasil tidak ditemukan setelah proses.")

elif pilihan == "2":
    check_export = glob.glob(os.path.join(dapur_folder, "data_export*.xlsx"))
    check_laporan = glob.glob(os.path.join(dapur_folder, "Laporan SHELL*.xlsx"))
    
    if check_export and check_laporan:
        print_msg("Menjalankan A_JustXlookupData.py...")
        try:
            subprocess.run([sys.executable, "A_JustXlookupData.py"], cwd=dapur_folder, check=True)
            
            check_laporan_after = glob.glob(os.path.join(dapur_folder, "Laporan SHELL*.xlsx"))
            if check_laporan_after:
                src_laporan = check_laporan_after[0]
                shutil.move(src_laporan, os.path.basename(src_laporan))
                print_msg("File Laporan SHELL berhasil diproses dan dipindahkan.")
        except subprocess.CalledProcessError:
            print_msg("Gagal menjalankan A_JustXlookupData.py.")
            input("--> Tekan enter untuk keluar")
            sys.exit()
    else:
        print_msg("File data_export atau Laporan SHELL tidak lengkap di folder Dapur untuk proses pilihan 2.")
        input("--> Tekan enter untuk keluar")
        sys.exit()

else:
    print_msg("Pilihan tidak valid.")
    sys.exit()

files_to_clean = glob.glob(os.path.join(dapur_folder, "*temp.xlsx")) + \
                 glob.glob(os.path.join(dapur_folder, "Laporan SHELL*.xlsx")) + \
                 glob.glob(os.path.join(dapur_folder, "data_export*.xlsx")) + \
                 glob.glob(os.path.join(dapur_folder, "*.pdf")) + \
                 glob.glob(os.path.join(dapur_folder, "*.PDF"))

if files_to_clean:
    print_msg("Membersihkan file sementara di folder Dapur...")
    for f in files_to_clean:
        try:
            os.remove(f)
        except:
            pass

print_msg("Selesai, tekan enter untuk keluar")
input()
