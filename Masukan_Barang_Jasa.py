import os
import shutil
import subprocess
import sys

def main():
    root_dir = os.getcwd()
    dapur_dir = os.path.join(root_dir, "Dapur")

    required_root_files = [
        "Accuratem.xls",
        "Coretaxm.xlsx"
    ]

    required_dapur_files = [
        "__init__.py",
        "1_AccCleaner&PshBrgJs.py",
        "2_CtxPshBrgJs.py",
        "3_AnalyticsBrgAccCtx.py",
        "4_AnalyticsJsAccCtx.py",
        "5_MergeHasil.py",
        "hbrg.txt",
        "hjv.txt"
    ]

    scripts_to_run = [
        "1_AccCleaner&PshBrgJs.py",
        "2_CtxPshBrgJs.py",
        "3_AnalyticsBrgAccCtx.py",
        "4_AnalyticsJsAccCtx.py",
        "5_MergeHasil.py"
    ]

    missing_root = [f for f in required_root_files if not os.path.exists(os.path.join(root_dir, f))]
    
    if missing_root:
        print(f"--> File berikut tidak ditemukan di folder utama: {', '.join(missing_root)}")
        input("Tekan enter untuk keluar...")
        sys.exit()

    if not os.path.exists(dapur_dir):
        print("--> Folder 'Dapur' tidak ditemukan.")
        input("Tekan enter untuk keluar...")
        sys.exit()

    missing_dapur = [f for f in required_dapur_files if not os.path.exists(os.path.join(dapur_dir, f))]

    if missing_dapur:
        print(f"--> File berikut tidak ditemukan di dalam folder Dapur: {', '.join(missing_dapur)}")
        input("Tekan enter untuk keluar...")
        sys.exit()

    for filename in os.listdir(dapur_dir):
        file_path = os.path.join(dapur_dir, filename)
        if os.path.isfile(file_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls") or "temp" in filename:
                if filename not in required_dapur_files:
                    try:
                        os.remove(file_path)
                    except:
                        pass

    try:
        shutil.copy(os.path.join(root_dir, "Accuratem.xls"), os.path.join(dapur_dir, "Accuratem.xls"))
        shutil.copy(os.path.join(root_dir, "Coretaxm.xlsx"), os.path.join(dapur_dir, "Coretaxm.xlsx"))
    except Exception as e:
        print(f"--> Gagal menyalin file ke folder Dapur: {e}")
        input("Tekan enter untuk keluar...")
        sys.exit()

    os.chdir(dapur_dir)

    for script in scripts_to_run:
        print(f"--> Sedang menjalankan {script}...")
        try:
            subprocess.run([sys.executable, script], check=True)
        except subprocess.CalledProcessError:
            print(f"--> Gagal menjalankan script: {script}")
            os.chdir(root_dir)
            input("Tekan enter untuk keluar...")
            sys.exit()
        except Exception as e:
            print(f"--> Terjadi kesalahan: {e}")
            os.chdir(root_dir)
            input("Tekan enter untuk keluar...")
            sys.exit()

    expected_results = ["Hasil_Analisis_Barang_Dan_Jasa.xlsx"]
    missing_results = [f for f in expected_results if not os.path.exists(f)]

    if missing_results:
        print(f"--> File hasil tidak ditemukan: {', '.join(missing_results)}")
        os.chdir(root_dir)
        input("Tekan enter untuk keluar...")
        sys.exit()

    try:
        for res_file in expected_results:
            shutil.copy(res_file, os.path.join(root_dir, res_file))
    except Exception as e:
        print(f"--> Gagal menyalin hasil kembali ke folder utama: {e}")
        os.chdir(root_dir)
        input("Tekan enter untuk keluar...")
        sys.exit()

    files_to_remove = expected_results + ["Accuratem.xls", "Coretaxm.xlsx"]
    
    for f in os.listdir('.'):
        if f in files_to_remove or f.endswith("temp.xlsx"):
            try:
                os.remove(f)
            except:
                pass

    os.chdir(root_dir)
    print("--> Proses selesai. File hasil telah disalin ke folder utama.")

if __name__ == "__main__":
    main()