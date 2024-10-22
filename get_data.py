import requests
import sys
import pandas
from datetime import date
from openpyxl import load_workbook, Workbook

import os

def clear_console():
    # Check the platform
    if os.name == 'nt':  # For Windows
        os.system('cls')
    else:  # For macOS/Linux
        os.system('clear')

def find_outliers(column):
    col_data = column.dropna()
    q1 = col_data.quantile(0.25)
    q3 = col_data.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr 
    return (column < lower_bound) | (column > upper_bound)

def excel_column_name(n):
    """Converts a zero-based column index to an Excel column name (e.g., 0 -> 'A', 26 -> 'AA')."""
    result = ''
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result
# Export to Excel with conditional formatting

clear_console()

carts = ["patient", "mla_ic", "mla_lr", "mla_mst", "mla_tst", "mla_psw", "mla_he", "mla_isw", "mtp_tst", "mtp_psw", "mtp_he", "mtp_isw", "ai_lr", "ai_mst", "ai_tst", "ca_ic", "ca_lr", "ca_mst", "ca_psw"] 
result = {}
df = pandas.DataFrame(columns=carts, index=[])

try:
    creds = open("credentials.txt", "r")
    lst = []
    for line in creds:
        lst.append(line.strip())
    username = lst[0]
    password = lst[1]
    creds.close()
except:
    username = None
    password = None

while True:
    if not username:
        username = input("Masukkan username: ")
        password = input("Masukkan password: ")
        creds = open("credentials.txt", "w")
        creds.write(username + "\n")
        creds.write(password)
        creds.close()
    login_url = "https://kinefeet.elgibor-solution.com/api/login"
    data = {
        "username": username,
        "password": password
    }

    try:
        response = requests.post(login_url, json=data)
    except:
        print("Terdapat kendala dalam koneksi ke situs Kinefeet. Pastikan anda terhubung ke internet.")
        try_again = input("Coba lagi? (Y/n) ")
        if (not try_again) or (try_again.lower() == "y"):
            continue
        else:
            exit()

    if response.status_code == 200:
        token = response.json().get("token")
        auth_token = "Bearer " + token
        headers = {"Authorization": auth_token}
        break

    print("Username atau password salah!")
    open("credentials.txt", "w").close()
    username = None
    password = None

clear_console()

while True:
    print("1) Cari nama pasien")
    print("2) Ambil data seluruh pasien")
    print("3) Ambil data pasien yang dibuat pada hari ini")
    print("4) Ambil data pasien yang dibuat pada bulan ini")
    print("5) Ambil data pasien yang dibuat setelah tanggal 8 Oktober 2024")
    print("6) Ambil data pasien yang dibuat pada rentang waktu tertentu")
    option = input("Ketik nomor dari operasi yang ingin dilakukan: ")

    if option == "1":
        while True:
            name = input("Masukkan nama pasien (maximal 1 pasien): ")
            page = 1
            patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page) + "&patient_name=" + name

            response = requests.get(patient_search_url, headers=headers).json()
            data = response.get("data")
            total = response.get("total")
            if total > 1:
                inp = input(f"Ada {total} pasien yang memenuhi kriteria nama. Tampilkan semuanya? (Y/n) ")
                if (not inp) or (inp == "y"):
                    index = 1
                    i = 1
                    while index <= total:
                        patient = data[(i - 1)]
                        print(i)
                        print(f"{patient.get('patient_name')}")
                        print(f"Created at: {patient.get('created_at')[0:10]}")
                        print(f"Last update: {patient.get('updated_at')[0:10]}")
                        index += 1
                        i += 1
                        if index % 10 == 1:
                            inp = input("Apakah pasien yang sesuai ada di list? (Y/n) ")
                            if (not inp) or (inp == "y"):
                                break
                            page += 1
                            patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page) + "&patient_name=" + name
                            response = requests.get(patient_search_url, headers=headers).json()
                            data = response.get("data")
                            i = 1
                            print()
                    inp = input("Pilih nomor pasien yang sesuai: ")
                    try:
                        patient = data[(int(inp) - 1)]
                    except:
                        print("Input tidak valid, pastikan input sesuai dengan nomor yang tersedia")
                        continue
                else:
                    continue
            elif total == 1:
                patient = data[0]
            else:
                print(f"Tidak ada pasien yang memenuhi kriteria nama {name}.")
                continue

            patient_name = patient.get("patient_name")

            print()
            print(f"{patient_name}")
            print(f"Created at: {patient.get('created_at')[0:10]}")
            print(f"Last update: {patient.get('updated_at')[0:10]}")

            confirmation = input("Apakah pasien yang dipilih sudah sesuai? (Y/n) ")

            if (not confirmation) or (confirmation.lower() == "y"):
                result["patient"] = patient_name
                for cart in carts:
                    if cart == "patient":
                        continue
                    if patient.get(cart) != None:
                        result[cart] = float(patient.get(cart))
            else:
                continue

            df.loc[len(df)] = result 
            print(df)

            inp = input("Pasien selanjutnya? Apabila sudah selesai, ketik n. (Y/n) ")
            if inp.lower() == "n":
                break
        break

    elif option == "2":
        page = 1
        patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
        response = requests.get(patient_search_url, headers=headers).json()
        data = response.get("data")
        total = response.get("total")
        option = input(f"Terdapat {total} jumlah pasien. Ambil seluruh data pasien tersebut? (Y/n) ")
        
        if option.lower() == "n":
            continue

        print(f"Mengambil {total} data, mohon tunggu...")
        index = 1
        i = 1
        while index <= total:
            print(f"({index}/{total})")
            patient = data[(i - 1)]
            patient_name = patient.get("patient_name")
            result["patient"] = patient_name
            for cart in carts:
                if cart == "patient":
                    continue
                if patient.get(cart) != None:
                    result[cart] = float(patient.get(cart)) 
            df.loc[len(df)] = result

            index += 1
            i += 1
            if index % 10 == 1:
                page += 1
                patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
                response = requests.get(patient_search_url, headers=headers).json()
                data = response.get("data")
                i = 1
            #endpoint = "https://kinefeet.elgibor-solution.com/api/checkup/detail/"
        break

    elif option == "3":                     #Hari ini
        today = date.today()
        print(f"Mengambil data pasien yang dibuat pada tanggal {today}...")
        page = 1
        patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
        response = requests.get(patient_search_url, headers=headers).json()
        data = response.get("data")
        total = response.get("total")
        i = 1
        index = 0
        while True:
            patient = data[(i - 1)]
            if patient.get("updated_at")[0:10] != str(today):
                break
            patient_name = patient.get("patient_name")
            result["patient"] = patient_name
            for cart in carts:
                if cart == "patient":
                    continue
                if patient.get(cart) != None:
                    result[cart] = float(patient.get(cart)) 
            df.loc[len(df)] = result
            i += 1
            index += 1
            if index % 10 == 1:
                page += 1
                patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
                response = requests.get(patient_search_url, headers=headers).json()
                data = response.get("data")
                i = 1
        print(f"Berhasil mengambil {index} data pasien pada tanggal {today}")
        break

    elif option == "4":
        today = str(date.today())
        this_month = today[5:-3]
        print(f"Mengambil data pasien yang dibuat pada bulan {this_month}...")
        page = 1
        patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
        response = requests.get(patient_search_url, headers=headers).json()
        data = response.get("data")
        total = response.get("total")
        i = 1
        index = 0
        while True:
            patient = data[(i - 1)]
            if patient.get("updated_at")[5:7] != str(this_month):
                break
            patient_name = patient.get("patient_name")
            result["patient"] = patient_name
            for cart in carts:
                if cart == "patient":
                    continue
                if patient.get(cart) != None:
                    result[cart] = float(patient.get(cart)) 
            df.loc[len(df)] = result
            i += 1
            index += 1
            if index % 10 == 1:
                page += 1
                patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
                response = requests.get(patient_search_url, headers=headers).json()
                data = response.get("data")
                i = 1
        print(f"Berhasil mengambil {index} data pasien pada bulan {this_month}")
        break

    elif option == "5":
        #8 Oktober 2024
        print(f"Mengambil data pasien yang dibuat hingga tanggal 8 Oktober 2024...")
        page = 1
        patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
        response = requests.get(patient_search_url, headers=headers).json()
        data = response.get("data")
        total = response.get("total")
        i = 1
        index = 0
        while True:
            patient = data[(i - 1)]
            if patient.get("id") < 104:           #Pasien Joan ID 104
                break
            patient_name = patient.get("patient_name")
            result["patient"] = patient_name
            for cart in carts:
                if cart == "patient":
                    continue
                if patient.get(cart) != None:
                    result[cart] = float(patient.get(cart)) 
            df.loc[len(df)] = result
            i += 1
            index += 1
            if index % 10 == 1:
                page += 1
                patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
                response = requests.get(patient_search_url, headers=headers).json()
                data = response.get("data")
                i = 1
        print(f"Berhasil mengambil {index} data pasien hingga tanggal 8 Oktober 2024")
        break

    elif option == "6":                                                  #Ambil data hingga tanggal tertentu
        cutoff_date = input("Masukkan tanggal (format: dd-mm-yyyy. Contoh: 01-02-2024 = 1 Februari 2024): ")
        cday = cutoff_date[:2]
        cmonth = cutoff_date[3:5]
        cyear = cutoff_date[6:10]
        valid_date = f"{cyear}-{cmonth}-{cday}"

        page = 1
        patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
        print("Mengambil data. Mohon tunggu...")
        response = requests.get(patient_search_url, headers=headers).json()
        data = response.get("data")
        total = response.get("total")
        i = 1
        index = 0
        while True:
            patient = data[(i - 1)]
            if patient.get("updated_at")[:10] < valid_date:   
                break
            patient_name = patient.get("patient_name")
            result["patient"] = patient_name
            for cart in carts:
                if cart == "patient":
                    continue
                if patient.get(cart) != None:
                    result[cart] = float(patient.get(cart)) 
            df.loc[len(df)] = result
            i += 1
            index += 1
            if index % 10 == 1:
                page += 1
                patient_search_url = "https://kinefeet.elgibor-solution.com/api/checkup?page=" + str(page)
                response = requests.get(patient_search_url, headers=headers).json()
                data = response.get("data")
                i = 1
            
            print(f"Telah mengambil {index} data...")
        print(f"Berhasil mengambil {index} data pasien hingga tanggal {cutoff_date}")
        break

    else:
        print("Tolong masukkan angka yang valid.")

#process dataframe
inp = input("Cari outlier dari data? (Y/n) ")
output_name = input("Masukkan nama dari spreadsheet yang akan dibuat: ")
output = f"data/{output_name}.xlsx"

if not (inp.lower() == "n"):
    print("Mencari outlier...")
    i = 1
    for col in df.columns:
        if col != "patient":
            df[f'{col}_is_outlier'] = find_outliers(df[col])
        #i += 1
        #if i == 4:
        #    break

with pandas.ExcelWriter(output, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Define a format for the outliers
    format_outlier = workbook.add_format({'bg_color': '#FF6666'})

    # Apply conditional formatting for each numeric column
    for col_num, col in enumerate(df.columns):
        if '_is_outlier' not in col and col != 'patient':
            # Calculate the range for this column (skip the header row)
            col_letter = excel_column_name(col_num)   # Convert column number to Excel column letter (A, B, etc.)
            data_range = f'{col_letter}2:{col_letter}{len(df) + 1}'  # Start at row 2, ending at the last data row

            # Corresponding outlier columns
            outlier_col_letter = excel_column_name(df.columns.get_loc(f'{col}_is_outlier'))

            # Apply conditional formatting for the data range based on the outlier column
            worksheet.conditional_format(data_range, {
                'type': 'formula',
                'criteria': f'=${outlier_col_letter}2=TRUE',
                'format': format_outlier
            })
    for col_num, col in enumerate(df.columns):
        if '_is_outlier' in col:
            worksheet.set_column(col_num, col_num, None, None, {'hidden': True})

print(f"Data berada dalam file {output}")
