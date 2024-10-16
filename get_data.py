import requests
import sys
import pandas
from datetime import date

import os

def clear_console():
    # Check the platform
    if os.name == 'nt':  # For Windows
        os.system('cls')
    else:  # For macOS/Linux
        os.system('clear')

def find_outliers(column):
    print(column)
    col_data = column.dropna()
    print(col_data)
    q1 = col_data.quantile(0.25)
    q3 = col_data.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr 
    print(f"lower bound = {lower_bound}")
    print(f"upper bound = {upper_bound}")
    print(f"iqr = {iqr}")
    print(f"q1 = {q1}")
    print(f"q3 = {q3}")
    return (column < lower_bound) | (column > upper_bound)

clear_console()

carts = ["patient", "mla_ic", "mla_lr", "mla_mst", "mla_tst", "mla_psw", "mla_he", "mla_isw", "mtp_tst", "mtp_psw", "mtp_he", "mtp_isw", "ai_lr", "ai_mst", "ai_tst", "ca_ic", "ca_lr", "ca_mst", "ca_psw"] 
result = {}
df = pandas.DataFrame(columns=carts, index=[])

while True:
    username = input("Masukkan username: ")
    password = input("Masukkan password: ")
    login_url = "https://kinefeet.elgibor-solution.com/api/login"
    data = {
        "username": username,
        "password": password
    }

    response = requests.post(login_url, json=data)

    if response.status_code == 200:
        token = response.json().get("token")
        auth_token = "Bearer " + token
        headers = {"Authorization": auth_token}
        break

    print("Username atau password salah!")

clear_console()

while True:
    print("1) Cari nama pasien")
    print("2) Ambil data seluruh pasien")
    print("3) Ambil data pasien yang dibuat pada hari ini")
    print("4) Ambil data pasien yang dibuat pada bulan ini")
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
                    patient = data[(int(inp) - 1)]
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

#process dataframe

print("Mencari outlier...")
i = 1
for col in df.columns:
    if col != "patient":
        df[f'{col}_is_outlier'] = find_outliers(df[col])
    #i += 1
    #if i == 4:
    #    break

with pandas.ExcelWriter('data/output.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    # Access the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Define a format for the outliers (red background)
    format_outlier = workbook.add_format({'bg_color': '#FF6666'})  # Light red background

    # Apply conditional formatting for each numeric column
    for col_num, col in enumerate(df.columns):
        if '_is_outlier' not in col and col != 'patient':
            # Calculate row range for data (2nd row onwards, 1-based index)
            start_row = 2
            end_row = len(df) + 1  # 1-based, so add 1
            range_str = f'{chr(65+col_num)}{start_row}:{chr(65+col_num)}{end_row}'
            
            # Apply conditional formatting based on the 'is_outlier' column
            outlier_col = f'{col}_is_outlier'
            outlier_col_num = df.columns.get_loc(outlier_col) + 1  # 1-based index for Excel

            worksheet.conditional_format(range_str, {
                'type':     'formula',
                'criteria': f'=$${chr(65+outlier_col_num)}{start_row}=TRUE',  # Reference the outlier column
                'format':   format_outlier
            })

#output_name = input("Masukkan nama dari file yang akan dibuat untuk menampung data: ")
#output = "data/" + output_name + ".xlsx"
#df.to_excel(output, index = False)
#print(f"Data ada di file {output}.")
