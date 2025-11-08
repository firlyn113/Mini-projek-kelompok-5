# Mini-projek-kelompok-5

import pandas as pd

#import  file excel
file_masuk =  r"data_wisuda_modif.xlsx"
data = pd.read_excel(file_masuk)

# Menentukan Grade
def grade(i):
    if i >= 3.75 :
        return "A"
    elif i >= 3.50:
        return "B+"
    elif i >= 3.00:
        return "B"
    elif i >=2.50:
        return "C"
    else:
        return "D"

# Menentukan Predikat
def predikat(row):
    ipk = row['IPK']
    studi = row['Lama Studi (Semester)']

    if ipk >= 3.75 and studi <= 8 :
        return "Cumlaude"
    elif ipk >= 3.50 and studi <=9:
        return "Sangat Memuaskan"
    elif ipk >= 3.00 :
        return "Memuaskan"
    else:
        return "Cukup"
    
data['Grade']= data['IPK'].apply(grade)
data['Predikat']= data.apply(predikat, axis=1)
data = data[['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Lama Studi (Semester)','Grade', 'Predikat','Tahun Wisuda']]

# Mengelompokan progam studi
jmlh_wisudawan = data.groupby('Program Studi')['NIM'].count().reset_index(name='Jumlah Wisudawan')

# Mencetak data wisudawan
print("=============================== DATA WISUDA ================================")
print(data)
print("=============================== DATA PRODI BANYAKNNYA JUMLAH YANG WISUDA ================================")
print(jmlh_wisudawan)


file_keluar=r"Rekap_hasil_wisuda.xlsx"
data.to_excel(file_keluar, index=False)

print(f"\nFile hasil telah disimpan ke: {file_keluar}")



import pandas as pd

#import  file excel
file_masuk =  r"C:/Users/MUHAMMAD ZAKI F P/Documents/File Tugas/data_wisuda_modif.xlsx"
data = pd.read_excel(file_masuk)

# Menentukan Grade
def grade(i):
    if i >= 3.75 :
        return "A"
    elif i >= 3.50:
        return "B+"
    elif i >= 3.00:
        return "B"
    elif i >=2.50:
        return "C"
    else:
        return "D"

# Menentukan Predikat
def predikat(row):
    ipk = row['IPK']
    studi = row['Lama Studi (Semester)']

    if ipk >= 3.75 and studi <= 8 :
        return "Cumlaude"
    elif ipk >= 3.50 and studi <=9:
        return "Sangat Memuaskan"
    elif ipk >= 3.00 :
        return "Memuaskan"
    else:
        return "Cukup"
    
data['Grade']= data['IPK'].apply(grade)
data['Predikat']= data.apply(predikat, axis=1)
data = data[['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Lama Studi (Semester)','Grade', 'Predikat','Tahun Wisuda']]

# Mengelompokan progam studi
jmlh_wisudawan = data.groupby('Program Studi')['NIM'].count().reset_index(name='Jumlah Wisudawan')

# Mencetak data wisudawan
print("=============================== DATA WISUDA ================================")
print(data)
print("=============================== DATA PRODI BANYAKNNYA JUMLAH YANG WISUDA ================================")
print(jmlh_wisudawan)


file_keluar=r"C:/Users/MUHAMMAD ZAKI F P/Documents/File Tugas/Rekap_hasil_wisuda.xlsx"
data.to_excel(file_keluar, index=False)

print(f"\nFile hasil telah disimpan ke: {file_keluar}")

