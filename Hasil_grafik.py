import pandas as pd
import matplotlib.pyplot as plt


#import  file excel
file_masuk = "data_wisuda_modif.xlsx"
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
jmlh_wisudawan = data.groupby('Program Studi')['NIM'].size().reset_index(name='Jumlah Wisudawan')

# Mencetak data wisudawan
print("=============================== DATA WISUDA ================================")
print(data)
print("=============================== DATA PRODI BANYAKNNYA JUMLAH YANG WISUDA ================================")
print(jmlh_wisudawan)


file_keluar="Rekap_hasil_wisuda_grafik.xlsx"
data.to_excel(file_keluar, index=False)

# ============  Diagram Batang  =============
plt.figure(figsize=(10, 5))
plt.bar(jmlh_wisudawan['Program Studi'], jmlh_wisudawan['Jumlah Wisudawan'], color='pink')
plt.title('Jumlah Wisudawan setiap Program Studi')
plt.xlabel('Program Studi')
plt.ylabel('Jumlah Wisudawan')
plt.xticks(rotation=30, ha='right')
plt.tight_layout()
plt.show()
plt.close()

# ===== Pie Chart Predikat =====
predikat_count = data['Predikat'].value_counts()
plt.figure(figsize=(9,18))
plt.pie(predikat_count, labels=predikat_count.index, autopct='%1.1f%%',
        startangle=90, colors=['orange','skyblue','violet','pink'])
plt.title(' Predikat Kelulusan')
plt.show()
