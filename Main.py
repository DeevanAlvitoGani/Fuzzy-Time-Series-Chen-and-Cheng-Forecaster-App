import pandas as pd
import numpy as np
import openpyxl
from prettytable import PrettyTable
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image



def process_file():
    try:
        # Open file dialog to select the CSV file
        filepath = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not filepath:
            return

        # Read the CSV file
        df = pd.read_csv(filepath)

        columns = df.columns.tolist()  # Get column names
        if len(columns) < 2:
            messagebox.showerror("Error", "The CSV file must have at least two columns.")
            return

        # Use the first two columns, regardless of their names
        a = df.iloc[:, 0].tolist()  # First column (Tanggal)
        b = df.iloc[:, 1].tolist()  # Second column (Nilai)
        
        d1 = max(b)
        d2 = min(b)
        
        #P adalah Universe Of Disclosure

        p = (d2,d1)

        #Menentukan jumlah data
        n = len(b)


        #Menentukan banyaknya kelas
        k = np.ceil(1 + 3.3*(np.log10(n)))



        #Menentukan Interval
        i = np.ceil((d1 - d2)/k)


        #membuat interval
        intervals = [d2 + int(i) * y for y in range(int(k) + 1)]



        #MEMBUAT INTERVAL DAN MEDIAN




        #Tabel1 (Tabel Interval dan Median)


        median = [(intervals[y] + intervals[y + 1]) / 2 for y in  range(len(intervals) - 1)]


        Table1 = PrettyTable(["Variabel", "Interval",  "Median"])

        rows = []
        for y in range(len(intervals) - 1):
            label = "A" + str(y + 1)  
            interval = f"{intervals[y]} - {intervals[y + 1]}"  
            median1 = f"{median[y]}"
            rows.append([label, interval, median1])  
                
        Table1.add_rows(rows) 


        print(Table1)




        #TABLE2 (Tabel Fuzzifikasi)


        Table2 = PrettyTable(["Tanggal", "Data Aktual", "Fuzzifikasi"])

        rows2 = []
        for y in range(len(b)):
            Tahun1 = f"{a[y]}"
            AQI1 = f"{b[y]}"
            
            fuzzy1 = 'undefined'
            
            for i in range(len(b) - 1):
                if intervals[i] <= b[y] <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1) 
                    break
                    
                
            
            rows2.append([Tahun1, AQI1, fuzzy1])    
        
        Table2.add_rows(rows2)


        # print(Table2)




        #TABLE3 (Fuzzifikasi dan Hubungan)


        Table3 = PrettyTable(["Tanggal", "Data Aktual", "Fuzzifikasi","Hubungan"])


        rows3 = []
        for y in range(len(b)):
            Tahun1 = f"{a[y]}"
            AQI1 = f"{b[y]}"
            
            fuzzy1 = 'undefined'
            
            for i in range(len(b) - 1):
                if intervals[i] <= b[y] <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1) 
                    break
            
            fuzzy2 = 'undefined'
            for i in range(len(b)):
                if y == 0 :
                    fuzzy2 = 'N/A'
                else:
                    if intervals[i] <= b[y-1] <= intervals[i + 1]:
                        fuzzy2 = "A" + str(i + 1)  
                        break
            
            if y == 0:
                hubunganfuzzy = 'N/A'      
                
            else: 
                hubunganfuzzy = f"{fuzzy2} - {fuzzy1}"
            
                
            
            
            rows3.append([Tahun1, AQI1, fuzzy1, hubunganfuzzy])   

        Table3.add_rows(rows3)


        # print(Table3)




        #Table5 (Fuzzy Logical Relationship Group)


        flrg = {}

        for y in range(1, len(b)):  # Dimulai dari 1 karena hari pertama tidak ada nilai
            fuzzy1 = 'undefined'
            fuzzy2 = 'undefined'
            
            #Fuzzifikasi Hari Ini
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y] <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1)
                    break
            
            #Fuzzifikasi Hari Sebelumnya
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y-1] <= intervals[i + 1]:
                    fuzzy2 = "A" + str(i + 1)
                    break
            
            #Membuat dictionary
            if fuzzy2 in flrg:
                if fuzzy1 not in flrg[fuzzy2]:
                    flrg[fuzzy2].append(fuzzy1)
            else:
                flrg[fuzzy2] = [fuzzy1]

        NilaiChen = {}
        for key, relations in flrg.items():
            median_fuzzy = [median[int(relation[1:])-1] for relation in relations]
            NilaiChen[key] = np.mean(median_fuzzy)


        Table5 = PrettyTable(["FLRG", "Relationship", "Nilai Prediksi"])

        for key, relations in flrg.items():
            relation_str = ', '.join(relations)
            prediction = round(NilaiChen[key], 1)  
            Table5.add_row([key, relation_str, prediction])


        print(Table5)
            





        #DEFUZZIFIKASI DAN MAPE
        #Tabel Chen 


        tablechen_rows = []
        mapechen = []  

        for y in range(len(b)):
            tanggal = a[y]  
            nilaiasli = b[y]  
            
            fuzzy1 = 'undefined'
            for i in range(len(intervals) - 1):
                if intervals[i] <= nilaiasli <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1)
                    break
                
            #Fuzzifikasi Hari Sebelumnya
            for i in range(len(intervals) - 1):
                if y == 0:
                    fuzzy2 = 'N/A'
                else: 
                    if intervals[i] <= b[y-1] <= intervals[i + 1]:
                        fuzzy2 = "A" + str(i + 1)
                        break

            if y == 0:
                PrediksiChen = 'N/A'  
            else:
                PrediksiChen = NilaiChen.get(fuzzy2, 'N/A')  
            
            if PrediksiChen != 'N/A':
                mape = np.ceil(abs((nilaiasli - PrediksiChen) / nilaiasli) * 100)
                mapechen.append(mape)  
            else:
                mape = 'N/A'
            
            
            tablechen_rows.append([tanggal, nilaiasli,  PrediksiChen, mape])


        TableResult = PrettyTable(["Tanggal", "Nilai", "Ramalan (Chen)", "MAPE"])

        for row in tablechen_rows:
            TableResult.add_row(row)

        avgmape = np.round(sum(mapechen) / len(mapechen), 2)



        # print("Fuzzy Chen:")
        # print(TableResult)
        # print(f"Average MAPE: {avgmape}")








            
            
        #Fuzzy CHENG

        #Matriks Bobot Relasi
        Matrixbobot = np.zeros((int(k), int(k)))


        for y in range(1, len(b)):
            fuzzy1 = fuzzy2 = 'undefined'
            
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y] <= intervals[i + 1]:
                    fuzzy1 = i
                    break
            
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y-1] <= intervals[i + 1]:
                    fuzzy2 = i
                    break
            
            # Isi Matrix nya
            if fuzzy1 != 'undefined' and fuzzy2 != 'undefined':
                Matrixbobot[fuzzy2][fuzzy1] += 1

        # print('Matrix Bobot Cheng:')        
        # print(Matrixbobot)



        # Matrix Normalisasi
        jumlahbaris = Matrixbobot.sum(axis=1, keepdims=True)
        matrixnormalisasi = np.divide(Matrixbobot, jumlahbaris, where=jumlahbaris!=0)
        matrixnormalisasi = np.round(matrixnormalisasi, 2) 


        # print("Normalisasi Matriks Bobot Cheng:")
        # print(matrixnormalisasi)



        #Kali dengan median

        result_matrix = np.zeros((int(k), int(k)))
        jumlahbaris = []

        for i in range(int(k)): 
            for j in range(int(k)):  
                result_matrix[i][j] = round(matrixnormalisasi[i][j] * median[j], 2) 
            jumlahbaris.append(round(np.sum(result_matrix[i]), 1)) 




        #Buat dictionary untuk cheng

        flrg2 = {}

        for y in range(1, len(b)):  # Dimulai dari 1 karena hari pertama tidak ada nilai
            fuzzy1 = 'undefined'
            fuzzy2 = 'undefined'
            
            #Fuzzifikasi Hari Ini
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y] <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1)
                    break
            
            #Fuzzifikasi Hari Sebelumnya
            for i in range(len(intervals) - 1):
                if intervals[i] <= b[y-1] <= intervals[i + 1]:
                    fuzzy2 = "A" + str(i + 1)
                    break
            
            #Membuat dictionary
            if fuzzy2 in flrg2:
                if fuzzy1 not in flrg[fuzzy2]:
                    flrg2[fuzzy2].append(fuzzy1)
            else:
                flrg2[fuzzy2] = [fuzzy1]

        NilaiCheng = {}
        for key, relations in flrg2.items():
            idx2 = int(key[1:]) - 1 
            NilaiCheng[key] = jumlahbaris[idx2]


        


        #Defuzzifikasi Cheng

        tablecheng_rows = []
        mapecheng = []  

        for y in range(len(b)):
            tanggal = a[y]  
            nilaiasli = b[y]  
            
            fuzzy1 = 'undefined'
            for i in range(len(intervals) - 1):
                if intervals[i] <= nilaiasli <= intervals[i + 1]:
                    fuzzy1 = "A" + str(i + 1)
                    break
                
            #Fuzzifikasi Hari Sebelumnya
            for i in range(len(intervals) - 1):
                if y == 0:
                    fuzzy2 = 'N/A'
                else: 
                    if intervals[i] <= b[y-1] <= intervals[i + 1]:
                        fuzzy2 = "A" + str(i + 1)
                        break

            if y == 0:
                prediksicheng = 'N/A'  
            else:
                prediksicheng = NilaiCheng.get(fuzzy2, 'N/A')  
            
            if prediksicheng != 'N/A':
                mape2 = np.ceil(abs((nilaiasli - prediksicheng) / nilaiasli) * 100)
                mapecheng.append(mape2)  
            else:
                mape2 = 'N/A'
            
            
            tablecheng_rows.append([tanggal, nilaiasli,  prediksicheng, mape2])


        TableResultCheng = PrettyTable(["Tanggal", "Nilai", "Ramalan (Cheng)", "MAPE"])

        for row in tablecheng_rows:
            TableResultCheng.add_row(row)

        avgmapecheng = np.round(sum(mapecheng) / len(mapecheng), 2)


        # print("\n")
        # print("Fuzzy Cheng:")
        # print(TableResultCheng)
        # print(f"Average MAPE: {avgmapecheng}")


        #Save File lebih dari 1
        Hasil = "Hasil Fuzzy Chen dan Cheng" 
        if not os.path.exists(Hasil):
            os.makedirs(Hasil)  

        namafile = "HasilFuzzy"
        fileextension = ".xlsx"
        filename = f"{namafile}{fileextension}"

        Hitungfile = 1
        while os.path.exists(os.path.join(Hasil, filename)):
            filename = f"{namafile}_{Hitungfile}{fileextension}"
            Hitungfile += 1

        output =  os.path.join(Hasil, filename)

        #Hasil ke Excel

        datachen = pd.DataFrame(tablechen_rows, columns=["Tanggal", "Nilai", "Ramalan (Chen)", "MAPE"])
        datacheng = pd.DataFrame(tablecheng_rows, columns=["Tanggal", "Nilai", "Ramalan (Cheng)", "MAPE"])
        
        

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            datachen.to_excel(writer, sheet_name='Hasil Chen', index=False)
            datacheng.to_excel(writer, sheet_name='Hasil Cheng', index=False)


        print("Results saved to HasilFuzzy.xlsx")


        # Example: display a success message after processing
        messagebox.showinfo("Success", "File processed successfully!")
    
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create a simple tkinter window for file selection
root = tk.Tk()
root.title("Fuzzy Chen and Cheng Method")
root.geometry("400x300")

# INstruksi
instructions = """Welcome to Fuzzy Chen and Cheng Method Processor!
1. Click the 'Select CSV File' button to load your file.
2. The application will process the file and generate results.
3. Make sure your CSV file contains 2 columns.
4. Make sure the csv is seperated by " , " and not " ; " 
5. Check your folder for the result"""


# Label
label = tk.Label(root, text=instructions, font=("Arial", 9), wraplength=350, justify="left")
label.pack(pady=10)

# Tombol
process_button = tk.Button(root, text="Select CSV File", command=process_file, height=2, width=20)
process_button.pack(pady=50)

# Run tkinter
root.mainloop()






