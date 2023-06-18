import pandas as pd
import xlsxwriter
import main


data = pd.read_csv('results-survey581489 (1).csv')

data.drop(data[pd.isna(data['M1. Płeć'])].index, inplace = True)
data.drop(data[data['BM01. Wymień znane Ci mniejszości narodowe i etniczne.'] == 'test'].index, inplace = True)
data.drop(columns='id. ID odpowiedzi', inplace=True)


#oddzielanie danych o ankietowanych
metryczka = data.iloc[:, -6:].copy()
metryczka = metryczka.iloc[:, [0, 2, 3, 4, 5, 1]]

excel = xlsxwriter.Workbook('wyniki.xlsx')
sheet = excel.add_worksheet()


formaty = main.tworzenie_formatow(excel) #Tablica, która zawiera wszystkie potrzebne formatowania komórek

row = 0
i = iter(range(1, len(data.columns[:37])+1))

row = main.tytul(sheet, '1. Badanie mniejszości', row, 0, formaty[0])


row = main.jednokrotne(sheet, data.iloc[:, 0], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 1], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 2], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 3], metryczka, row, next(i), formaty)
row = main.wielokrotne(sheet, data.iloc[:, 4:13], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 13], metryczka, row, next(i), formaty)
row = main.wielokrotne(sheet, data.iloc[:, 14:20], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 21], metryczka, row, next(i), formaty)
row = main.kilka_jednokrotne(sheet, data.iloc[:, 22:24], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 20], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 24], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 25], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 26], metryczka, row, next(i), formaty)
row = main.kilka_jednokrotne(sheet, data.iloc[:, 27:29], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 29], metryczka, row, next(i), formaty)
row = main.kilka_jednokrotne(sheet, data.iloc[:, 30:32], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 32], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 33], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 34], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 35], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, data.iloc[:, 36], metryczka, row, next(i), formaty)

i = iter(range(1, len(metryczka.columns)+1))

row = main.tytul(sheet, '2. Metryczka', row, 0, formaty[0])
row = main.jednokrotne(sheet, metryczka.iloc[:, 0], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, metryczka.iloc[:, 1], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, metryczka.iloc[:, 2], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, metryczka.iloc[:, 3], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, metryczka.iloc[:, 4], metryczka, row, next(i), formaty)
row = main.jednokrotne(sheet, metryczka.iloc[:, 5], metryczka, row, next(i), formaty)

excel.close()