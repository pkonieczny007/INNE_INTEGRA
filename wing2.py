import pandas as pd
import xlwings as xw
import os

# Określ nazwę pliku Excel
nazwa_pliku = 'wykaz.xlsx'
nazwa_pliku_org = 'wykaz_org.xlsx'

# Używając pandas, wczytaj plik Excel
sciezka_pliku = os.path.join(os.getcwd(), nazwa_pliku)
df = pd.read_excel(sciezka_pliku)

# Zapisz kopię oryginalnego DataFrame jako wykaz_org.xlsx przed modyfikacją
df.to_excel(nazwa_pliku_org, index=False)

# Zamień wartości w kolumnach 'kategoria' i 'Bezeichnung' na ciągi znaków
df['kategoria'] = df['kategoria'].astype(str)
df['Bezeichnung'] = df['Bezeichnung'].astype(str)

# Dodaj nową kolumnę 'nazwa', łącząc kolumny 'kategoria' i 'Bezeichnung' z przecinkiem, tylko tam gdzie 'TYP' == 'INNE'
df['nazwa'] = df.apply(lambda x: x['kategoria'] + ', ' + x['Bezeichnung'] if x['TYP'] == 'INNE' else '', axis=1)

# Zapisz zmodyfikowany DataFrame z powrotem do pliku Excel za pomocą xlwings
with xw.App(visible=False) as app:
    wb = xw.Book(sciezka_pliku)
    sheet = wb.sheets[0]
    sheet.range('A1').options(index=False, headers=True).value = df
    wb.save()
    wb.close()
