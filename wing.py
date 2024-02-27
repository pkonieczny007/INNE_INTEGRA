import pandas as pd
import xlwings as xw
import os

# Określ nazwę pliku Excel w katalogu, w którym znajduje się skrypt
nazwa_pliku = 'wykaz.xlsx'

# Używając pandas, wczytaj plik Excel
sciezka_pliku = os.path.join(os.getcwd(), nazwa_pliku)
df = pd.read_excel(sciezka_pliku)

# Sprawdź, czy kolumny 'kategoria' i 'Bezeichnung' istnieją w danych
if 'kategoria' in df.columns and 'Bezeichnung' in df.columns:
    # Dodaj nową kolumnę 'nazwa', łącząc dane z kolumn 'kategoria' i 'Bezeichnung' przecinkiem
    df['nazwa'] = df['kategoria'] + ', ' + df['Bezeichnung']
else:
    print("Błąd: Brak jednej z wymaganych kolumn ('kategoria' lub 'Bezeichnung').")

# Zapisz zmodyfikowany DataFrame z powrotem do pliku Excel
with xw.App(visible=False) as app:
    wb = xw.Book(sciezka_pliku)
    sheet = wb.sheets[0]
    sheet.range('A1').options(index=False).value = df
    wb.save()
    wb.close()
