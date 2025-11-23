# 📦 invoice_group_excel

Biblioteka Python służąca do automatycznej agregacji, filtrowania i generowania uporządkowanych arkuszy Excel na podstawie jednego pliku źródłowego. Narzędzie tworzy gotowe raporty kosztowe, grupuje dane, dodaje sumy, formatuje i eksportuje całość do nowego pliku XLSX.

## 🚀 Najważniejsze funkcje

- Automatyczne wczytywanie danych z pliku Excel.
- Grupowanie danych po wybranych kolumnach.
- Tworzenie wielu arkuszy na podstawie warunków filtrowania.
- Automatyczne wykrywanie kolumn liczbowych i dodawanie sum na końcu.
- Brak zależności od stałych nazw kolumn – biblioteka dostosowuje się do struktury danych.
- Sortowanie wyników według wybranych kolumn.
- Zapisywanie i stylowanie wygenerowanych arkuszy.

## 📁 Jak to działa (opis ogólny)

1. Użytkownik wskazuje plik wejściowy XLSX.
2. Biblioteka analizuje jego strukturę.
3. Na podstawie poleceń:
   - grupuje dane,
   - filtruje,
   - tworzy nowe arkusze,
   - sortuje,
   - dodaje sumy.
4. Wynik zapisywany jest jako nowy plik Excel.

## 📥 Instalacja

```
pip install invoice_group_excel
```

## 🧰 Przykładowe użycie

```python
from invoice_group_excel import InvoiceAggregator, CFG

def main() -> None:
    agregator = InvoiceAggregator(CFG.input_data_file, ['Tabela1'])
    agregator.add_sheets_grouped_by_key(key_to_group=['Task'], sort_by='Type')
    agregator.add_sheet_with_conditions(
        'Tabela1',
        output_sheet_name='illia_',
        condition_fn=lambda row: row["Type"] == 'illia',
        sort_by='Task'
    )
    agregator.update_input_data()
    agregator.save_and_style(CFG.output_data_file)
```

## 🔧 Konfiguracja

```python
class Config:
    data_dir: str = 'data'
    input_data_file: str = data_dir + '/costs_list_data.xlsx'
    output_data_file: str = data_dir + '/output_costs_list_data.xlsx'

    date_columns = ["B"]
    date_format = "MM-DD-YYYY"

    summary_col_description = ("A", "SUMMARY")

CFG: Config = Config()
```

## 🧪 Testy

```
pytest
```

## 🤝 Wkład

Pull requesty mile widziane.

## 📄 Licencja

MIT
