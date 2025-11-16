from typing import Literal
from pathlib import Path
import pandas as pd


class PivotTableManager:

    def __init__(self, file_path: str | None = None) -> None:
        # Sciezka do pliku Excel

        if file_path is None:
            self.file_path = None
            self.source_data: dict[str, pd.DataFrame] = {}
        else:
            self.file_path: Path = Path(file_path)
            self.source_data: dict[str, pd.DataFrame] = self._read_source_data()
        # Przechowuje wszystkie arkusze z pliku Excel jako DataFrames
        # sheet_name ustawiony na None to znaczy:
        # "wczytaj wszystkie arkusze z pliku Excel" jako słownik nazwa_arkusza -> DataFrame.
        # Gdybys ustawil sheet_name na "Sprzedaz" wczytalby tylko arkusz "Sprzedaz" jako DataFrame.

        # Przechowuje utworzone tabele przestawne (pivot tables)
        self.pivots: dict[str, pd.DataFrame] = {}

    def _read_source_data(self,file_path: str | None = None) -> dict[str, pd.DataFrame]:
        if file_path: self.file_path = Path(file_path)
        if self.file_path is None: raise ValueError("No file is specified to read data from")
        return pd.read_excel(self.file_path, sheet_name=None)

    def create_pivot(
            self,
            name: str,
            source_sheet_name: str,
            index: list[str],
            columns: list[str] | None = None,
            values: list[str] | None = None,
            agr_func: Literal['sum', 'mean', 'count', 'median', 'min', 'max'] = 'sum',
            source_file_path: Path | None = None
    ) -> None:
        """
        Tworzy nową tabelę przestawną i zapisuje ją pod unikalną nazwą.

        :param name: Nazwa tabeli przestawnej
        :param source_sheet_name: Nazwa arkusza źródłowego
        :param index: Kolumny indeksujące
        :param columns: Kolumny kolumnujące (opcjonalnie)
        :param values: Kolumny wartości (opcjonalnie)
        :param agr_func: Funkcja agregująca
        :param source_file_path: Zdefiniować nowy source file
        """
        if source_file_path:
            self.file_path = Path(source_file_path)
            self.source_data: dict[str, pd.DataFrame] = self._read_source_data()

        if name in self.pivots:
            raise ValueError(f'Pivot table "{name}" already exists')

        self._create(name, source_sheet_name, index, columns, values, agr_func)

    def get_pivot(self, name: str) -> pd.DataFrame | None:
        return self.pivots.get(name)

    def update_pivot(
            self,
            name: str,
            sheet_name: str,
            index: list[str],
            columns: list[str] | None = None,
            values: list[str] | None = None,
            aggfunc: Literal['sum', 'mean', 'count', 'median', 'min', 'max'] = 'sum'
    ) -> None:
        if name not in self.pivots:
            raise ValueError(f'Pivot table "{name}" not found')

        self._create(name, sheet_name, index, columns, values, aggfunc)

    def remove_pivot(self, name: str) -> None:
        if name in self.pivots:
            del self.pivots[name]

    def list_pivots(self) -> list[str]:
        return list(self.pivots.keys())

    def save_pivots_to_excel(self, output_path: str | None = None) -> None:
        target_path = Path(output_path) if output_path else self.file_path
        mode: Literal['a', 'w'] = 'a' if target_path.exists() else 'w'
        with pd.ExcelWriter(target_path, mode=mode) as writer:
            for name, df in self.pivots.items():
                df.to_excel(writer, sheet_name=name, if_sheet_exists='replace')

    # Zwracamy liste arkuszy z oryginalnego pliku
    def _list_source_sheets(self) -> list[str]:
        return list(self.source_data.keys())

    def _create(
            self,
            name: str,
            sheet_name: str,
            index: list[str],
            columns: list[str] | None = None,
            values: list[str] | None = None,
            aggfunc: Literal['sum', 'mean', 'count', 'median', 'min', 'max'] = 'sum'
    ) -> None:
        df = self.source_data.get(sheet_name)
        if df is None:
            raise ValueError(f'Sheet "{sheet_name}" not found')

        pivot = pd.pivot_table(
            df,
            # Lista kolumn, które staną się rzędami (wierszami) w tabeli przestawnej
            index=index,
            # (Opcjonalnie) kolumny, które staną się kolumnami w tabeli przestawnej
            columns=columns,
            # (Opcjonalnie) kolumny, na których wykona się agregacja (np. suma, średnia)
            values=values,
            # Funkcja agregująca: "sum", "mean", "count", "max", "min"
            aggfunc=aggfunc,
            # Dodaje wiersz i kolumnę z sumami (lub inną agregacją)
            margins=True,
            # Nazwa wiersza/kolumny z sumami (Total zamiast domyślnego All)
            margins_name='Total',
            # Usuwa wiersze / kolumny, gdzie wszystkie wartości są NaN (jeśli powstaną podczas pivotu)
            dropna=True
        )
        self.pivots[name] = pivot
