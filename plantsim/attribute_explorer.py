"""
Copyright (c) 2021 Tilo van Ekeris / TMDT, University of Wuppertal
Distributed under the MIT license, see the accompanying
file LICENSE or https://opensource.org/licenses/MIT
"""

from enum import Enum
from pathlib import Path

from plantsim.pandas_table import PandasTable


class AttributeExplorerMode(Enum):
    WATCH = "Watch"
    EDIT = "Edit"
    READ = "Read"


class AttributeExplorer:
    def __init__(self, plantsim, object_name: str) -> None:
        self.plantsim = plantsim
        self.object_name = object_name
        self._mode: AttributeExplorerMode = AttributeExplorerMode(self.plantsim.get_value(f"{self.object_name}.Mode"))

    @property
    def mode(self) -> AttributeExplorerMode:
        self._mode: AttributeExplorerMode = AttributeExplorerMode(self.plantsim.get_value(f"{self.object_name}.Mode"))
        return self._mode

    @mode.setter
    def set_mode(self, mode: AttributeExplorerMode) -> None:
        self.plantsim.set_value(f"{self.object_name}.Mode", mode.value)
        self._mode = mode

    @property
    def explorer_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.ExplorerTable")

    def import_explorer_table(self, path: str | Path, sheet: str | None = None) -> None:
        if self.mode != AttributeExplorerMode.EDIT:
            raise Exception("Attribute Explorer must be in edit mode.")

        path = Path(path).absolute()

        suffix = path.suffix.lower()
        if suffix == ".xlsx" or suffix == ".xls":
            read_function = f'readExcelFile("{path}", "{sheet}")' if sheet else f'readExcelFile("{path}")'
        elif suffix == ".xml":
            read_function = f'readXMLFile("{path}")'
        else:
            read_function = f'readFile("{path}")'

        command = f'var t: table; t.create; t.ColumnIndex := True; t.RowIndex := True; t.{read_function}; {self.object_name}.ExplorerTable := t'
        self.plantsim.execute_simtalk(command, from_path_context=False)

    def export_explorer_table(self, path: str | Path, sheet: str | None = None) -> None:
        path = Path(path).absolute()

        suffix = path.suffix.lower()
        if suffix == ".xlsx" or suffix == ".xls":
            write_function = f'writeExcelFile("{path}", "{sheet}")' if sheet else f'writeExcelFile("{path}")'
        elif suffix == ".xml":
            write_function = f'writeXMLFile("{path}")'
        else:
            write_function = f'writeFile("{path}")'

        command = f'{self.object_name}.ExplorerTable.{write_function}'
        self.plantsim.execute_simtalk(command, from_path_context=False)

    @property
    def object_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.ObjectTable")

    @property
    def attribute_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.AttributeTable")

    @property
    def query_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.QueryTable")
