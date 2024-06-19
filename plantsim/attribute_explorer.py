"""
Copyright (c) 2021 Tilo van Ekeris / TMDT, University of Wuppertal
Distributed under the MIT license, see the accompanying
file LICENSE or https://opensource.org/licenses/MIT
"""

from enum import Enum

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

    @property
    def object_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.ObjectTable")

    @property
    def attribute_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.AttributeTable")

    @property
    def query_table(self) -> PandasTable:
        return PandasTable(self.plantsim, f"{self.object_name}.QueryTable")
