import tempfile
from pathlib import Path

import pandas as pd


class PandasTable:
    def __init__(self, plantsim, object_name: str):
        """
        Pandas Table mapping for PlantSim Tables (e.g., DataTable, ExplorerTable)
        - stores table in a .txt file in a temporary directory
        - returns that table as a pandas Dataframe
        :param plantsim: Plantsim instance (with loaded model) that is queried
        :param table_name: The object name within Plantsim relative to the current path context
        """
        self.plantsim = plantsim
        self._table_name = object_name

        self.update()

    def __repr__(self):
        return repr(self.table)

    def update(self) -> pd.DataFrame:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file_path = Path(temp_dir) / f"{self._table_name}.txt"
            self.plantsim.execute_simtalk(f'{self._table_name}.writeFile("{temp_file_path}")')
            self._table = pd.read_csv(temp_file_path, delimiter="\t")
        return self._table

    @property
    def table(self):
        return self._table
