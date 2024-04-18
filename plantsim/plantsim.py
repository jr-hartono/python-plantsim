"""
Copyright (c) 2021 Tilo van Ekeris / TMDT, University of Wuppertal
Distributed under the MIT license, see the accompanying
file LICENSE or https://opensource.org/licenses/MIT
"""

import shutil
from enum import Enum
from pathlib import Path
from typing import Any

import win32com
import win32com.client as win32

from .attribute_explorer import AttributeExplorer
from .error_code import ErrorCode


class LicenseType(Enum):
    PROFESSIONAL = "Professional"
    STANDARD = "Standard"
    APPLICATION = "Application"
    RUNTIME = "Runtime"
    RESEARCH = "Research"
    EDUCATIONAL = "Educational"
    STUDENT = "Student"


class PlantSim:
    def __init__(
        self,
        *,
        license_type: LicenseType,
        version: str | None = None,
        visible: bool = True,
        trust_models: bool = False,
    ) -> None:
        dispatch_string: str = "Tecnomatix.PlantSimulation.RemoteControl"
        if version:
            dispatch_string += f".{version}"

        try:
            self.plantsim = win32.gencache.EnsureDispatch(dispatch_string)
        except AttributeError as e:
            if e.name == "CLSIDToClassMap":
                shutil.rmtree(win32com.__gen_path__)  # cleanup win32com cache
                self.plantsim = win32.gencache.EnsureDispatch(dispatch_string)

        self.plantsim.SetVisible(visible)
        self.plantsim.SetTrustModels(trust_models)

        self.license_type: str = license_type.value
        try:
            self.plantsim.SetLicenseType(self.license_type)
        except BaseException as e:
            if ErrorCode.extract(e.args) == -2147221503:
                raise Exception(
                    f"The license type {self.license_type} does not seem to exist. Make sure it is a valid Plant Simulation license type."
                ) from e

        self.path_context: str = ""
        self.event_controller: str = ""

    def load_model(self, filepath: Path | str) -> None:
        try:
            self.plantsim.LoadModel(str(filepath))
        except BaseException as e:
            if ErrorCode.extract(e.args) == -2147221503:
                raise Exception(
                    f'The license server or the selected license type "{self.license_type}" is not available.\n'
                    "Make sure that the license server is up and running and you can connect to it (VPN etc.).\n"
                    f'Make sure that a valid license of type "{self.license_type}" is available in the license server.'
                ) from e

    def set_path_context(self, path_context: str) -> None:
        self.path_context = path_context
        self.plantsim.SetPathContext(self.path_context)

    def set_event_controller(self, path: str | None = None) -> None:
        self.event_controller = path if path is not None else f"{self.path_context}.EventController"

    def reset_simulation(self) -> None:
        if not self.event_controller:
            raise Exception("You need to set an event controller first!")
        self.plantsim.ResetSimulation(self.event_controller)

    def start_simulation(self) -> None:
        if not self.event_controller:
            raise Exception("You need to set an event controller first!")
        self.plantsim.StartSimulation(self.event_controller)

    def is_simulation_running(self) -> bool:
        return self.plantsim.IsSimulationRunning()

    def stop_simulation(self) -> None:
        self.plantsim.StopSimulation()

    def get_object(self, object_name: str) -> AttributeExplorer | Any:
        # "Smart" getter that has some limited ability to decide which kind of object to return"""
        internal_class_name = self.get_value(f"{object_name}.internalClassName")

        if internal_class_name == "AttributeExplorer":
            # Attribute explorer that dynamically fills a table
            return AttributeExplorer(self, object_name)
        if internal_class_name == "NwData":
            # Normal string
            return self.get_value(object_name)
        # Fallback: Return raw value of object
        return self.get_value(object_name)

    def get_value(self, object_name: str) -> Any:
        return self.plantsim.GetValue(object_name)

    def set_value(self, object_name: str, value: Any) -> None:
        self.plantsim.SetValue(object_name, value)

    def execute_simtalk(
        self,
        command_string: str,
        *parameter: tuple[Any, ...],
        from_path_context: bool = True,
    ) -> None:
        """
        Execute a SimTalk command according to COM documentation:
        PlantSim.ExecuteSimTalk("->real; return 3.14159")
        PlantSim.ExecuteSimTalk("param r:real->real; return r*r", 3.14159)

        :param command_string: Command to be executed
        :param parameter: (optional); parameter, if command contains a parameter to be set
        :param from_path_context: if true, command should be formulated from the path context
        """
        command_string = f"{self.path_context}.{command_string}" if from_path_context else f".{command_string}"

        if parameter:
            self.plantsim.ExecuteSimTalk(command_string, parameter)
        else:
            self.plantsim.ExecuteSimTalk(command_string)

    def quit(self) -> None:
        self.plantsim.Quit()
