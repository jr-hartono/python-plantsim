"""
Copyright (c) 2021 Tilo van Ekeris / TMDT, University of Wuppertal
Distributed under the MIT license, see the accompanying
file LICENSE or https://opensource.org/licenses/MIT
"""

import inspect
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
        model: Path | str | None = None,
        *,
        license_type: LicenseType | str | None = None,
        version: str | None = None,
        visible: bool = True,
        trust_models: bool = False,
        path_context: str = ".Models.Model",
        event_controller: str = ".Models.Model.EventController",
    ) -> None:
        dispatch_string: str = "Tecnomatix.PlantSimulation.RemoteControl"
        if version:
            dispatch_string += f".{version}"

        try:
            self._plantsim = win32.gencache.EnsureDispatch(dispatch_string)
        except AttributeError as e:
            if e.name == "CLSIDToClassMap":
                shutil.rmtree(win32com.__gen_path__)  # cleanup win32com cache
                self._plantsim = win32.gencache.EnsureDispatch(dispatch_string)

        self.visible = visible
        self.trust_models = trust_models
        if license_type is not None:
            self.license_type = license_type
        if model is not None:
            self.model = model

        self.path_context = path_context
        self.event_controller = event_controller

    @property
    def model(self) -> Path:
        return self._model

    @model.setter
    def model(self, model: Path | str) -> None:
        self._model = Path(model).absolute()
        try:
            self._plantsim.LoadModel(str(self._model))
        except BaseException as e:
            if ErrorCode.extract(e.args) == -2147221503:
                raise Exception(
                    f'The license server or the selected license type "{self.license_type}" is not available.\n'
                    "Make sure that the license server is up and running and you can connect to it (VPN etc.).\n"
                    f'Make sure that a valid license of type "{self.license_type}" is available in the license server.'
                ) from e

    @property
    def license_type(self) -> LicenseType:
        return self._license_type

    @license_type.setter
    def license_type(self, license_type: LicenseType | str) -> None:
        self._license_type = LicenseType(license_type) if isinstance(license_type, str) else license_type
        try:
            self._plantsim.SetLicenseType(self._license_type.value)
        except BaseException as e:
            if ErrorCode.extract(e.args) == -2147221503:
                raise Exception(
                    f"The license type {self.license_type} does not seem to exist. Make sure it is a valid Plant Simulation license type."
                ) from e

    @property
    def visible(self) -> bool:
        return self._visible

    @visible.setter
    def visible(self, visible: bool) -> None:
        self._visible = visible
        self._plantsim.SetVisible(visible)

    @property
    def trust_models(self) -> bool:
        return self._trust_models

    @trust_models.setter
    def trust_models(self, trust_models: bool) -> None:
        self._trust_models = trust_models
        self._plantsim.SetTrustModels(trust_models)

    @property
    def path_context(self) -> str:
        return self._path_context

    @path_context.setter
    def path_context(self, path_context: str) -> None:
        self._path_context = path_context
        self._plantsim.SetPathContext(self._path_context)

    @property
    def event_controller(self) -> str:
        return self._event_controller

    @event_controller.setter
    def event_controller(self, path: str) -> None:
        self._event_controller = path

    def reset_simulation(self, event_controller: str | None = None) -> None:
        event_controller = event_controller if event_controller is not None else self.event_controller
        self._plantsim.ResetSimulation(event_controller)

    def start_simulation(
        self,
        event_controller: str | None = None,
        *,
        reset: bool = True,
        seed: int | None = None,
        wait_until_finished: bool = True,
    ) -> None:
        event_controller = event_controller if event_controller is not None else self.event_controller

        if reset:
            self.reset_simulation(event_controller)

        if seed is not None:
            self.set_value(f"{event_controller}.RandomNumbersVariant", seed)

        self._plantsim.StartSimulation(event_controller)

        while self.is_simulation_running() and wait_until_finished:
            pass

    def is_simulation_running(self) -> bool:
        return self._plantsim.IsSimulationRunning()

    def stop_simulation(self) -> None:
        self._plantsim.StopSimulation()

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
        return self._plantsim.GetValue(object_name)

    def set_value(self, object_name: str, value: Any) -> None:
        self._plantsim.SetValue(object_name, value)

    def execute_simtalk(self, command_string: str, *parameter: tuple[Any, ...], from_path_context: bool = False) -> Any:
        """
        Execute a SimTalk command according to COM documentation:
        PlantSim.ExecuteSimTalk("->real; return 3.14159")
        PlantSim.ExecuteSimTalk("param r:real->real; return r*r", 3.14159)

        :param command_string: Command to be executed
        :param parameter: (optional); parameter, if command contains a parameter to be set
        :param from_path_context: if true, command should be formulated from the path context
        """
        if from_path_context:
            command_string = f"{self.path_context}.{command_string}"

        # Allow multiline commands
        command_string = inspect.cleandoc(command_string)

        if parameter:
            return self._plantsim.ExecuteSimTalk(command_string, parameter)

        return self._plantsim.ExecuteSimTalk(command_string)

    def quit(self) -> None:
        self._plantsim.Quit()
