from abc import ABC, abstractmethod

from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet


class TransformStrategy(ABC):
    """
    Interface for all supported versions of the transform algorithm over the input Excel sheet.
    """

    # Define the red font
    red_font = Font(color="FF0000")

    @abstractmethod
    def transform_excel_sheet(self, sheet: Worksheet, start_column_index: int, count_column_name: str) -> None:
        pass
