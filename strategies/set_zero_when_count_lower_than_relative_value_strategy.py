from openpyxl.comments import Comment
from openpyxl.styles import Alignment, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from strategies.transform_strategy import TransformStrategy


class SetZeroWhenNumberLowerThanRelativeValueStrategy(TransformStrategy):
    """
    Strategy that sets to zero the integer and percentage values in pairs of columns when the integer value
    relative to the count column (in percentage) is less than a relative value (in percentage).
    """

    DEFAULT_MIN_RELATIVE_VALUE_PERCENTAGE = 0.1

    __GRAY_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    __COLUMN_ADDED_BY_SCRIPT_COMMENT = Comment('This column was added by the Python script', '')

    def __init__(self, min_relative_value_percentage=None):
        self.min_relative_value_percentage = min_relative_value_percentage or self.DEFAULT_MIN_RELATIVE_VALUE_PERCENTAGE

    def transform_excel_sheet(self, sheet: Worksheet, start_column_index: int, count_column_name: str) -> None:
        # Add new columns for the percentage values
        col = start_column_index
        while col < sheet.max_column:
            # Insert a new column to the right of the current column
            sheet.insert_cols(col + 1)
            # Give a name to the cell in the first row of the new column
            new_col_name_cell = sheet.cell(row=1, column=col + 1, value=f"{sheet.cell(row=1, column=col).value} percentage")
            # Modify the style of the cell in the first row of the current column
            new_col_name_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            new_col_name_cell.fill = self.__GRAY_FILL
            new_col_name_cell.comment = self.__COLUMN_ADDED_BY_SCRIPT_COMMENT

            col += 3  # Move to the next pair of columns (having into account the new added column)

        # Iterate over the rows starting from the second row
        for row in range(2, sheet.max_row + 1):
            # Get the count value for the current row
            count_column_index = self.__get_column_index_from_name(sheet, count_column_name)
            count_value = sheet.cell(row=row, column=count_column_index).value

            col = start_column_index
            while col < sheet.max_column:
                int_value_cell = sheet.cell(row=row, column=col)
                relative_value_cell = sheet.cell(row=row, column=col + 1)  # cell added by the script to store the relative value (in percentage)
                perc_value_cell = sheet.cell(row=row, column=col + 2)

                # Check if the integer value relative to the count column is less than the relative value
                if int_value_cell.value is not None:
                    relative_value_percentage = int_value_cell.value / count_value * 100
                    relative_value_cell.value = relative_value_percentage
                    if relative_value_percentage < self.min_relative_value_percentage:
                        int_value_cell.value = 0
                        int_value_cell.font = self.red_font  # Set the text color to red

                        relative_value_cell.font = self.red_font  # Set the text color to red
                        relative_value_cell.comment = Comment(  # Set comment explaining the change
                            f'{relative_value_percentage} < {self.min_relative_value_percentage}, so the cells to the left and right '
                            f'were set to 0, because the number of values is not significant enough.',
                            '')

                        perc_value_cell.value = 0
                        perc_value_cell.font = self.red_font  # Set the text color to red

                col += 3  # Move to the next pair of columns (having into account the new added column)

    def __get_column_index_from_name(self, sheet: Worksheet, column_name: str):
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value == column_name:
                return col
        return None
