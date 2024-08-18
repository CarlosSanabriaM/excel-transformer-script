import openpyxl

from strategies.set_zero_when_count_lower_than_absolute_value_strategy import SetZeroWhenNumberLowerThanAbsoluteValueStrategy
from strategies.set_zero_when_count_lower_than_relative_value_strategy import SetZeroWhenNumberLowerThanRelativeValueStrategy

STRATEGIES_MAP = {
    'ABSOLUTE': SetZeroWhenNumberLowerThanAbsoluteValueStrategy(),
    'RELATIVE': SetZeroWhenNumberLowerThanRelativeValueStrategy()
}


# noinspection PyShadowingNames
def transform_excel(transform_strategy, input_file, output_file, start_column_name, count_column_name):
    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active

    # Find the start column index
    start_column_index = None
    for col in range(1, sheet.max_column + 1):
        if sheet.cell(row=1, column=col).value == start_column_name:
            start_column_index = col
            break

    if start_column_index is None:
        raise ValueError(f"Column '{start_column_name}' not found in the Excel file.")

    # Use the selected strategy to transform the sheet
    strategy = STRATEGIES_MAP.get(transform_strategy)
    strategy.transform_excel_sheet(sheet, start_column_index, count_column_name)

    # Save the modified workbook preserving the formatting
    wb.save(output_file)
    print(f"Modified Excel successfully saved as '{output_file}'")


if __name__ == "__main__":
    # Update these values
    transform_strategy = 'RELATIVE'  # Replace with your desired strategy
    input_excel_file = 'input.xlsx'  # Replace with your input file name
    output_excel_file = 'output.xlsx'  # Replace with your desired output file name
    start_column_name = 'First column to modify'  # Replace with your specific start column name
    count_column_name = 'Num events'  # Replace with your specific count column name

    transform_excel(transform_strategy, input_excel_file, output_excel_file, start_column_name, count_column_name)
