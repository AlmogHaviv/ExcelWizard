import pandas as pd
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side


def read_excel_file(file_path):
    """
    Read data from an Excel file, perform necessary transformations, and prepare the data for further processing.

    Args:
    - file_path (str): Path to the Excel file.

    Returns:
    - tuple: A tuple containing a list of DataFrames grouped by role, and the headers row DataFrame.
    """
    try:
        # Read the first row to retrieve headers information
        first = pd.read_excel(file_path)
        headers_row = first.iloc[[0]]
        headers_row.insert(0, 'Unnamed - 1', '')

        # Read the Excel files, skipping the first row to avoid duplication
        df = pd.read_excel(file_path, header=1)

        # Drop unnecessary columns
        columns_to_drop = ['Department Units', 'M.W.D', 'month', 'Special notes']
        df.drop(columns=columns_to_drop, inplace=True)

        # Add new columns with default values
        df['Entry Type'] = 'Budget'
        df['Employee ID'] = None
        df['Exp Type'] = None
        df['Jira name'] = None
        df['Employee Name'] = 'Total'
        df['Approved by'] = None

        # Reorganize the order of columns
        cols = list(df.columns)
        cols.insert(0, cols.pop(cols.index('Approved by')))
        cols.insert(0, cols.pop(cols.index('Exp Type')))
        cols.insert(0, cols.pop(cols.index('Employee Name')))
        cols.insert(0, cols.pop(cols.index('Jira name')))
        cols.insert(0, cols.pop(cols.index('Department')))
        cols.insert(0, cols.pop(cols.index('Role Ending')))
        cols.insert(0, cols.pop(cols.index('Exp Type')))
        cols.insert(0, cols.pop(cols.index('Employee ID')))
        cols.insert(0, cols.pop(cols.index('Entry Type')))
        df = df[cols]

        # Group the concatenated DataFrame by 'Department'
        grouped_data = df.groupby('Department')

        # Initialize a list to store the 6 DataFrames
        grouped_dfs = []

        # Iterate over the groups and create a DataFrame for each group
        for group_name, group_df in grouped_data:
            if group_name != "Total":
                # Drop the 'Role Ending' column as it's not needed anymore
                group_df.drop(columns=['Department'], inplace=True)
                # Append the DataFrame for the current group to the list
                grouped_dfs.append([shorten_name(group_name), group_df])

        return grouped_dfs, headers_row

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def read_excel_file_for_page_2(file_path):
    """
    Read data from an Excel file, perform necessary transformations, and prepare the data for further processing.

    Args:
    - file_path (str): Path to the Excel file.

    Returns:
    - tuple: A tuple containing a list of DataFrames grouped by role, and the headers row DataFrame.
    """
    try:
        # Read the first row to retrieve headers information
        first = pd.read_excel(file_path)
        headers_row = first.iloc[[0]]
        headers_row.insert(0, 'Unnamed - 1', '')

        # Read the Excel files, skipping the first row to avoid duplication
        df = pd.read_excel(file_path, header=1, sheet_name=1)

        # Drop unnecessary columns
        columns_to_drop = ['Department Units', 'M.W.D', 'month', 'Special notes']
        df.drop(columns=columns_to_drop, inplace=True)

        # Add new columns with default values
        df['Entry Type'] = 'Budget'
        df['Employee ID'] = None
        df['Exp Type'] = None
        df['Jira name'] = None
        df['Employee Name'] = 'Total'
        df['Approved by'] = None

        # Reorganize the order of columns
        cols = list(df.columns)
        cols.insert(0, cols.pop(cols.index('Approved by')))
        cols.insert(0, cols.pop(cols.index('Exp Type')))
        cols.insert(0, cols.pop(cols.index('Employee Name')))
        cols.insert(0, cols.pop(cols.index('Jira name')))
        cols.insert(0, cols.pop(cols.index('Department')))
        cols.insert(0, cols.pop(cols.index('Role Ending')))
        cols.insert(0, cols.pop(cols.index('Exp Type')))
        cols.insert(0, cols.pop(cols.index('Employee ID')))
        cols.insert(0, cols.pop(cols.index('Entry Type')))
        df = df[cols]

        # Group the concatenated DataFrame by 'Department'
        grouped_data = df.groupby('Role Ending')

        # Initialize a list to store the 6 DataFrames
        grouped_dfs = []

        # Iterate over the groups and create a DataFrame for each group
        for group_name, group_df in grouped_data:
            if group_name != "Total":
                # Drop the 'Role Ending' column as it's not needed anymore
                group_df.drop(columns=['Role Ending'], inplace=True)
                # Append the DataFrame for the current group to the list
                grouped_dfs.append([shorten_name(group_name), group_df])

        return grouped_dfs

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def manipulate_data(data, first_row, output_directory):
    """
    Perform data manipulation tasks.

    Args:
    - data (pandas DataFrame): DataFrame containing the input data.

    Returns:
    - pandas DataFrame: DataFrame containing the manipulated data.
    """
    # Get a list of the headers
    header_row = first_row.columns.tolist()
    for idx, name in enumerate(header_row):
        if "Unnamed" == name[:7]:
            header_row[idx] = None
    header_row2 = data[0][1].columns.tolist()
    # Iterate over each DataFrame
    for idx, df in enumerate(data):
        # Create a new Excel workbook
        wb = Workbook()

        # Create a new sheet
        ws = wb.active

        # Add the first row from my_row to the Excel sheet
        ws.append(header_row)

        # Add the first row from df to the Excel sheet
        ws.append(header_row2)

        # Append all other rows from df to the Excel sheet
        for row in df[1].itertuples(index=False):
            ws.append(list(row))

        style_excel(ws)

        # Define the path for the Excel file
        if df[0] == "PLM" or df[0] == "PlantGrowth":
            file_name = f"{output_directory}/{df[0]}_5_24_U.xlsx"
        elif df[0] == "QA":
            file_name = f"{output_directory}/PLM_5_24 QA.xlsx"
        elif df[0] == "GreenHouseControlled":
            file_name = f"{output_directory}/GreenHouse_5_24.xlsx"
        elif df[0] == 'FieldNonControlled':
            file_name = f"{output_directory}/Non Controlled GreenHouse_5_24.xlsx"
        else:
            file_name = f"{output_directory}/{df[0]}_5_24.xlsx"

        # Save the workbook to an Excel file
        wb.save(file_name)

        print(file_name)


# Define a function to extract the shortened name using regex
def shorten_name(name):
    split_string = re.split(r'\s+|-', name)
    return "".join(split_string[1:])


def style_excel(ws):
    # Define fill and font styles
    light_blue_fill = PatternFill(start_color='CCE5FF', end_color='CCE5FF', fill_type='solid')
    soft_green_fill = PatternFill(start_color='D3EAD3', end_color='D3EAD3', fill_type='solid')
    david_font = Font(name='David')

    # Apply styles to the second row
    for cell in ws[2]:
        cell.fill = light_blue_fill

    # Apply styles to columns 2-5 from the third row to the end
    for row in ws.iter_rows(min_row=3, min_col=2, max_col=10):
        for cell in row:
            cell.fill = soft_green_fill

    # Apply font style to all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.font = david_font

    # Add black borders between all parts of the data
    max_row = ws.max_row
    max_col = ws.max_column
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                         bottom=Side(style='thin'))

    for i in range(1, max_row + 1):
        for j in range(1, max_col + 1):
            ws.cell(row=i, column=j).border = thin_border


def main():
    input_file_path = os.path.abspath("../data/Experimental_Project_days_template.xlsx")
    output_file_path = os.path.abspath("../data/")
    # Step 1: Read input Excel file
    input_data, first_row = read_excel_file(input_file_path)
    second_page_data = read_excel_file_for_page_2(input_file_path)
    for df in second_page_data:
        input_data.append(df)

    if input_data is None:
        return

    # Step 2: excel it
    manipulate_data(input_data, first_row, output_file_path)


if __name__ == "__main__":
    main()
