import pandas as pd
import os
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side


def main():
    input_path = os.path.join('..', 'data', 'jira_data.xlsx')
    output_path = ""
    df = create_df_for_cpbc(input_path)
    create_reports_from_data(df, output_path)
    return


def create_df_for_cpbc(path):
    try:
        # Read the execl file to create a df to work with
        expanded_df = pd.read_excel(path)

        df = expanded_df[['Time Spent', 'Assignee', 'Sprint', 'Custom field (Budget)']]
        df['Time Spent (Days)'] = df['Time Spent'] / 3600 / 8
        df.drop(columns=['Time Spent'], inplace=True)

        df_by_departments = departments_df(df)

        print(df_by_departments)

        filtered_df = formation_of_df_for_excel(df)

        print(filtered_df)
        return filtered_df

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def departments_df(df: object):
    # Construct the relative path to the CSV file
    csv_file_path = os.path.join('..', 'config', 'worker-names.csv')

    # Read the CSV file
    worker_config = pd.read_csv(csv_file_path)

    # Create a mapping from assignee names to their corresponding columns in df1
    name_to_column = {}
    for column in worker_config.columns:
        for name in worker_config[column].dropna():
            name_to_column[name] = column

    # Add the corresponding column information to df2
    df['Team'] = df['Assignee'].map(name_to_column)

    # Drop rows where 'Time Spent' is NaN
    df = df.dropna(subset=['Time Spent (Days)'])

    # Initialize a dictionary to store the DataFrames
    category_dfs = {}

    # Loop through each category column in df1
    for category in worker_config.columns:
        # Get the list of assignees for the current category
        assignees = worker_config[category].dropna().tolist()
        # Filter df2 based on the assignees list
        filtered_df = df[df['Assignee'].isin(assignees)]
        # Apply the clean_string function to the 'Custom field (Budget)' column
        filtered_df['Custom field (Budget)'] = filtered_df['Custom field (Budget)'].apply(clean_string)
        filtered_df.drop(columns=['Team'], inplace=True)
        # Group by 'Sprint' and 'Custom field (Budget)' and sum the 'Time Spent (Days)'
        df_aggregated = filtered_df.groupby(['Sprint', 'Custom field (Budget)'], as_index=False)['Time Spent (Days)'].sum()
        # Pivot the DataFrame without aggregation
        pivot_df = df_aggregated.pivot(index='Sprint', columns='Custom field (Budget)', values='Time Spent (Days)')
        # Fill NaN values with 0
        pivot_df = pivot_df.fillna(0)
        # Reset index to convert the pivoted DataFrame back to a normal DataFrame
        pivot_df = pivot_df.reset_index()
        # Print the resulting DataFrame
        print(pivot_df)
        # Store the filtered DataFrame in the dictionary
        category_dfs[category] = pivot_df
        # Save the resulting DataFrame to an Excel file
        pivot_df.to_excel(f'{category}pivoted_output.xlsx', index=False)

    return category_dfs


def clean_string(input_string):
    # Use regex to remove the content inside parentheses along with the space before it
    cleaned_string = re.sub(r'\s*\(.*?\)', '', input_string)
    return cleaned_string


def formation_of_df_for_excel(df):
    # Add new columns with default values
    df['Entry Type'] = 'ACTUAL'
    df['Employee ID'] = None
    df['Exp Type'] = 'Ongoing task'
    df['Jira name'] = None
    df['Employee Name'] = 'Total'
    df['Approved by'] = None
    df['New Project 1'] = None
    df['New Project 2'] = None

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
            # Append the DataFrame for the current group to the list
            grouped_dfs.append([shorten_name(group_name), group_df])

    return grouped_dfs
