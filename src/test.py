import CPBC_all
import os


input_path = os.path.join('..', 'data', 'jira_data.xlsx')

CPBC_all.create_df_for_cpbc(input_path)
