import CPBC_all
import os


input_path = os.path.join('..', 'data', 'jira_data.xlsx')

dfs = CPBC_all.create_df_for_cpbc(input_path)

for dept, df in dfs.items():
    CPBC_all.create_full_scale_for_excel(dept, df, input_path)
