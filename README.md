# ExcelWizard

ExcelWizard is a Python automation tool designed to transform and manipulate Excel files with ease.
It is particularly useful for generating specific reports from raw data files.

## Features

- Read data from Excel files
- Perform data transformation and cleanup
- Generate new Excel files with processed data
- Apply custom formatting and styling to Excel sheets
- Supports generating specific reports for different departments (e.g., CPBC and CPBE reports)

## Installation

1. Clone the repository: git clone https://github.com/your-username/ExcelWizard.git
2. Install the required dependencies: pip install -r requirements.txt
   
## Usage

if you want to use the CPBE flag:
   1. Place your input Excel file(s) in the `data` directory under the name 'Experimental_Project_days_template.xlsx'.
   2. Run the `main.py` script: python main.py -e
   3. The processed Excel files will be generated in the `data` directory.
if you want to use the CPBC flag:
   1. Place your input Excel file(s) in the data directory under the name jira_data.xlsx.
   2. To generate a CPBC report for a specific department:
         python main.py -c department_name
      Replace department_name with the specific department you want the report for (e.g., Algo, Bi, Dev).
   3. The processed Excel files will be generated in the data directory.

## Project Structure

- CPBE_first_page.py: Script to generate the CPBE report.
- CPBC_all.py: Script to generate the CPBC report for all departments.
- data/: Directory where input Excel files should be placed and where output files will be generated.
- config/: Directory containing configuration files such as worker-names.csv and fte_contract.csv.
- main.py: Main entry point for the script, parses arguments and triggers report generation.
- requirements.txt: List of dependencies required for the project.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
