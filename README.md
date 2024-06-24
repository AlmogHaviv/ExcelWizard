# ExcelWizard

ExcelWizard is a Python automation tool designed to transform and manipulate Excel files with ease. It is particularly useful for generating specific reports from raw data files, including CPBC (Cross-Platform Business Center) and CPBE (Cross-Platform Business Excellence) reports.

## Features

- Read data from Excel files
- Perform data transformation and cleanup
- Generate new Excel files with processed data
- Apply custom formatting and styling to Excel sheets
- Support for generating specific reports for different departments (e.g., CPBC reports for Algo, Bi, Dev, Product, Devops, System Architect, and CPB Directors)
- Support for generating CPBE reports

## Installation

1. Clone the repository: git clone https://github.com/your-username/ExcelWizard.git
2. Install the required dependencies: pip install -r requirements.txt
   
## Usage

if you want to use the CPBC script:
   1. Place your input Excel file(s) in the `data` directory.
   2. Run the `CPBC_all.py` script with the desired options:
      Options:
      - `-all`: Generate reports for all departments
      - `-product`: Generate Product report
      - `-devops`: Generate Devops report
      - `-algo`: Generate Algo report
      - `-dev`: Generate Dev report
      - `-bi`: Generate BI report
      - `-sa`: Generate System Architect report
      - `-cpb`: Generate CPB Directors report
      - `-input_file INPUT_FILE`: Specify the input file name
      - `-wd WORKING_DIRECTORY`: Specify the working directory
      please notice you have to fullfill -input_file and -wb and one of the other flags, otherwise it will return an error

if you want to use the CPBE script:
   1. Place your input Excel file(s) in the `data` directory.
   2. Run the `CPBE_first_page.py` script with the desired options:
      Options:
      - `-input_file INPUT_FILE`: Specify the input file name
      - `-wd WORKING_DIRECTORY`: Specify the working directory
      please notice you have to fullfill -input_file and -wb, otherwise it will return an error

## Project Structure

- CPBE_first_page.py: Script to generate the CPBE report.
- CPBC_all.py: Script to generate the CPBC report for all departments.
- data/: Directory where input Excel files should be placed and where output files will be generated.
- config/: Directory containing configuration files such as worker-names.csv and fte_contract.csv.
- requirements.txt: List of dependencies required for the project.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
