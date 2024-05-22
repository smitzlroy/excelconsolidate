# Excel Consolidator

This Python script consolidates specific data from multiple Excel files into a single workbook. It is specifically designed to handle workbooks related to VMware environments, extracting critical information about virtual machines from a worksheet named "vInfo".

## Features

- **Worksheet Specificity**: Targets only the 'vInfo' worksheet in each Excel file.
- **Column Filtering**: Consolidates only selected columns relevant to VMware infrastructure management.
- **User-Friendly**: Utilizes a graphical file picker to select multiple Excel files for processing.

## Prerequisites

Before running this script, ensure you have the following installed:

- Python 3.6 or higher
- `openpyxl` library

You can install `openpyxl` using pip:


## Usage

To use the script, follow these steps:

1. Clone this repository or download the script to your local machine.
2. Ensure that your Python environment is set up with the necessary dependencies.
3. Run the script using a Python interpreter:


4. Select the Excel files you wish to consolidate when prompted by the file dialog.

The script will create or update a file named `consolidated_vInfo.xlsx` in the script's directory, containing the consolidated data.

## Columns Consolidated

The script specifically consolidates the following columns:

- VM
- Powerstate
- Template
- CPUs
- Memory
- Provisioned MiB
- In Use MiB
- Unshared MiB
- Datacenter
- Cluster
- Host
- OS according to the configuration file
- OS according to the VMware Tools

## Contributing

Contributions to this project are welcome! Please fork the repository and submit a pull request with your enhancements. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.
