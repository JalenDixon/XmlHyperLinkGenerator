# File Hyperlink Generator

This script generates an Excel (.xlsx) file that contains hyperlinks to all files in a specified directory.

## Description

The Python script scans a defined directory for files. For each file, it creates a hyperlink to the file's location and adds it to an Excel spreadsheet. The spreadsheet also includes the file's name and its position in the directory's list of files.

The script uses the `os`, `pandas`, and `xlsxwriter` modules to achieve this. `os` is used to interact with the file system, `pandas` is used to create the data structure for the spreadsheet, and `xlsxwriter` is used to write the data to an Excel file.

## Setup & Execution

### Dependencies

This script requires the following Python packages:

- pandas
- xlsxwriter

You can install these packages using pip:

```sh
pip install pandas xlsxwriter
```
### Execution
To run the script, navigate to its directory in a terminal window and run the following command:

```sh 
python main.py
```

Before running, ensure you've updated the folder_path and output_dir variables at the top of the script to match your specific use case.

Output
The script creates an Excel file in the specified output directory. The file is named "Hyperlinks_<current_date>.xlsx". The Excel file includes three columns: 'Exhibit Number', 'Description', and 'Hyperlink'. The 'Exhibit Number' column contains the file's position in the directory listing, the 'Description' column contains the file's name, and the 'Hyperlink' column contains a clickable hyperlink to the file.

You can save this text as `README.md` in the same directory as the script. This README can serve as a basic guide for someone else trying to understand what the script does and how to use it.
