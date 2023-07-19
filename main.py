import os
from datetime import datetime
from pathlib import Path
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory

def main():
    # Ask the user for the input directory
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    folder_path = askdirectory(title='Please select the directory to get files from') # show an "Open" dialog box and return the path
    if not folder_path: # if user clicked cancel, then folder_path will be an empty string
        print("No input directory selected. Exiting.")
        return

    # Ask the user for the output directory
    output_dir = askdirectory(title='Please select the output directory') # show an "Open" dialog box and return the path
    if not output_dir: # if user clicked cancel, then output_dir will be an empty string
        print("No output directory selected. Exiting.")
        return

    os.makedirs(output_dir, exist_ok=True)

    # Get the current date
    date = datetime.now().strftime('%m-%d-%Y')

    # Output file
    output_file = os.path.join(output_dir, f'Hyperlinks_{date}.xlsx')

    # Check if the file already exists and remove it
    if os.path.exists(output_file):
        os.remove(output_file)

    # Prepare a DataFrame to hold the data
    data = {'Exhibit Number': [], 'Description': [], 'Hyperlink': []}

    # Loop through each file in the directory
    for i, filename in enumerate(os.listdir(folder_path), start=1):
        file_path = os.path.join(folder_path, filename)

        # Create a Path object and then get the URI
        file_uri = Path(file_path).as_uri()

        # Add the exhibit number (i), description (filename) and hyperlink to the data
        data['Exhibit Number'].append(i)
        data['Description'].append(filename)
        data['Hyperlink'].append(file_uri)

    df = pd.DataFrame(data)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Convert the DataFrame to an XlsxWriter Excel object. Note that we turn off
    # the default header and skip one row to allow us to insert a user defined
    # header.
    df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add a header format.
    header_format = workbook.add_format({'bold': True, 'text_wrap': True})

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Write hyperlinks
    for row_num, link in enumerate(df['Hyperlink'], start=1):
        worksheet.write_url(row_num, 2, link)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()


if __name__ == "__main__":
    main()
