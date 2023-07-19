import os
from datetime import datetime
from pathlib import Path
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory


def request_directory_from_user(prompt_message):
    Tk().withdraw()  # Hide the root window as we don't want a full GUI
    directory_path = askdirectory(title=prompt_message)  # Open dialog box and return the path
    return directory_path


def create_output_file_path(output_dir):
    current_date = datetime.now().strftime('%m-%d-%Y')
    output_file_name = f'Hyperlinks_{current_date}.xlsx'
    output_file_path = os.path.join(output_dir, output_file_name)
    return output_file_path


def generate_data_from_files(directory):
    data = {'Exhibit Number': [], 'Description': [], 'Hyperlink': []}

    for file_number, file_name in enumerate(os.listdir(directory), start=1):
        file_path = os.path.join(directory, file_name)
        file_uri = Path(file_path).as_uri()

        data['Exhibit Number'].append(file_number)
        data['Description'].append(file_name)
        data['Hyperlink'].append(file_uri)

    return data


def create_excel_file(dataframe, output_file_path):
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

    dataframe.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    header_format = workbook.add_format({'bold': True, 'text_wrap': True})

    for column_number, column_name in enumerate(dataframe.columns.values):
        worksheet.write(0, column_number, column_name, header_format)

    for row_number, hyperlink in enumerate(dataframe['Hyperlink'], start=1):
        worksheet.write_url(row_number, 2, hyperlink)

    writer.close()


def main():
    input_dir = request_directory_from_user('Please select the directory to get files from')
    if not input_dir:
        print("No input directory selected. Exiting.")
        return

    output_dir = request_directory_from_user('Please select the output directory')
    if not output_dir:
        print("No output directory selected. Exiting.")
        return

    os.makedirs(output_dir, exist_ok=True)
    output_file_path = create_output_file_path(output_dir)

    if os.path.exists(output_file_path):
        os.remove(output_file_path)

    file_data = generate_data_from_files(input_dir)
    dataframe = pd.DataFrame(file_data)

    create_excel_file(dataframe, output_file_path)


if __name__ == "__main__":
    main()
