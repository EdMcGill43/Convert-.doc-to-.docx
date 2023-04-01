import os
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import messagebox

# Set the folder path where the Word files are located
folder_path = '/path/to/word/files'

# Get all the .docx files in the folder
docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.docx')]

# Extract the data from each Word file and store it in a list
data_list = []
for docx_file in docx_files:
    document = Document(docx_file)
    for para in document.paragraphs:
        data_list.append([para.text])

# Create a Pandas dataframe with the extracted data
df = pd.DataFrame(data_list, columns=['Data'])

# Write the dataframe to an Excel file
excel_file = '/path/to/excel/file.xlsx'
df.to_excel(excel_file, index=False)

# Show a pop-up message when the script is finished running
root = tk.Tk()
root.withdraw()
messagebox.showinfo(title='Extraction Complete', message='The data extraction and conversion to Excel file is completed.')
