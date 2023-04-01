import os
import glob
import comtypes.client
import tkinter as tk
from tkinter import filedialog, messagebox

# Create a GUI window to select input directory
root = tk.Tk()
root.withdraw()
input_dir = filedialog.askdirectory(title='Select Input Directory')

# Change backslashes to forward slashes in input directory
input_dir = input_dir.replace('\\', '/')

# Get list of .doc files in directory
doc_files = glob.glob(os.path.join(input_dir, '*.doc'))

# Define function to convert .doc to .docx
def doc_to_docx(doc_file):
    # Define input and output filenames
    input_file = os.path.abspath(doc_file)
    output_file = os.path.abspath(doc_file + 'x')
    
    # Load Word COM server
    word = comtypes.client.CreateObject('Word.Application')
    
    # Load input file
    doc = word.Documents.Open(input_file)
    
    # Save as .docx file
    doc.SaveAs(output_file, FileFormat=16)
    
    # Close input file and Word COM server
    doc.Close()
    word.Quit()

# Convert each .doc file to .docx
for doc_file in doc_files:
    doc_to_docx(doc_file)

# Show message box to indicate completion
messagebox.showinfo(title='Conversion Complete', message='All files have been converted to .docx format.')
