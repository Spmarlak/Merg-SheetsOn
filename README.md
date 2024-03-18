# Merg-SheetsOn
# Merge mutiple sheets on excel to one 


import pandas as AJ
from tkinter import *
from tkinter import filedialog
import os
filepath = ""
def openFile():
    global filepath
    filepath = filedialog.askopenfilename()
    print(filepath)
    window.destroy()
window = Tk()
button = Button(text="Select Excel File", command=openFile)
button.pack()
window.mainloop()
if filepath:
    AJxls = AJ.ExcelFile(filepath)
    AJdataframes = []
    for sheet_name in AJxls.sheet_names:
        df = AJ.read_excel(filepath, sheet_name=sheet_name)
        AJdataframes.append(df)
    combined_data = AJ.concat(AJdataframes, ignore_index=True)
    directory_path = os.path.dirname(filepath)
    combined_file_path = os.path.join(directory_path, 'Data_Combined_GUI.xlsx')
    combined_data.to_excel(combined_file_path, index=False)
    print("Data from all Sheets combined and saved to", combined_file_path)
else:
    print("No file selected")
