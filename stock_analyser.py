from analyse_data import analyze_file
from tkinter import filedialog
from tkinter import *

window = Tk()
window.title('Stock Analyzer')


# Actions
def file_selected():
    file_selected = filedialog.askopenfilename()
    input_file_path.delete(0, "end")
    input_file_path.insert(0, file_selected)
    # cacheFile = open("cache.txt", "w")
    # cacheFile.write(folder_selected)
    # cacheFile.close()
    print(input_file_path.get())


def folder_selected():
    folder_selected = filedialog.askdirectory()
    output_file_path.delete(0, "end")
    output_file_path.insert(0, folder_selected)
    # cacheFile = open("cache.txt", "w")
    # cacheFile.write(folder_selected)
    # cacheFile.close()
    final_output_dest_path = output_file_path.get() + "/" + "Result.xlsx"
    output_file_path.delete(0, "end")
    output_file_path.insert(0, final_output_dest_path)


def analyze():
    excel_data_file_path = input_file_path.get()
    analyzed_excel_data_file_path = output_file_path.get()
    analyze_file(excel_data_file_path, analyzed_excel_data_file_path)


# Input file labels
input_dest_label = Label(window, text="INPUT EXCEL FILE DIR ", font=("Calibri", 12))
input_dest_label.grid(column=0, row=0, sticky="W")
output_dest_label = Label(window, text="OUTPUT FILE NAME ", font=("Calibri", 12))
output_dest_label.grid(column=0, row=1, sticky="W")

# Path to files
input_file_path = Entry(window, width=65)
input_file_path.grid(column=1, row=0)
output_file_path = Entry(window, width=65)
output_file_path.grid(column=1, row=1)

# Buttons
btn_browse_input_file = Button(window, text="Browse", command=file_selected)
btn_browse_input_file.grid(column=2, row=0)
btn_browse_input_path = Button(window, text="Browse", command=folder_selected)
btn_browse_input_path.grid(column=2, row=1)
btn_analyze = Button(window, text="Analyze!", font=("Calibri", 12, "bold"), fg="green", command=analyze)
btn_analyze.grid(column=1, row=2, pady=40, padx=130, sticky="SW")

window.geometry('650x150')

if __name__ == '__main__':
    window.mainloop()