import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import openpyxl
import os


def process_data():
    source_file_path = source_file_var.get()
    target_file_path = target_file_var.get()
    subjects_of_interest = subjects_entry.get().split(',')  # Assumes subjects are entered as a comma-separated list

    try:
        df1 = pd.read_excel(source_file_path, index_col="Student ID")
        df1 = df1.rename(columns=lambda x: x.replace(' ', '_'))
        count_per_value = df1['Student_Name'].value_counts()
        df1 = df1[df1['Student_Name'].map(count_per_value) == 1]
        if df1[df1['Course_Name'].isin(subjects_of_interest)].shape[0] > 0:
            df1 = df1[df1['Course_Name'].isin(subjects_of_interest)]
            df1.to_excel(target_file_path, index=True)
            messagebox.showinfo("Success", f"Num of Eligible Students: {df1.shape[0]}\nDataFrame saved to {target_file_path}")
        else:
            messagebox.showerror("Error: Please Enter valid course ID")


    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def browse_source_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx;.xls")])
    source_file_var.set(file_path)

def browse_target_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    target_file_var.set(file_path)

# GUI setup
app = tk.Tk()
app.title("Student Course filtration")
window_width = 800
window_height = 400

screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
app.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

font_style = ("Helvetica", 12)
bg_color = "#4BCFFA"  # Light gray background color
label_bg_color = "#4BCFFA"
button_color = "#4CAF50"  # Green color for buttons
button_fg_color = "white"  # White text color for buttons
 
app.configure(bg=bg_color)

source_file_var = tk.StringVar()
target_file_var = tk.StringVar()

tk.Label(app, text="Source Excel File:", font=font_style, bg = "#4BCFFA").pack(pady=10)
tk.Entry(app, textvariable=source_file_var, width=50).pack(pady=10)
tk.Button(app, text="Browse", command=browse_source_file, font=font_style, bg=button_color, fg=button_fg_color).pack(pady=10)

tk.Label(app, text="Target Excel File:", font=font_style, bg = "#4BCFFA").pack(pady=10)
tk.Entry(app, textvariable=target_file_var, width=50).pack(pady=10)
tk.Button(app, text="Browse", command=browse_target_file, font=font_style, bg=button_color, fg=button_fg_color).pack(pady=10)

tk.Label(app, text="Subjects of Interest (Eg: EG-316):", font=font_style, bg = "#4BCFFA").pack(pady=10)
subjects_entry = tk.Entry(app, width=50)
subjects_entry.pack(pady=10)

tk.Button(app, text="Proceed", command=process_data, font=font_style, bg=button_color, fg=button_fg_color).pack(pady=10)


app.mainloop()