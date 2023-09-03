import os
import openpyxl
import tkinter as tk
from tkinter import filedialog

def edit_excel_files():
    folder_path = folder_path_var.get()
    start_cell = start_cell_var.get()
    end_cell = end_cell_var.get()
    cell_value = cell_value_var.get()
    selected_sheet = sheet_var.get()

    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)

            try:
                # Buka file Excel
                workbook = openpyxl.load_workbook(file_path)

                # Pilih sheet yang dipilih oleh pengguna
                sheet = workbook[selected_sheet]

                # Ekstrak kolom dan baris dari sel awal dan sel akhir
                start_column_letter = start_cell[0]
                start_row_number = int(start_cell[1:])
                end_column_letter = end_cell[0]
                end_row_number = int(end_cell[1:])

                # Edit data pada sel dalam rentang yang ditentukan
                for row_number in range(start_row_number, end_row_number + 1):
                    for column_letter in range(openpyxl.utils.column_index_from_string(start_column_letter), openpyxl.utils.column_index_from_string(end_column_letter) + 1):
                        cell = sheet.cell(row=row_number, column=column_letter)
                        cell.value = cell_value

                # Simpan perubahan
                workbook.save(file_path)
                result_label.config(text=f"Data diubah pada {file_path}")
            except Exception as e:
                result_label.config(text=f"Terjadi kesalahan pada {file_path}: {str(e)}")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_path_var.set(folder_selected)

# Buat GUI
root = tk.Tk()
root.title("Aplikasi Excel Editor")
# Atur latar belakang root window menjadi biru pastel
root.configure(bg='#b4c8d3')

folder_path_var = tk.StringVar()
start_cell_var = tk.StringVar()
end_cell_var = tk.StringVar()
cell_value_var = tk.StringVar()
sheet_var = tk.StringVar()

folder_label = tk.Label(root, text="Pilih folder yang berisi file Excel:",font=("Poppins",14),bg='#b4c8d3')
folder_label.pack(pady=5)
    
browse_button = tk.Button(root, text="Browse", command=browse_folder,font=("Poppins",14),bg='#ffffff')
browse_button.pack(pady=5)

start_cell_label = tk.Label(root, text="Sel Awal (contoh: A1):",font=("Poppins",14),bg='#b4c8d3')
start_cell_label.pack(pady=5)

start_cell_entry = tk.Entry(root, textvariable=start_cell_var,font=("Poppins",14),bg='#ffffff')
start_cell_entry.pack(pady=5)

end_cell_label = tk.Label(root, text="Sel Akhir (contoh: B5):",font=("Poppins",14),bg='#b4c8d3')
end_cell_label.pack(pady=5)

end_cell_entry = tk.Entry(root, textvariable=end_cell_var,font=("Poppins",14),bg='#ffffff')
end_cell_entry.pack(pady=5)

cell_value_label = tk.Label(root, text="Nilai yang akan dimasukkan:",font=("Poppins",14),bg='#b4c8d3')
cell_value_label.pack(pady=5)

cell_value_entry = tk.Entry(root, textvariable=cell_value_var,font=("Poppins",14),bg='#ffffff')
cell_value_entry.pack(pady=5)

sheet_label = tk.Label(root, text="Nama Sheet (contoh: Sheet1):",font=("Poppins",14),bg='#b4c8d3')
sheet_label.pack(pady=5)

sheet_entry = tk.Entry(root, textvariable=sheet_var,font=("Poppins",14),bg='#ffffff')
sheet_entry.pack(pady=5)

edit_button = tk.Button(root, text="Running", command=edit_excel_files,font=("Poppins",14),bg='#ffffff')
edit_button.pack(pady=5)

result_label = tk.Label(root, text="",font=("Poppins",14),bg='#b4c8d3')
result_label.pack(pady=5)

root.mainloop()
