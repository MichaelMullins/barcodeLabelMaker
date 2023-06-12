import os
import pandas as pd
import tkinter as tk
import tkinter.filedialog as fd
from barcode.codex import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from PIL import Image


class App:
    def __init__(self, root):
        self.import_button = tk.Button(root, text="Import", command=self.import_from_excel)
        self.import_button.grid(row=0, column=0, padx=5, pady=5)
        self.entries = []
        for i in range(10):
            row = []
            for j in range(3):
                entry = tk.Entry(root)
                entry.grid(row=i+1, column=j, padx=5, pady=5)
                row.append(entry)
            self.entries.append(row)

        submit_button = tk.Button(root, text="Submit", command=self.generate_barcodes)
        submit_button.grid(row=11, column=1, pady=10)

    def import_from_excel(self):
        file_path = fd.askopenfilename(filetypes=(("Excel Files", "*.xlsx"),))
        if not file_path:
            return
        df = pd.read_excel(file_path, header=None)
        data = df.values.flatten()
        for i, content in enumerate(data):
            if pd.notna(content):
                self.entries[i // 3][i % 3].delete(0, 'end')
                self.entries[i // 3][i % 3].insert(0, str(content))

    def generate_barcodes(self):
        barcode_dir = 'barcodes'
        if not os.path.exists(barcode_dir):
            os.makedirs(barcode_dir)
        else:
            # Remove existing barcode images
            for file_name in os.listdir(barcode_dir):
                if file_name.endswith('.png'):
                    os.remove(os.path.join(barcode_dir, file_name))

        barcode_files = []
        index = 1
        for row in self.entries:
            for entry in row:
                content = entry.get()
                file_path = os.path.join(barcode_dir, f'barcode{index}.png')
                if content:
                    with open(file_path, 'wb') as f:
                        Code128(content, writer=ImageWriter()).write(f)
                else:
                    # If input is empty, generate a blank white png
                    img = Image.new('RGB', (100, 100), color=(255, 255, 255))
                    img.save(file_path)
                barcode_files.append(file_path.replace('\\', '\\\\'))
                index += 1

        self.create_excel(barcode_files)

    def create_excel(self, barcode_files):
        wb = Workbook()
        ws = wb.active
        ws.append(["Barcode"+str(i+1) for i in range(len(barcode_files))])  # Headers
        ws.append(barcode_files)  # File paths
        wb.save("barcode_directories.xlsx")

root = tk.Tk()
root.title("Barcode Generator")
app = App(root)
root.mainloop()
