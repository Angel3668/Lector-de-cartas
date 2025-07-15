import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from pyzbar.pyzbar import decode
from openpyxl import Workbook
import subprocess

# Cropping coordinates for two variants (left, upper, right, lower)
VARIANT_1_BOX = (50, 50, 300, 300)
VARIANT_2_BOX = (100, 100, 350, 350)


class ScannerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide main window
        self.files = filedialog.askopenfilenames(title="Selecciona las imagenes")
        if not self.files:
            sys.exit()
        self.progress_win = tk.Toplevel()
        self.progress_win.title("Escaneando")
        self.cancel = False

        self.progress = ttk.Progressbar(self.progress_win, length=300, mode="determinate")
        self.progress.pack(padx=10, pady=10)
        self.label = tk.Label(self.progress_win, text="0 de {}".format(len(self.files)))
        self.label.pack()
        self.result_label = tk.Label(self.progress_win, text="")
        self.result_label.pack()
        cancel_btn = tk.Button(self.progress_win, text="Cancelar", command=self.cancel_scan)
        cancel_btn.pack(pady=5)

        self.scan_files()
        self.root.mainloop()

    def cancel_scan(self):
        self.cancel = True

    def read_code(self, image):
        for variant in (VARIANT_1_BOX, VARIANT_2_BOX):
            cropped = image.crop(variant)
            decoded = decode(cropped)
            if decoded:
                return decoded[0].data.decode('utf-8')
        return ""

    def scan_files(self):
        results = []
        for idx, f in enumerate(self.files, start=1):
            if self.cancel:
                break
            try:
                img = Image.open(f)
                code = self.read_code(img)
                results.append((os.path.basename(f), code))
                self.result_label.config(text=code)
            except Exception as e:
                results.append((os.path.basename(f), f"Error: {e}"))
            self.progress['value'] = idx * 100 / len(self.files)
            self.label.config(text=f"{idx} de {len(self.files)}")
            self.progress_win.update()

        if results and not self.cancel:
            self.save_results(results)
        else:
            self.progress_win.destroy()

    def get_default_excel_path(self):
        first_dir = os.path.dirname(self.files[0])
        recargas_dir = os.path.join(first_dir, 'recargas')
        os.makedirs(recargas_dir, exist_ok=True)
        base = os.path.basename(first_dir)
        excel_path = os.path.join(recargas_dir, base + '.xlsx')
        count = 1
        while os.path.exists(excel_path):
            excel_path = os.path.join(recargas_dir, f"{base}_{count}.xlsx")
            count += 1
        return excel_path

    def save_results(self, data):
        wb = Workbook()
        ws = wb.active
        ws.append(["Imagen", "Codigo"])
        for row in data:
            ws.append(row)

        excel_path = self.get_default_excel_path()
        try:
            wb.save(excel_path)
            self.open_file(excel_path)
        except Exception:
            # Ask user for alternative
            excel_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
            if excel_path:
                wb.save(excel_path)
                self.open_file(excel_path)
        self.progress_win.destroy()

    def open_file(self, path):
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', path])
            else:
                subprocess.call(['xdg-open', path])
        except Exception:
            messagebox.showwarning("Advertencia", "No se pudo abrir el archivo de Excel")


if __name__ == '__main__':
    ScannerApp()
