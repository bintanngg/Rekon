import tkinter as tk
from tkinter import ttk
import os

class RekonsiliasiUI:
    def __init__(self, root, logic):
        self.root = root
        self.logic = logic
        self.root.title("Rekon v1.1 | App by Bintanngg")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.image_path = os.path.join(self.script_dir, 'cat.png')
        self.app_icon = tk.PhotoImage(file=self.image_path)
        self.root.iconphoto(False, self.app_icon)
        
        # Variables
        self.kolom_ket = tk.StringVar()
        self.kolom_debet = tk.StringVar()
        self.kolom_kredit = tk.StringVar()
        self.progress_var = tk.IntVar()

        # Style
        self.root.configure(bg='#f0f0f0')
        style = ttk.Style()
        style.theme_use('clam')

        # Create widgets
        self.create_widgets()

    def create_widgets(self):
        self.create_title()
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        self.create_file_section(main_frame)
        self.create_column_section(main_frame)
        self.create_results_section(main_frame)
        self.create_buttons(main_frame)
        self.create_status_bar()

    def create_title(self):
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=60)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="Aplikasi Rekonsiliasi", font=("Arial", 16, "bold"), bg='#2c3e50', fg='white').pack(pady=15)

    def create_file_section(self, parent):
        file_frame = ttk.LabelFrame(parent, text="1. Pilih File Excel", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        ttk.Button(file_frame, text="Pilih File Excel", command=self.logic.select_files).pack(fill=tk.X, pady=5)
        self.file_label = tk.Label(file_frame, text="Belum ada file yang dipilih", font=("Arial", 9), fg="#7f8c8d", wraplength=500, justify=tk.LEFT)
        self.file_label.pack(fill=tk.X, pady=5)

    def create_column_section(self, parent):
        column_frame = ttk.LabelFrame(parent, text="2. Pilih Kolom", padding="10")
        column_frame.pack(fill=tk.X, pady=10)
        column_frame.columnconfigure(1, weight=1)
        labels = ["Kolom Transaksi (Keterangan):", "Kolom Debet:", "Kolom Kredit:"]
        vars = [self.kolom_ket, self.kolom_debet, self.kolom_kredit]
        combos = []
        for i, (label, var) in enumerate(zip(labels, vars)):
            ttk.Label(column_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            combo = ttk.Combobox(column_frame, textvariable=var, state='readonly')
            combo.grid(row=i, column=1, sticky=tk.EW, padx=10)
            combos.append(combo)
        self.combo_ket, self.combo_debet, self.combo_kredit = combos

    def create_results_section(self, parent):
        self.result_frame = ttk.LabelFrame(parent, text="3. Hasil Rekonsiliasi", padding="10")
        self.result_frame.pack(fill=tk.X, pady=10)
        self.result_text = tk.Text(self.result_frame, height=8, font=("Arial", 9), fg="#7f8c8d", wrap=tk.WORD, state=tk.DISABLED)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self.result_text.insert(tk.END, "Hasil akan ditampilkan di sini setelah memproses")

    def create_buttons(self, parent):
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="Proses Rekonsiliasi", command=self.logic.process_reconciliation).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset", command=self.reset_form).pack(side=tk.LEFT, padx=5)

    def create_status_bar(self):
        self.status_var = tk.StringVar(value="Siap")
        status_frame = tk.Frame(self.root, bg='#ecf0f1', bd=1, relief=tk.SUNKEN)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = tk.Label(status_frame, textvariable=self.status_var, bg='#ecf0f1', fg='#34495e', anchor=tk.W, padx=10, pady=5)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.RIGHT, padx=10, pady=5)

    def reset_form(self):
        self.kolom_ket.set("")
        self.kolom_debet.set("")
        self.kolom_kredit.set("")
        self.file_label.config(text="Belum ada file yang dipilih", fg="#7f8c8d")
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Hasil akan ditampilkan di sini setelah memproses")
        self.result_text.config(fg="#7f8c8d", state=tk.DISABLED)
        self.combo_ket['values'] = []
        self.combo_debet['values'] = []
        self.combo_kredit['values'] = []
        self.status_var.set("Siap")
        self.progress_var.set(0)
