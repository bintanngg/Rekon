'''
APLIKASI GUI PYTHON UNTUK REKONSILIASI
AKUN HARUS MEMILIKI TRANS REFERENCE YANG SAMA UNTUK TIAP TRANSAKSINYA
'''

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from pathlib import Path

class RekonsiliasiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Rekonsiliasi V1.0 by Bintang")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        # Variable untuk menyimpan data
        self.datapath = tk.StringVar()
        self.df = None
        self.kolom_ket = tk.StringVar()
        self.kolom_debet = tk.StringVar()
        self.kolom_kredit = tk.StringVar()
        
        # Style
        self.root.configure(bg='#f0f0f0')
        style = ttk.Style()
        style.theme_use('clam')
        
        # Create main frame
        self.create_widgets()
    
    def create_widgets(self):
        # Title
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=60)
        title_frame.pack(fill=tk.X)
        
        title_label = tk.Label(
            title_frame, 
            text="🏦 Aplikasi Rekonsiliasi",
            font=("Arial", 16, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(pady=15)
        
        # Main content frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="1. Pilih File Excel", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(file_frame, text="📁 Pilih File Excel", command=self.select_file).pack(fill=tk.X, pady=5)
        
        self.file_label = tk.Label(
            file_frame,
            text="Belum ada file yang dipilih",
            font=("Arial", 9),
            fg="#7f8c8d",
            wraplength=500,
            justify=tk.LEFT
        )
        self.file_label.pack(fill=tk.X, pady=5)
        
        # 2. Column Selection Section
        column_frame = ttk.LabelFrame(main_frame, text="2. Pilih Kolom", padding="10")
        column_frame.pack(fill=tk.X, pady=10)
        
        # Kolom Keterangan
        ttk.Label(column_frame, text="Kolom Transaksi (Keterangan):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.combo_ket = ttk.Combobox(
            column_frame, 
            textvariable=self.kolom_ket, 
            state='readonly',
            width=40
        )
        self.combo_ket.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # Kolom Debet
        ttk.Label(column_frame, text="Kolom Debet:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.combo_debet = ttk.Combobox(
            column_frame, 
            textvariable=self.kolom_debet, 
            state='readonly',
            width=40
        )
        self.combo_debet.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        # Kolom Kredit
        ttk.Label(column_frame, text="Kolom Kredit:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.combo_kredit = ttk.Combobox(
            column_frame, 
            textvariable=self.kolom_kredit, 
            state='readonly',
            width=40
        )
        self.combo_kredit.grid(row=2, column=1, sticky=tk.W, padx=10)
        
        # 3. Output Section
        output_frame = ttk.LabelFrame(main_frame, text="3. Nama File Output", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(output_frame, text="Nama file output:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar(value="hasil_rekonsiliasi.xlsx")
        ttk.Entry(output_frame, textvariable=self.output_var, width=45).grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # 4. Results Section (akan ditampilkan setelah proses)
        self.result_frame = ttk.LabelFrame(main_frame, text="4. Hasil Rekonsiliasi", padding="10")
        self.result_frame.pack(fill=tk.X, pady=10)
        
        self.result_text = tk.Label(
            self.result_frame,
            text="Hasil akan ditampilkan di sini setelah memproses",
            font=("Arial", 9),
            fg="#7f8c8d",
            justify=tk.LEFT
        )
        self.result_text.pack(fill=tk.X)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(
            button_frame, 
            text="▶ Proses Rekonsiliasi", 
            command=self.process_reconciliation
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="🔄 Reset", 
            command=self.reset_form
        ).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Siap")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bg='#ecf0f1',
            fg='#34495e',
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=10,
            pady=5
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Pilih file Excel untuk Rekonsiliasi",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            # Validasi file
            if not os.path.exists(file_path):
                messagebox.showerror("Error", f"File tidak ditemukan: {file_path}")
                return
            
            if not (file_path.endswith('.xlsx') or file_path.endswith('.xls')):
                messagebox.showerror("Error", "File harus berformat Excel (.xlsx atau .xls)")
                return
            
            try:
                # Baca file
                self.df = pd.read_excel(file_path)
                self.datapath.set(file_path)
                
                # Update label
                filename = os.path.basename(file_path)
                self.file_label.config(text=f"✓ {filename}", fg="#27ae60")
                
                # Update combobox dengan nama kolom
                columns = list(self.df.columns)
                self.combo_ket['values'] = columns
                self.combo_debet['values'] = columns
                self.combo_kredit['values'] = columns
                
                self.status_var.set(f"File loaded: {filename} ({len(self.df)} rows)")
                messagebox.showinfo("Sukses", f"File berhasil dimuat!\n\nKolom tersedia:\n" + "\n".join([f"• {col}" for col in columns]))
                
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membaca file: {str(e)}")
                self.status_var.set("Error reading file")
    
    def process_reconciliation(self):
        # Validasi input
        if self.df is None:
            messagebox.showwarning("Peringatan", "Silahkan pilih file terlebih dahulu!")
            return
        
        if not self.kolom_ket.get() or not self.kolom_debet.get() or not self.kolom_kredit.get():
            messagebox.showwarning("Peringatan", "Silahkan pilih semua kolom!")
            return
        
        try:
            self.status_var.set("Memproses...")
            self.root.update()
            
            kolom_ket = self.kolom_ket.get()
            kolom_debet = self.kolom_debet.get()
            kolom_kredit = self.kolom_kredit.get()
            
            # Konversi ke numeric
            df_copy = self.df.copy()
            df_copy[kolom_debet] = pd.to_numeric(df_copy[kolom_debet], errors='coerce')
            df_copy[kolom_kredit] = pd.to_numeric(df_copy[kolom_kredit], errors='coerce')
            
            # Cek NaN
            if df_copy[kolom_debet].isna().any() or df_copy[kolom_kredit].isna().any():
                messagebox.showerror("Error", "Ada nilai non-numeric di kolom debet atau kredit!")
                self.status_var.set("Error: Non-numeric values")
                return
            
            # Proses rekonsiliasi
            rekap = df_copy.groupby(kolom_ket)[[kolom_debet, kolom_kredit]].sum()
            rekap['Selisih'] = rekap[kolom_debet] - rekap[kolom_kredit]
            
            # Cari transaksi bermasalah
            keterangan_selisih = rekap[rekap['Selisih'] != 0].index.tolist()
            data_selisih = df_copy[df_copy[kolom_ket].isin(keterangan_selisih)]
            
            # Hitung statistik
            total_selisih = rekap['Selisih'].sum()
            jumlah_bermasalah = len(keterangan_selisih)
            
            # Simpan file
            output_file = self.output_var.get()
            if not (output_file.endswith('.xlsx') or output_file.endswith('.xls')):
                output_file += '.xlsx'
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                data_selisih.to_excel(writer, sheet_name='Transaksi Bermasalah', index=False)
                rekap.loc[keterangan_selisih].to_excel(writer, sheet_name='Rekap Selisih')
            
            # Tampilkan hasil
            hasil_text = f"""
✓ Rekonsiliasi Selesai!

📊 Hasil:
  • Total Selisih: Rp {total_selisih:,.2f}
  • Jumlah Transaksi Bermasalah: {jumlah_bermasalah}
  • Total Transaksi: {len(df_copy)}

💾 File Output:
  • {os.path.abspath(output_file)}
            """.strip()
            
            self.result_text.config(text=hasil_text, fg="#27ae60", font=("Arial", 9, "normal"))
            self.status_var.set(f"Sukses: {jumlah_bermasalah} transaksi bermasalah ditemukan")
            
            messagebox.showinfo(
                "Sukses",
                f"Rekonsiliasi berhasil!\n\n"
                f"Total Selisih: Rp {total_selisih:,.2f}\n"
                f"Transaksi Bermasalah: {jumlah_bermasalah}\n\n"
                f"File disimpan: {output_file}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memproses rekonsiliasi:\n{str(e)}")
            self.status_var.set("Error during processing")
    
    def reset_form(self):
        self.datapath.set("")
        self.kolom_ket.set("")
        self.kolom_debet.set("")
        self.kolom_kredit.set("")
        self.output_var.set("hasil_rekonsiliasi.xlsx")
        self.file_label.config(text="Belum ada file yang dipilih", fg="#7f8c8d")
        self.result_text.config(text="Hasil akan ditampilkan di sini setelah memproses", fg="#7f8c8d")
        self.df = None
        self.combo_ket['values'] = []
        self.combo_debet['values'] = []
        self.combo_kredit['values'] = []
        self.status_var.set("Siap")


if __name__ == "__main__":
    root = tk.Tk()
    app = RekonsiliasiApp(root)
    root.mainloop()
