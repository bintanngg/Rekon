import pandas as pd
import threading
import logging
import os
from tkinter import filedialog, messagebox
import tkinter as tk

# Constants
DEFAULT_FILE_TYPES = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
EXCEL_ENGINE = 'openpyxl'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
WARNING_NO_FILE = "Silahkan pilih file terlebih dahulu!"
WARNING_NO_COLUMNS = "Silahkan pilih semua kolom!"
ERROR_NON_NUMERIC = "Ada nilai non-numeric di kolom debet atau kredit!"
ERROR_EMPTY_FILE = "File Excel kosong atau tidak valid!"
ERROR_PARSE_FILE = "Format file Excel tidak dapat diparsing!"
ERROR_PROCESSING = "Gagal memproses rekonsiliasi:\n{error}"
SUCCESS_MESSAGE = "Sukses: {issues} transaksi bermasalah ditemukan"
CANCELLED_MESSAGE = "Cancelled"
NORMAL_TRX_SHEET = 'Normal Trx'
PROBLEM_TRX_SHEET = 'Problem Trx'
DIFFERENCE_SUMMARY_SHEET = 'Difference Summary'
DEFAULT_OUTPUT_FILENAME = "rekonsiliasi_output.xlsx"
RESULTS_TEXT_TEMPLATE = """
Rekonsiliasi Selesai!

Hasil:
  • Total Transaksi: {total_transaksi}
  • Jumlah Transaksi Normal: {jumlah_normal}
  • Jumlah Transaksi Bermasalah: {jumlah_bermasalah}
  • Total Selisih: Rp {total_selisih:,.2f}

File Output:
  • {output_file}
""".strip()

class RekonsiliasiLogic:
    """
    Logic class for handling reconciliation operations.
    Manages file selection, processing, and export of reconciliation results.
    """
    def __init__(self, ui):
        self.ui = ui
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(filename='rekonsiliasi.log', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

    def select_files(self):
        """
        Opens file dialog to select an Excel file and loads it into a dataframe.
        Handles large files by checking size and using chunking if necessary.
        """
        file_path = filedialog.askopenfilename(
            title="Pilih file Excel untuk Rekonsiliasi",
            filetypes=DEFAULT_FILE_TYPES
        )

        if file_path:
            # Validate file
            if not os.path.exists(file_path):
                messagebox.showerror("Error", f"File not found: {file_path}")
                return

            if not (file_path.endswith('.xlsx') or file_path.endswith('.xls')):
                messagebox.showerror("Error", "File must be in Excel format (.xlsx or .xls)")
                return

            # Check file size for chunking (10MB threshold)
            file_size = os.path.getsize(file_path)
            chunk_size = 10000 if file_size > 10 * 1024 * 1024 else None  # 10MB

            try:
                # Read file with chunking if large
                if chunk_size:
                    chunks = pd.read_excel(file_path, chunksize=chunk_size)
                    self.df = pd.concat(chunks, ignore_index=True)
                else:
                    self.df = pd.read_excel(file_path, engine=EXCEL_ENGINE)

                self.filename = os.path.basename(file_path)
                logging.info(f"File loaded: {file_path}, Size: {file_size} bytes")

                # Check if DataFrame is empty
                if self.df.empty:
                    messagebox.showerror("Error", ERROR_EMPTY_FILE)
                    self.ui.status_var.set("Error: Empty file")
                    return

            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file: {str(e)}")
                self.ui.status_var.set("Error reading file")
                logging.error(f"Error reading file {file_path}: {str(e)}")
                return

            # Update label
            self.ui.file_label.config(text=f"File loaded: {self.filename}", fg="#000000")

            # Update combobox dengan nama kolom dari file
            columns = list(self.df.columns)
            self.ui.combo_ket['values'] = columns
            self.ui.combo_debet['values'] = columns
            self.ui.combo_kredit['values'] = columns

            self.ui.status_var.set(f"File loaded: {len(self.df)} rows")

    def process_reconciliation(self):
        """
        Initiates reconciliation processing in a separate thread to keep UI responsive.
        """
        if not self.validate_inputs():
            return

        # Run in thread to avoid freezing UI
        threading.Thread(target=self._process_reconciliation_thread).start()

    def validate_inputs(self):
        """
        Validates that a file is loaded and all columns are selected.
        Returns True if valid, False otherwise.
        """
        if not hasattr(self, 'df'):
            messagebox.showwarning("Peringatan", WARNING_NO_FILE)
            return False

        kolom_ket = self.ui.kolom_ket.get()
        kolom_debet = self.ui.kolom_debet.get()
        kolom_kredit = self.ui.kolom_kredit.get()

        if not kolom_ket or not kolom_debet or not kolom_kredit:
            messagebox.showwarning("Peringatan", WARNING_NO_COLUMNS)
            return False

        # Additional validation: check if columns exist in dataframe
        if kolom_ket not in self.df.columns or kolom_debet not in self.df.columns or kolom_kredit not in self.df.columns:
            messagebox.showerror("Error", "Kolom yang dipilih tidak ada di file!")
            return False

        # Check if debit and credit columns are different
        if kolom_debet == kolom_kredit:
            messagebox.showerror("Error", "Kolom Debet dan Kredit harus berbeda!")
            return False

        return True

    def _process_reconciliation_thread(self):
        try:
            self.ui.status_var.set("Processing...")
            self.ui.progress_var.set(10)
            self.ui.root.update()

            kolom_ket = self.ui.kolom_ket.get()
            kolom_debet = self.ui.kolom_debet.get()
            kolom_kredit = self.ui.kolom_kredit.get()

            # Single file processing
            df_copy = self.df.copy()
            self.ui.progress_var.set(30)
            df_copy[kolom_debet] = pd.to_numeric(df_copy[kolom_debet], errors='coerce')
            df_copy[kolom_kredit] = pd.to_numeric(df_copy[kolom_kredit], errors='coerce')

            # Detect rows with non-numeric values in debit and credit columns
            debit_nan_rows = df_copy.index[df_copy[kolom_debet].isna()].tolist()
            credit_nan_rows = df_copy.index[df_copy[kolom_kredit].isna()].tolist()

            if debit_nan_rows or credit_nan_rows:
                error_message = "Nilai non-numeric ditemukan pada baris:\n"
                if debit_nan_rows:
                    error_message += f" - Kolom Debet: {', '.join(str(r + 2) for r in debit_nan_rows)}\n"
                    # +2 because dataframe index starts at 0 and Excel rows start at 1 with header row
                if credit_nan_rows:
                    error_message += f" - Kolom Kredit: {', '.join(str(r + 2) for r in credit_nan_rows)}\n"
                messagebox.showerror("Error", error_message)
                self.ui.status_var.set("Error: Non-numeric values")
                self.ui.progress_var.set(0)
                return

            self.ui.progress_var.set(50)

            # Proses rekonsiliasi
            rekap = df_copy.groupby(kolom_ket)[[kolom_debet, kolom_kredit]].sum()
            rekap['Selisih'] = rekap[kolom_debet] - rekap[kolom_kredit]

            keterangan_selisih = rekap[rekap['Selisih'] != 0].index.tolist()
            # Handle NaN in filtering for problem transactions
            mask_selisih = df_copy[kolom_ket].isin(keterangan_selisih)
            if any(pd.isna(k) for k in keterangan_selisih):
                mask_selisih = mask_selisih | df_copy[kolom_ket].isna()
            data_selisih = df_copy[mask_selisih]

            keterangan_normal = rekap[rekap['Selisih'] == 0].index.tolist()
            # Handle NaN in filtering for normal transactions
            mask_normal = df_copy[kolom_ket].isin(keterangan_normal)
            if any(pd.isna(k) for k in keterangan_normal):
                mask_normal = mask_normal | df_copy[kolom_ket].isna()
            data_normal = df_copy[mask_normal]

            total_selisih = rekap['Selisih'].sum()
            jumlah_bermasalah = len(data_selisih)
            jumlah_normal = len(data_normal)

            self.ui.progress_var.set(70)

            # Choose location and name for output file
            output_file = filedialog.asksaveasfilename(
                title="Save Reconciliation Output File",
                defaultextension=".xlsx",
                filetypes=DEFAULT_FILE_TYPES,
                initialfile=DEFAULT_OUTPUT_FILENAME
            )

            if not output_file:
                self.ui.status_var.set("Cancelled")
                self.ui.progress_var.set(0)
                return

            self.ui.progress_var.set(90)

            with pd.ExcelWriter(output_file, engine=EXCEL_ENGINE) as writer:
                data_normal.to_excel(writer, sheet_name=NORMAL_TRX_SHEET, index=False)
                data_selisih.to_excel(writer, sheet_name=PROBLEM_TRX_SHEET, index=False)
                rekap.loc[keterangan_selisih].to_excel(writer, sheet_name=DIFFERENCE_SUMMARY_SHEET)

            # Display results
            hasil_text = RESULTS_TEXT_TEMPLATE.format(
                total_transaksi=len(df_copy),
                jumlah_normal=jumlah_normal,
                jumlah_bermasalah=jumlah_bermasalah,
                total_selisih=total_selisih,
                output_file=output_file
            )

            self.ui.result_text.config(state=tk.NORMAL)
            self.ui.result_text.delete(1.0, tk.END)
            self.ui.result_text.insert(tk.END, hasil_text)
            self.ui.result_text.config(fg="#000000", state=tk.DISABLED)
            self.ui.status_var.set(SUCCESS_MESSAGE.format(issues=jumlah_bermasalah))
            self.ui.progress_var.set(100)

            logging.info(f"Reconciliation completed. Output: {output_file}, Issues: {jumlah_bermasalah}")

        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "File Excel kosong atau tidak valid!")
            self.ui.status_var.set("Error: Empty file")
            logging.error("Empty data error")
        except pd.errors.ParserError:
            messagebox.showerror("Error", "Format file Excel tidak dapat diparsing!")
            self.ui.status_var.set("Error: Parse error")
            logging.error("Parser error")
        except Exception as e:
            messagebox.showerror("Error", ERROR_PROCESSING.format(error=str(e)))
            self.ui.status_var.set("Error during processing")
            logging.error(f"Processing error: {str(e)}")
        finally:
            self.ui.progress_var.set(0)

    def export_results(self):
        """
        Exports the current results displayed in the UI to a text file.
        """
        if not hasattr(self.ui, 'result_text') or not self.ui.result_text.get(1.0, tk.END).strip():
            messagebox.showwarning("Peringatan", "Tidak ada hasil untuk diekspor!")
            return

        results_text = self.ui.result_text.get(1.0, tk.END).strip()
        output_file = filedialog.asksaveasfilename(
            title="Save Results to Text File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="reconciliation_results.txt"
        )

        if output_file:
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(results_text)
                messagebox.showinfo("Sukses", f"Hasil berhasil diekspor ke {output_file}")
                logging.info(f"Results exported to {output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal mengekspor hasil: {str(e)}")
                logging.error(f"Export error: {str(e)}")
