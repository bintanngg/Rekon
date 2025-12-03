'''
!!ATTENTION!!
1. Data must be in EXCEL FORMAT!
2. Data must have the same trans reference (naming) for each transaction.
3. Data must have Debit and Credit columns.
4. How the application works
    a. Use the Naming column as a reference for debit-credit
    b. If debit-credit is not equal to 0 then it is considered a problematic transaction.
'''

from ui import RekonsiliasiUI
from logic import RekonsiliasiLogic
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    logic = RekonsiliasiLogic(None)  # placeholder
    ui = RekonsiliasiUI(root, logic)
    logic.ui = ui
    root.mainloop()
