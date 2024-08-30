import tkinter as tk
from gui import GUI
from excel_handler import ExcelHandler

def main():
    root_window = tk.Tk()
    excel_handler = ExcelHandler()
    GUI(root_window, excel_handler)

    root_window.mainloop()


if __name__ == '__main__':
    main()

'''
    Скомпилировать в .exe
    auto-py-to-exe
'''