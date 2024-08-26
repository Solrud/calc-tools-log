import tkinter as tk
from tkinter import filedialog, messagebox
from constants import *

class GUI:
    def __init__(self, window, excel_handler):
        self.window = window
        self.excel_handler = excel_handler

        self.tools_log_file = TOOLS_LOG_PATH
        self.svod_excel_file = SVOD_EXCEL_PATH
        self.output_path = OUTPUT_PATH


        self.draw_widgets()
        self.define_window_options()

    def define_window_options(self):
        self.window.title('Обработчик ToolsLog/Svod')
        self.window.geometry('500x300')

        self.write_path_to(self.tools_log_file, self.file_tools_log_input)
        self.write_path_to(self.svod_excel_file, self.file_svod_excel_input)
        self.write_path_to(self.output_path, self.to_choose_ouput_path_input)

    def draw_widgets(self):
        # Выбор ToolsLog Excel
        self.to_choose_file_tools_log_btn = tk.Button(self.window, text='Выберите ToolsLog.xlsx файл',
                                                      command=self.select_file_tools_log_file)
        self.to_choose_file_tools_log_btn.grid(column=1, row=1, padx=20, pady=30)

        self.file_tools_log_input = tk.Text(self.window, width=30, height=1, state='disabled')
        self.file_tools_log_input.grid(column=2, row=1)

        # Выбор Svod Excel
        self.to_choose_file_svod_excel_btn = tk.Button(self.window, text='Выберите SvodExcel.xlsx файл',
                                                       command=self.select_file_svod_excel_file)
        self.to_choose_file_svod_excel_btn.grid(column=1, row=2)

        self.file_svod_excel_input = tk.Text(self.window, width=30, height=1, state='disabled')
        self.file_svod_excel_input.grid(column=2, row=2)

        # Выбор пути создания обработанного файла Excel
        self.to_choose_ouput_path_btn = tk.Button(self.window, text='Выберите путь вывода',
                                                  command=self.select_output_path)
        self.to_choose_ouput_path_btn.grid(column=1, row=3, pady=30)

        self.to_choose_ouput_path_input = tk.Text(self.window, width=30, height=1, state='disabled')
        self.to_choose_ouput_path_input.grid(column=2, row=3)

        # Кнопка обработки файлов
        self.to_handle_files_btn = tk.Button(self.window, text='Обработать', bg='#b7fa84', font=('Arial Bold', 13),
                                             command=self.define_command_handler_files_btn)
        self.to_handle_files_btn.grid(column=2, row=4, pady=40)

    def select_file_tools_log_file(self): # Метод выбора ToolsLog.xlsx
        self.tools_log_file = filedialog.askopenfilename(title='Выберите ToolsLog.xlsx',
                                                         filetypes=[("Excel файлы", "*.xlsx *.xls")])
        self.write_path_to(self.tools_log_file, self.file_tools_log_input)

    def select_file_svod_excel_file(self): # Метод выбора SvodExcel.xlsx
        self.svod_excel_file = filedialog.askopenfilename(title='Выберите SvodExcel.xlsx',
                                                         filetypes=[("Excel файлы", "*.xlsx *.xls")])
        self.write_path_to(self.svod_excel_file, self.file_svod_excel_input)


    def select_output_path(self):
        self.output_path = filedialog.askdirectory(title='Выберите путь вывода нового файла')
        self.write_path_to(self.output_path, self.to_choose_ouput_path_input)

    def write_path_to(self, file_path, text_file_path):
        if (file_path):
            text_file_path.configure(state='normal')
            text_file_path.delete(1.0, tk.END)
            text_file_path.insert("end", file_path)
            text_file_path.configure(state='disabled')





    def define_command_handler_files_btn(self):
        if (self.tools_log_file and self.svod_excel_file and self.output_path):
            # messagebox.showinfo('Успех!','Программа приступила к выполнению')
            self.excel_handler.handler(self.tools_log_file, self.svod_excel_file, self.output_path)
        else:
            messagebox.showerror('Ошибка!','Одно из полей не заполнено')