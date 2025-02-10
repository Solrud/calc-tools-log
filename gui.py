import os
import time
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
        self.file_name = OUTPUT_FILE_NAME

        self.draw_widgets()
        self.define_window_options()

    @staticmethod
    def show_info_messagebox(label, text):
        messagebox.showinfo(label, text)

    @staticmethod
    def show_error_messagebox(label, text):
        messagebox.showerror(label, text)

    @staticmethod
    def show_warning_messagebox(label, text):
        messagebox.showwarning(label, text)

    def define_window_options(self):
        self.window.title('Обработчик ToolsLog/Svod')
        self.window.geometry('480x385')

        self.write_path_in_input_text(self.tools_log_file, self.file_tools_log_input)
        self.write_path_in_input_text(self.svod_excel_file, self.file_svod_excel_input)
        self.write_path_in_input_text(self.output_path, self.to_choose_ouput_path_input)
        self.write_path_in_input_text(self.file_name, self.file_name_txt, False)

    def draw_widgets(self):
        # Выбор ToolsLog Excel
        self.to_choose_file_tools_log_btn = tk.Button(self.window, text='Выберите ToolsLog.xlsx файл',
                                                      command=self.select_file_tools_log_file)
        self.to_choose_file_tools_log_btn.grid(column=1, row=1, padx=20, pady=30)

        self.file_tools_log_input = tk.Entry(self.window, width=40, state='readonly', fg='#00540e')
        self.file_tools_log_input.grid(column=2, row=1)

        # Выбор Svod Excel
        self.to_choose_file_svod_excel_btn = tk.Button(self.window, text='Выберите SvodExcel.xlsx файл',
                                                       command=self.select_file_svod_excel_file)
        self.to_choose_file_svod_excel_btn.grid(column=1, row=2)

        self.file_svod_excel_input = tk.Entry(self.window, width=40, state='disabled', fg='#00540e')
        self.file_svod_excel_input.grid(column=2, row=2)

        # Выбор пути создания обработанного файла Excel
        self.to_choose_ouput_path_btn = tk.Button(self.window, text='Выберите путь вывода',
                                                  command=self.select_output_path)
        self.to_choose_ouput_path_btn.grid(column=1, row=3, pady=30)

        self.to_choose_ouput_path_input = tk.Entry(self.window, width=40, state='disabled', fg='#00540e')
        self.to_choose_ouput_path_input.grid(column=2, row=3)

        # Название файла
        self.file_name_lbl = tk.Label(self.window, text='Название файла:')
        self.file_name_lbl.grid(column=1, row=4)

        self.file_name_txt = tk.Entry(self.window, width=40)
        self.file_name_txt.grid(column=2, row=4)

        # Кнопка обработки файлов
        self.to_handle_files_btn = tk.Button(self.window, text='Обработать', bg='#b7fa84', font=('Arial Bold', 13),
                                             command=self.define_command_handler_files_btn)
        self.to_handle_files_btn.grid(column=2, row=5, pady=40)

        # Предупреждение о загрузке
        self.wait_info_txt = tk.Label(self.window, text='', fg='#780000')
        self.wait_info_txt.grid(column=1, row=5)

        # Версия ПО
        self.version_lbl = tk.Label(self.window, text=f'Версия v.{APP_VERSION} от {APP_VERSION_DATE}', fg='#780000')
        self.version_lbl.grid(column=1, row=6)

    def select_file_tools_log_file(self):  # Метод выбора ToolsLog.xlsx
        init_dir = os.path.dirname(self.tools_log_file)

        selected_file = filedialog.askopenfilename(
            title='Выберите ToolsLog.xlsx',
            filetypes=[("Excel файлы", "*.xlsx *.xls")],
            initialdir=init_dir
        )
        if selected_file:
            self.tools_log_file = selected_file
            self.write_path_in_input_text(self.tools_log_file, self.file_tools_log_input)

    def select_file_svod_excel_file(self):  # Метод выбора SvodExcel.xlsx
        init_dir = os.path.dirname(self.svod_excel_file)

        selected_file = filedialog.askopenfilename(
            title='Выберите SvodExcel.xlsx',
            filetypes=[("Excel файлы", "*.xlsx *.xls")],
            initialdir=init_dir
        )
        if selected_file:
            self.svod_excel_file = selected_file
            self.write_path_in_input_text(self.svod_excel_file, self.file_svod_excel_input)

    def select_output_path(self):
        selected_path = filedialog.askdirectory(title='Выберите путь вывода нового файла')
        if selected_path:
            self.output_path = selected_path
            self.output_path = self.output_path + '/' if self.output_path and self.output_path[-1] != '/' else self.output_path
            self.write_path_in_input_text(self.output_path, self.to_choose_ouput_path_input)

    def write_path_in_input_text(self, file_path, text_file_path, disable=True):
        text_file_path.configure(state='normal')  # Разблокируем поле для редактирования

        # Проверяем тип виджета
        if isinstance(text_file_path, tk.Text):  # Если это tk.Text
            text_file_path.delete(1.0, tk.END)   # Используем индексы для tk.Text
            text_file_path.insert("end", file_path)
        else:  # Если это tk.Entry
            text_file_path.delete(0, tk.END)     # Используем индексы для tk.Entry
            text_file_path.insert(0, file_path)

        if disable:
            text_file_path.configure(state='readonly')  # Блокируем поле обратно

    def define_command_handler_files_btn(self):
        file_name_txt = self.file_name_txt.get().strip()
        if self.tools_log_file.strip() and self.svod_excel_file.strip() and self.output_path.strip() and file_name_txt.strip():
            self.wait_info_txt.configure(text='Идёт обработка, подождите... \nПримерно 1мин. 22 сек.')
            self.to_handle_files_btn.configure(bg='#ffc4c4', state='disabled')
            self.wait_info_txt.update()
            self.to_handle_files_btn.update()

            self.excel_handler.init_values(self.tools_log_file, self.svod_excel_file, self.output_path, file_name_txt)
            self.excel_handler.handler()

            self.to_handle_files_btn.configure(bg="#b7fa84", state='normal')
            self.wait_info_txt.configure(text='')
        else:
            messagebox.showerror('Ошибка!', 'Одно из полей не заполнено')