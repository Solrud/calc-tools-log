from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from datetime import datetime
from utils import check_file_exists
from gui import GUI
from constants import OUTPUT_FILE_TYPE
class ExcelHandler:
    def handler(self, tools_log_path, svod_excel_path, output_path, file_name):
        try:
            st = datetime.now()
            wb_new = Workbook()
            ws_new = wb_new.active
            ws_new.title = 'Данные обработки'
            ws_new.append(['UID', 'Раздел', 'Тип_инструмента', 'Обозначение', 'Наименование', 'Производитель',
             'Доп._параметры', 'Остаток_ИРК_20_1', 'Остаток_ИРК_21_1', 'Остаток_ИРК_22_1', 'Остаток_ИРК_23_1',
             'Остаток_ИРК_20_2', 'Остаток_ИРК_23_2', 'Выдано раз', 'Выдано количество суммарно',
             'Дата_посл.выдачи', 'Приход раз', 'Приход количество суммарно', 'Дата_посл.прихода', 'Ячейка', 'Стелаж'])
            fill_svod_color = PatternFill(start_color='69fff5', end_color='69fff5', fill_type='solid')
            fill_tools_log_color = PatternFill(start_color='a8cae4', end_color='a8cae4', fill_type='solid')
            for cell in ws_new['1:1'[0]]:
                if (cell.value in ['Раздел', 'Тип_инструмента', 'Выдано раз', 'Выдано количество суммарно', 'Дата_посл.выдачи',
                             'Приход раз', 'Приход количество суммарно', 'Дата_посл.прихода']):
                    cell.fill = fill_tools_log_color
                else:
                    cell.fill = fill_svod_color
                cell.alignment = Alignment(vertical='center')
                cell.font = Font(bold=True)
                cell.border = Border( left=Side(style='medium'),
                                      right=Side(style='medium'),
                                      top=Side(style='medium'),
                                      bottom=Side(style='medium'))
            ws_new.row_dimensions[1].height = 45


            wb_svod = load_workbook(svod_excel_path)
            ws_svod = wb_svod['Лист1']
            ws_svod_tuple = tuple(ws_svod.iter_rows(values_only=True))
            ws_svod_dict_uid = {}
            for tup in ws_svod_tuple:
                uid_dict = tup[0]
                ws_svod_dict_uid[uid_dict] = tup

            wb_tools_log = load_workbook(tools_log_path)
            ws_tools_log = wb_tools_log['Лист1']

            iter = 0

            lst_uid = -1
            biggest_date_vidacha = '01.01.1900'
            biggest_date_prihod = '01.01.1900'
            sum_vidano = 0
            count_vidano = 0
            sum_prihod = 0
            count_prihod = 0
            razdel = ''
            type_instrument = ''
            for row in ws_tools_log.rows:
                iter += 1

                current_uid = row[0].value
                current_date = row[1].value
                current_vidano = row[3].value
                current_prihod = row[6].value
                current_razdel = row[9].value
                current_type_instrument = row[10].value

                if row[0].value != None and type(row[0].value) is int:

                    if current_uid == lst_uid:
                        pass
                    else:
                        if lst_uid != -1:
                            if biggest_date_vidacha == '01.01.1900':
                                biggest_date_vidacha = ''
                            if biggest_date_prihod == '01.01.1900':
                                biggest_date_prihod = ''

                            obozn = ''
                            naim = ''
                            proizv = ''
                            dop_param = ''
                            ost_irk_20_1 = ''
                            ost_irk_21_1 = ''
                            ost_irk_22_1 = ''
                            ost_irk_23_1 = ''
                            ost_irk_20_2 = ''
                            ost_irk_23_2 = ''
                            stelazh = ''
                            yacheika = ''

                            row_svod_by_uid = ws_svod_dict_uid.get(lst_uid)
                            if row_svod_by_uid:
                                obozn = row_svod_by_uid[1]
                                naim = row_svod_by_uid[2]
                                proizv = row_svod_by_uid[3]
                                dop_param = row_svod_by_uid[4]
                                ost_irk_20_1 = row_svod_by_uid[6]
                                ost_irk_21_1 = row_svod_by_uid[8]
                                ost_irk_22_1 = row_svod_by_uid[10]
                                ost_irk_23_1 = row_svod_by_uid[12]
                                ost_irk_20_2 = row_svod_by_uid[14]
                                ost_irk_23_2 = row_svod_by_uid[16]
                                stelazh = row_svod_by_uid[18]
                                yacheika = row_svod_by_uid[19]

                            push_row = [lst_uid, razdel, type_instrument, obozn, naim, proizv, dop_param,
                                        ost_irk_20_1, ost_irk_21_1, ost_irk_22_1, ost_irk_23_1, ost_irk_20_2,
                                        ost_irk_23_2, count_vidano, sum_vidano, biggest_date_vidacha, count_prihod,
                                        sum_prihod, biggest_date_prihod, stelazh, yacheika]
                            ws_new.append(push_row)

                        lst_uid = current_uid
                        biggest_date_vidacha = '01.01.1900'
                        biggest_date_prihod = '01.01.1900'
                        sum_vidano = 0
                        count_vidano = 0
                        sum_prihod = 0
                        count_prihod = 0
                        razdel = ''
                        type_instrument = ''

                    razdel = current_razdel
                    type_instrument = current_type_instrument

                    if current_vidano > 0:
                        sum_vidano += current_vidano
                        count_vidano += 1
                        if datetime.strptime(current_date, '%d.%m.%Y') > datetime.strptime(biggest_date_vidacha, '%d.%m.%Y'):
                            biggest_date_vidacha = current_date

                    if current_prihod > 0:
                        sum_prihod += current_prihod
                        count_prihod += 1
                        if datetime.strptime(current_date, '%d.%m.%Y') > datetime.strptime(biggest_date_prihod, '%d.%m.%Y'):
                            biggest_date_prihod = current_date
                else:
                    continue

            for row in ws_new.rows:
                for cell in row:
                    if(row != row[1]):
                        ws_new.column_dimensions[cell.column_letter].auto_size = True
                        cell.border = Border(left=Side(style='thin'),
                                         right=Side(style='thin'),
                                         top=Side(style='thin'),
                                         bottom=Side(style='thin'))
            ws_new.column_dimensions['B'].width = 30
            ws_new.column_dimensions['E'].width = 50

            output_full_path = output_path + file_name + OUTPUT_FILE_TYPE
            wb_new.save(output_full_path)
            if (check_file_exists(output_full_path)):
                fn = datetime.now()
                code_work_time = (fn-st)
                GUI.show_info_messagebox('Успех', f'Файл создан по пути: {output_full_path}. ' +
                                         f'\nВремя выполнения: {code_work_time}')
            else:
                GUI.show_error_messagebox('Ошибка', 'Файл не создался по выбранному пути')
        except Exception as e:
            print(e)
            GUI.show_error_messagebox('Ошибка!', 'Во время загрузки и обработки файлов произошла ошибка')


        #163871 uid который повторяется по датам 2 раза, но с разным типом инструмента
        # нужно брать тот который в svode

        # Время загрузки wb_svod =  0:00:17.699499
        # Время итераци ws_svod 0:00:01.460051
        # Время загрузки wb_tools_log = 0:00:24.833626