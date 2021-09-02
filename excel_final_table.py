import os
import xlsxwriter


NAME_OF_XLSX = 'final_table.xlsx'


class DateFilter:
    def __init__(self, start_date: str, stop_date):
        self._start_date = start_date
        self._stop_date = stop_date

        self._start_month = int(start_date.split('.')[0])
        self._start_year = int(start_date.split('.')[1])
        self._stop_month = int(stop_date.split('.')[0])
        self._stop_year = int(stop_date.split('.')[1])

    def get_titles_for_excel(self) -> list:
        titles = list()

        for year in range(self._start_year, self._stop_year + 1):
            if year == self._start_year:
                start_month = self._start_month
            else:
                start_month = 1

            if year == self._stop_year:
                stop_month = self._stop_month
            else:
                stop_month = 12

            for month in range(start_month, stop_month + 1):
                new_title = f"{month}.{year}"
                if len(new_title) == 6:
                    new_title = "0" + new_title
                titles.append(new_title)

        return titles

    def get_data_paths(self):
        pass


if __name__ == '__main__':
    START = '05.2021'
    STOP = '08.2021'
    date_filter = DateFilter(START, STOP)

    HEADERS = [{'header': 'НИОКР'}, *[{'header': date} for date in date_filter.get_titles_for_excel()]]

    if os.path.exists(NAME_OF_XLSX):
        os.remove(NAME_OF_XLSX)
    workbook = xlsxwriter.Workbook(NAME_OF_XLSX)
    worksheet = workbook.add_worksheet()
    BORDER_COLOR = '#ffffff'
    first_header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#6282db', 'border': 1, 'font_size': 18, 'border_color': BORDER_COLOR})
    main_header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#6282db', 'border': 1, 'font_size': 15, 'border_color': BORDER_COLOR})
    # header_format = workbook.add_format({'bold': True,
    #                                      'align': 'center',
    #                                      'valign': 'vcenter',
    #                                      'fg_color': '#D7E4BC',
    #                                      'border': 1})
    first_ordinary_format = workbook.add_format({'fg_color': '#6282db', 'border': 1, 'font_size': 14, 'border_color': BORDER_COLOR})
    plus_ordinary_format = workbook.add_format({'align': 'center', 'border': 1, 'font_size': 14, 'fg_color': '#1e9a33', 'border_color': BORDER_COLOR})
    minus_ordinary_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'font_size': 14, 'fg_color': '#ff4a4a', 'border_color': BORDER_COLOR})

    # center_format = workbook.add_format({'align': 'center'})
    worksheet.set_row(0, 25)
    NUM_OF_PROJECTS = 6
    for i in range(NUM_OF_PROJECTS):
        worksheet.set_row(i + 1, 18)
    # worksheet.set_cols(10)
    worksheet.set_column(0, 0, 45)
    NUM_OF_DATES = 4
    worksheet.set_column(1, NUM_OF_DATES, 10)
    DATA = [['НИОКР', *date_filter.get_titles_for_excel()],
            ['Телевизор', '+', '+', '—', '—'], ['Машина', '—', '—', '+', '+'],
            ['BigBabyTape', '—', '—', '+', '—'], ['Лампа', '+', '+', '+', '—'],
            ['Сосна', '—', '+', '—', '+'], ['Тополь', '+', '—', '—', '—']]
    # worksheet.add_table('A1:E7', {'data': DATA,
    #                               'header_row': False})
    # worksheet.add_table('A1:E7')
    for row in range(7):
        for col in range(5):
            if row == 0:
                if col == 0:
                    worksheet.write(row, col, DATA[row][col], first_header_format)
                else:
                    worksheet.write(row, col, DATA[row][col], main_header_format)
            else:
                if col == 0:
                    worksheet.write(row, col, DATA[row][col], first_ordinary_format)
                else:
                    if DATA[row][col] == '+':
                        worksheet.write(row, col, DATA[row][col], plus_ordinary_format)
                    else:
                        worksheet.write(row, col, DATA[row][col], minus_ordinary_format)
    workbook.close()