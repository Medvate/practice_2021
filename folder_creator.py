import os


MONTH_NAMES = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
               "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


def create_year_folder(year: int or str) -> None:
    """
    Создает в папке с данными папку для года и месяца в нем.

    :param year: Год, для которого нужно создать папку.
    :return: None.
    """
    cwd = os.getcwd()
    data_path = os.path.join(cwd, 'DATA')
    year_path = os.path.join(data_path, str(year))

    if not os.path.exists(data_path):
        os.mkdir(data_path)

    if not os.path.exists(year_path):
        os.mkdir(year_path)

    for month_counter in range(12):
        name_of_folder = f"{month_counter + 1} [{MONTH_NAMES[month_counter]}]"
        month_path = os.path.join(year_path, name_of_folder)

        if not os.path.exists(month_path):
            os.mkdir(month_path)


if __name__ == '__main__':
    YEARS = [2018, 2019, 2020, 2021]

    for year in YEARS:
        create_year_folder(year)

