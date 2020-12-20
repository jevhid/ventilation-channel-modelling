import time

import sub_methods as sm


def main():
    path = "C:/Users/TurboNotik-PC/Dropbox/AAU/2_semester/Semesterprojekt/py/"
    file_template = path + "Skabelon.xlsx"

    file_name = input('enter file path [leave empty for default path]: ')
    if len(file_name) == 0:
        file_name = path + "Data.xlsx"
    print('data file selected', file_name)

    sheet_index = int(input("enter sheet number [leave empty for default value]: ") or 1) - 1
    print('sheet selected: ', sheet_index + 1)

    df = sm.normalize_data(file_name, sheet_index)
    sheet = sm.purge_data(file_template)

    column_name_list = ["Area", "Number", "Rumnavne", "Specified Supply Airflow",
                        "Specified Return Airflow", "Room: Department", "Rumnavne"]
    column_name_list = list(str(input("enter column name (space separated)[leave empty for default value]: ")).split()
                            or column_name_list)
    start = time.time()
    print('columns selected ', column_name_list)
    print('script running')

    for column_name in column_name_list:
        start_column = sm.get_column_index(column_name, sheet)
        sm.append_to_excel(start_column, column_name, file_template, df)

    print('success')
    end = time.time()
    print('time elapsed: ', round((end - start), 2), 'seconds')
    print('')


if __name__ == "__main__":
    main()

k = input("press enter to exit")
