import openpyxl
import csv


def get_data_fo_f2(worksheet):
    data_fo_f2 = []
    for row in worksheet.iter_rows(min_row=10, max_row=40, min_col=2, max_col=25, values_only=True):
        data_fo_f2.append(row)

    return data_fo_f2


def get_data_days():
    data_dates = []
    for i in range(32):
        data_dates.append(i + 1)

    return data_dates


def get_data_times():
    data_times = []
    for i in range(24):
        data_times.append(i)

    return data_times


def write_data(file_out, header, year, month, data_days, data_times, data_fo_f2, data_fo_f2_correct):
    with open(file_out, 'w', newline='') as archivo_csv:
        writer = csv.writer(archivo_csv)
        writer.writerow(header)
        for row_fo_f2, row_fo_f2_c, day in zip(data_fo_f2, data_fo_f2_correct, data_days):
            for fo_f2, fo_f2_c, time in zip(row_fo_f2, row_fo_f2_c, data_times):
                print("day: ", day, "\ttime: ", time, "\tfo_f2: ", fo_f2, "\tfo_f2_correct: ", fo_f2_c)
                writer.writerow([year, month, day, time, fo_f2, fo_f2_c])

        archivo_csv.close()

    print("file created :)")


def main():
    workbook = openpyxl.load_workbook('./file.xlsx')
    sheet = workbook.sheetnames[0]
    worksheet = workbook[sheet]
    data_fo_f2 = get_data_fo_f2(worksheet)

    sheet = workbook.sheetnames[2]
    worksheet = workbook[sheet]
    data_fo_f2_correct = get_data_fo_f2(worksheet)

    data_times = get_data_times()
    data_days = get_data_days()
    month = worksheet['F6'].value
    year = worksheet['H6'].value
    header = ['year', 'month', 'day', 'time', 'foF2', 'foF2_']

    file = "output.csv"

    write_data(file, header, year, month, data_days, data_times, data_fo_f2, data_fo_f2_correct)

    workbook.close()


if __name__ == '__main__':
    main()
