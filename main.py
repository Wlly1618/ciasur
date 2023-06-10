import openpyxl
import csv
import calendar


def get_data_fo_f2(worksheet, max_row):
    data_fo_f2 = []
    for row in worksheet.iter_rows(min_row=10, max_row=max_row + 9, min_col=2, max_col=25, values_only=True):
        data_fo_f2.append(row)

    return data_fo_f2


def get_data_times():
    data_times = []
    for i in range(24):
        data_times.append(i)

    return data_times


def write_data(file_out, header, year, month, data_times, data_fo_f2, data_fo_f2_correct):
    with open(file_out, 'w', newline='') as archivo_csv:
        writer = csv.writer(archivo_csv)
        writer.writerow(header)
        i = 1
        for row_fo_f2, row_fo_f2_c in zip(data_fo_f2, data_fo_f2_correct):
            for fo_f2, fo_f2_c, time in zip(row_fo_f2, row_fo_f2_c, data_times):
                # print("day: ", i, "\t time: ", time, "\tfo_f2: ", fo_f2, "\tfo_f2_correct: ", fo_f2_c)
                writer.writerow([year, month, i, time, fo_f2, fo_f2_c])
            i += 1

        archivo_csv.close()

    print("file created :)")


def check_data(col_data_fo_f2, col_data_fo_f2_correct):
    for r1, r2 in zip(col_data_fo_f2, col_data_fo_f2_correct):
        if r1 != r2:
            return False

    return True


def main():
    workbook = openpyxl.load_workbook('./file.xlsx')
    sheet = workbook.sheetnames[0]
    worksheet = workbook[sheet]

    data_times = get_data_times()
    month = worksheet['F6'].value
    year = worksheet['H6'].value
    amo_days = calendar.monthrange(year=year, month=month)
    data_fo_f2 = get_data_fo_f2(worksheet, amo_days[1])
    col_data_fo_f2 = []
    for row in worksheet.iter_rows(min_row=10, max_row=amo_days[1] + 9, min_col=1, max_col=1, values_only=True):
        col_data_fo_f2.append(row)

    sheet = workbook.sheetnames[2]
    worksheet = workbook[sheet]
    data_fo_f2_correct = get_data_fo_f2(worksheet, amo_days[1])
    col_data_fo_f2_correct = []
    for row in worksheet.iter_rows(min_row=10, max_row=amo_days[1] + 9, min_col=1, max_col=1, values_only=True):
        col_data_fo_f2_correct.append(row)

    header = ['year', 'month', 'day', 'time', 'foF2', 'foF2_']
    file = "output.csv"

    flag = check_data(col_data_fo_f2, col_data_fo_f2_correct)

    if flag == 1:
        write_data(file, header, year, month, data_times, data_fo_f2, data_fo_f2_correct)
    else:
        print("data bad")

    workbook.close()


if __name__ == '__main__':
    main()
