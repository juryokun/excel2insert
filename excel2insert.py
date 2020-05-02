import openpyxl
import json
import sys


def main():
    # load excel
    with open('settings.json', 'r') as f:
        try:
            settings = json.load(f)

            book = openpyxl.load_workbook(settings['dataFile'])
            sheet = book[settings['dataSheet']]
            output = settings['outputFile']

        except Exception as e:
            print("settings.jsonを読み込めませんでした。形式を見直してください。")
            print(e)
            sys.exit(1)

    # get table name
    table_name = get_table_name(sheet)

    # get data_types
    data_types = get_data_types(sheet)

    # set header
    write_header(output, table_name, data_types)

    # set body
    write_body(sheet, output, data_types)

    # set footer
    write_footer(output)


def get_table_name(sheet):
    return get_excel_value(sheet, 1, 2)


def get_data_types(sheet):
    data_types = {}
    excel_column = 1
    while is_exist_value(sheet, 4, excel_column):
        column_name = get_excel_value(sheet, 4, excel_column)
        data_type = get_excel_value(sheet, 3, excel_column)
        if data_type == None or data_type == '':
            data_type = 'var'

        data_types[column_name] = data_type

        excel_column += 1

    return data_types


def get_excel_value(sheet, row, column):
    return sheet.cell(row=row, column=column).value


def is_exist_value(sheet, row, column):
    value = get_excel_value(sheet, row, column)
    if value != None and value != '':
        return True
    return False


def write_header(output, table_name, data_types):
    value = "INSERT INTO " + table_name + "("
    for index, column_name in enumerate(data_types.keys()):
        value = value + column_name
        if index != len(data_types)-1:
            value = value + ","
    value = value + ") values\n"

    with open(output, mode='w') as f:
        f.write(value)


def write_footer(output):
    with open(output, mode='a') as f:
        f.write("\n;")


def write_body(sheet, output, data_types):
    row = 5
    while is_exist_value(sheet, row, 1):
        if row != 5:
            dataset = ",\n    ("
        else:
            dataset = "    ("

        for index, value in enumerate(data_types.values()):
            if index != 0:
                dataset = dataset + ","
            data = get_excel_value(sheet, row, index+1)
            if data == "null":
                dataset = dataset + data
                continue
            if data == None or data == "":
                dataset = dataset + "\'\'"
                continue
            if value == "var":
                dataset = dataset + "\'" + str(data) + "\'"
                continue
            dataset = dataset + str(data)
        dataset = dataset + ")"
        with open(output, mode='a') as f:
            f.write(dataset)

        row += 1


main()
