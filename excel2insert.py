import openpyxl
import json
import sys


def main():
    """
    メイン処理
    """
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

    try:
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
    except Exception as e:
        print(e)
        sys.exit(1)

    sys.exit()


def get_table_name(sheet):
    """
    テーブル名取得

    Returns
    ----------
    str or int or None
    """
    return get_excel_value(sheet, 1, 2)


def get_data_types(sheet):
    """
    各カラムの型を取得

    Parameters
    ----------
    sheet : object
        EXCELシートオブジェクト

    Returns
    ----------
    data_types : dict => {column_name: data_type}
    """
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
    """
    該当セルの値を取得する

    Parameters
    ----------
    sheet : object
        EXCELシートオブジェクト
    row : int
        行番号
    column : int
        列番号

    Returns
    ---------
    str or int or None
    """
    return sheet.cell(row=row, column=column).value


def is_exist_value(sheet, row, column):
    """
    該当セルに値があるかをチェックする

    Returns
    ----------
    bool
        値があれば True、なければ False
    """
    value = get_excel_value(sheet, row, column)
    if value != None and value != '':
        return True
    return False


def write_header(output, table_name, data_types):
    """
    ヘッダーをファイルに書き込む
    """
    value = "INSERT INTO " + table_name + "("
    for index, column_name in enumerate(data_types.keys()):
        value = value + column_name
        if index != len(data_types)-1:
            value = value + ","
    value = value + ") values\n"

    with open(output, mode='w') as f:
        f.write(value)


def write_footer(output):
    """
    フッターをファイルに書き込む
    """
    with open(output, mode='a') as f:
        f.write("\n;")


def write_body(sheet, output, data_types):
    """
    インサートするデータをファイルに書き込む
    """
    row = 5
    while is_exist_value(sheet, row, 1):
        # 最初の行はスキップ
        if row != 5:
            dataset = ",\n    ("
        else:
            dataset = "    ("

        for index, value in enumerate(data_types.values()):
            # 最初の列はスキップ
            if index != 0:
                dataset = dataset + ","

            # データを取得し適宜形式を変更する
            data = get_excel_value(sheet, row, index+1)
            if value == "var":
                dataset = dataset + change_variable_format(data)
            elif value in ("int", "statement"):
                dataset = dataset + change_raw_format(data)
            else:
                raise Exception("型: "+value+" には対応していません")

        dataset = dataset + ")"

        # ファイルに書き込む
        with open(output, mode='a') as f:
            f.write(dataset)

        row += 1


def change_variable_format(data):
    """
    var型のフォーマットに変形

    Returns
    ----------
    str
    """
    if data == "null" or data == None or data == "":
        return change_except_format(data)
    return "\'" + str(data) + "\'"


def change_raw_format(data):
    """
    特に加工の必要ないフォーマットを対象に変形
        例外ケース以外は文字列に変更するだけ

    Returns
    ----------
    str
    """
    if data == "null" or data == None or data == "":
        return change_except_format(data)
    return str(data)


def change_except_format(data):
    """
    特殊なデータが入ってる場合の変形処理

    Returns
    ----------
    str
    """
    if data == "null":
        return data
    if data == None or data == "":
        return "\'\'"


main()
