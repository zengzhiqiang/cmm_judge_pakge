def judge_test_data(stand_data, up_tol, low_tol, test_data):
    '''
    这是一个判定被测数据是否合格的函数
    :param stand_data: 标准值。type = float 传入数据可为"/"
    :param up_tol: 上公差。type = float 传入数据可为"/"
    :param low_tol: 下公差。type = float 传入数据可为"/"
    :param test_data: 测试数据。type = list。 数组成员 type = float
    :return:
    1、数据均满足要求，且上下公差均不为"/":"OK"
    2、数据不满足要求："NO", 和一个记录错误数据位置的列表。
    3、数据满足要求，但上公差或者下公差为"/"或者标准值为"/" 或者上公差，下公差同时为"/": "/"
    '''
    err_data = {}
    if stand_data == "/" or (up_tol == "/" and low_tol == "/"):
        '''
        标准值为空，或者上下公差同时为空。直接返回不判定的结果。
        '''
        return "/", err_data
    elif up_tol == "/":
        status, err_data = judge_test_dataum(stand_data, test_data, low_tol=low_tol)
        if status:
            return "/", err_data
        else:
            return "NO", err_data
    elif low_tol == "/":
        status, err_data = judge_test_dataum(stand_data, test_data, up_tol=up_tol)
        if status:
            return "/", err_data
        else:
            return "NO", err_data
    else:
        status, err_data = judge_test_dataum(stand_data, test_data, up_tol=up_tol, low_tol=low_tol)
        if status:
            return "OK", err_data
        else:
            return "NO", err_data

def judge_test_dataum(stand_data, test_data, up_tol=10000, low_tol=10000):
    '''

    :param stand_data:
    :param up_tol:
    :param low_tol:
    :param test_data:
    :return:
    '''
    status = True
    i = 0
    err_data = {}
    for test_dataum in test_data:
        if test_dataum >= stand_data - low_tol and test_dataum <= stand_data + up_tol:
            #print(stand_data, up_tol, low_tol, test_dataum)
            pass
        else:
            i += 1
            status = False
            err_data.update({i: test_dataum})
            #print(err_data)
    return status, err_data

import os
import xlrd
import re

def judge_a_workbook(file_path, file):
    work_book_statu = True  #记录该报告有没有错误的标记
    work_book = xlrd.open_workbook(file_path)
    sheet_names = work_book.sheet_names()
    title_name = re.search(r'20\d\d[-]\d\d[-]\d\d\d', file).group()
    for sheet_name in sheet_names:
        table = work_book.sheet_by_name(sheet_name)
        report_name_data = table.row_values(1)
        name_match = False
        for report_name_datum in report_name_data:
            report_name = re.search(r'20\d\d[-]\d\d[-]\d\d\d', report_name_datum)
            if report_name and report_name.group() == title_name:
                name_match = True
        if name_match:
            pass
        else:
            work_book_statu = False
            print("报告编号错了！")
        rows = table.nrows
        cols = table.ncols
        for row in range(0, rows):
            data = table.row_values(row)
            if type(data[0]) == float:
                try:
                    stand = float(re.search(r"\d+(\.\d+)?", str(data[1])).group())
                except:
                    stand = "/"
                try:
                    up_tol = float(re.search(r"\d+(\.\d+)?", str(data[2])).group())
                except:
                    up_tol = "/"
                try:
                    low_tol = float(re.search(r"\d+(\.\d+)?", str(data[3])).group())
                except:
                    low_tol = "/"
                judge = str(data[len(data) - 2]).replace(" ", "")
                test_data = []
                for col in range(4, cols - 2):
                    if table.cell(row, col).value and table.cell(row, col).value != "/":
                            test_datums = str(table.cell(row, col).value).split("、")
                            for test_datum in test_datums:
                                test_data.append(float(test_datum))
                result, err_data = judge_test_data(stand, up_tol, low_tol, test_data)
                if result == judge:
                    pass
                elif result == "OK":
                    print("第", str(data[0]) + "行应该判定为: OK")
                    work_book_statu = False
                elif result == "/":
                    print("第", str(data[0]) + "行应该判定为: /")
                    work_book_statu = False
                else:
                    work_book_statu = False
                    print("第", str(data[0]) + "行错误数据有下：")
                    for k, err_datum in err_data.items():
                        print("第" + str(k) + "个数值为：" + str(err_datum) + "，超差！")
    return work_book_statu

if __name__ == "__main__":
    while True:
        files = []
        count = False  #记录是否有报告
        path = input("请填入报告所在文件夹路径：")
        try:
            files = os.listdir(path)
        except:
            count = True
            print("无法找到指定路径：" + path)
            print("按任意键继续")
        for file in files:
            if file[-4:] == '.xls' or file[-5:] == '.xlsx':
                if re.search(r'20\d\d[-]\d\d[-]\d\d\d', file):
                    count = True
                    print("正在检查" + file + "...")
                    file_path = os.path.join(path, file)
                    work_book_statu = judge_a_workbook(file_path, file)
                    if work_book_statu:
                        print(file + "报告无误！")
        if count:
            pass
        else:
            print("该文件夹下没有报告，命名错误？！")
            print("按任意键继续")
        input()