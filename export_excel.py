import xlrd
import xlwt
import os


def read_file():
    path = './input'
    for root, dirs, files in os.walk(path):
        for file in files:
            if os.path.splitext(file)[1] == ".xlsx":
                read_excel(root + '/' + file)


def read_excel(filename):
    """
    :param filename:
    :return:
    """
    excel = xlrd.open_workbook(filename)
    sheet = excel.sheet_by_index(0)
    names = sheet.col_values(7)

    d = {}
    for (i, v) in enumerate(names):
        if d.__contains__(v):
            d[v].append(i)
        else:
            d[v] = [i]
    d.pop('')
    d.pop('历史审批人姓名')
    write_excel(d, sheet)


def write_excel(d, sheet):
    """
    :param d:
    :param sheet:
    :return:
    """
    title = sheet.row_values(0)[0]
    for k, ele in d.items():
        stu = [sheet.row_values(0), sheet.row_values(1), sheet.row_values(2), sheet.row_values(3)]

        for e in ele:
            stu.append(sheet.row_values(e))

        book = xlwt.Workbook()  # 新建一个excel
        style = xlwt.XFStyle()
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
        sheet_n = book.add_sheet('sheet1')  # 添加一个sheet页
        row = 0  # 控制行

        for st in stu:
            col = 0  # 控制列
            for s in st:  # 再循环里面list的值，每一列
                sheet_n.write(row, col, s, style)
                col += 1
            row += 1
        #     指定目录
        path = r"./output/" + title + '_' + k + ".xlsx"
        book.save(path)


if __name__ == '__main__':
    read_file()
