
import os
import copy
import re

# 计数器类
class row_object:
    def __init__(self, row_index, count):
        self.count = count
        self.rowIndex = row_index


# 定义匹配中文数字的正则表达式
pattern = re.compile(r'^[一二三四五六七八九十百千万亿]+$', re.UNICODE)


# 判断输入字符串是否为中文数字
def is_chinese_number(s):
    return bool(pattern.match(s))


# 获取文件名称
def get_file_name(xls_file):
    arr_list = xls_file.split('\\')
    name = arr_list[len(arr_list) - 1]
    new_name = name[0:len(name) - 4]
    if name.endswith(".xlsx") or name.endswith(".XLSX"):
        new_name = name[0:len(name) - 5]
    return new_name


# 创建文件夹
def create_directory_if_not_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


# 写入标题信息
def write_title(sheet, row, num):
    value = sheet['A1'].value
    if value or row[num] != '√':
        return
    else:
        # 标题信息
        sheet['A1'] = row[3] + '\n' + row[4]
        sheet['D1'] = row[2]
        sheet['F1'] = 'PAGE'


# 写入分类
def write_categories(sheet, row, obj):
    # 插入一行新的
    sheet['A' + str(obj.rowIndex - 1)] = row[1]
    sheet['B' + str(obj.rowIndex - 1)] = row[2]


# 写入一行数据
def write_row(sheet, row, obj, num):
    # 插入一行新的
    if row[num] != '√':
        return
    else:
        # 插入一行新的
        sheet['A' + str(obj.rowIndex - 2)] = obj.count
        sheet['B' + str(obj.rowIndex - 2)] = row[2]
        sheet['C' + str(obj.rowIndex - 2)] = row[3]
        sheet['C' + str(obj.rowIndex - 1)] = row[4]
        obj.count = obj.count + 1


# 复制一行
def copy_row(sheet, obj, num):
    # 插入一行新的
    source_row = sheet[obj.rowIndex]
    # 插入行
    for cell in source_row:
        target_cell = sheet.cell(row=obj.rowIndex + num, column=cell.column)
        target_cell.value = cell.value
        # 设置单元格格式
        target_cell.fill = copy.copy(cell.fill)

        if cell.has_style:
            # target_cell._style = copy.copy(cell._style)
            target_cell.font = copy.copy(cell.font)
            target_cell.border = copy.copy(cell.border)
            target_cell.fill = copy.copy(cell.fill)
            target_cell.alignment = copy.copy(cell.alignment)

    sheet.row_dimensions[obj.rowIndex + num].height = 25

