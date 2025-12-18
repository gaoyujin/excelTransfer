from openpyxl.styles import Font
import re


def write_title_v2(sheet, row):
    # 标题信息
    sheet['A1'] = row[0]
    sheet['D1'] = row[3]
    sheet['F1'] = 'PAGE'


# 处理【图 名】内容
def split_text(text):
    # 正则匹配最后一个换行符（/n），并捕获前后两部分
    match = re.search(r'^(.*?)(\n)(.*)$', text)
    if match:
        # 中文部分：从开头到最后一个/n前（不包含/n）
        chinese_text = match.group(1)
        # 英文部分：最后一个/n之后的内容
        non_chinese_text = match.group(3)
    else:
        # 若无/n，中文部分为整个字符串，英文部分为空
        chinese_text = text
        non_chinese_text = ""
    return chinese_text, non_chinese_text


# 写入一行数据
def write_row_v2(sheet, row, obj, num):
    # 插入一行新的
    chinese_part, non_chinese_part = split_text(row[2])
    non_chinese_part = strip_prefixes(non_chinese_part)
    # 插入一行新的
    sheet['A' + str(obj.rowIndex - 2)] = row[0]
    sheet['B' + str(obj.rowIndex - 2)] = row[1]
    sheet['C' + str(obj.rowIndex - 2)] = chinese_part
    sheet['C' + str(obj.rowIndex - 1)] = non_chinese_part
    obj.count = obj.count + 1


def write_categories_v2(sheet, row, obj):
    arr = split_string(row[0])
    # 插入一行新的
    sheet['A' + str(obj.rowIndex - 1)] = arr[0]
    sheet['B' + str(obj.rowIndex - 1)] = arr[1]
    # 创建一个加粗的字体对象
    bold_font_a = Font(bold=True)
    sheet['A' + str(obj.rowIndex - 1)].font = bold_font_a
    bold_font_b = Font(bold=True)
    sheet['B' + str(obj.rowIndex - 1)].font = bold_font_b


# 切割部分的标题
def split_string(text):
    parts = text.split('.')
    return [part.strip() for part in parts if part.strip()]


# 去除英文文字头部多出来的内容
def strip_prefixes(input_str):
    prefixes = ['\n', '-\n', ' \n']
    for prefix in prefixes:
        if input_str.startswith(prefix):
            return input_str[len(prefix):]
    return input_str


# 判断是否跳过标题行
def contains_keywords(text):
    if type(text) == str:
        keywords = ["序号", "图名", "图号"]
        for keyword in keywords:
            if keyword in text:
                return True
        return False
    else:
        return False
