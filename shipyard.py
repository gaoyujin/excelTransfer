import common
from openpyxl.worksheet.pagebreak import Break


# 写入船厂一行数据
def write_sheet_data(wb, row, row_idx, obj):
    if row[0] == '':
        return
    else:
        # 根据索引获取指定sheet
        sheet = wb.worksheets[0]

        if row_idx > 0:
            if row[3] == '' and row[4] == '':
                # 插入一行新的
                common.copy_row(sheet, obj, 1)

                obj.rowIndex = obj.rowIndex + 1

                if obj.rowIndex > 10:
                    page_break = Break(obj.rowIndex-2)  # 创建分页对象
                    # 其中i或者j为行号或者列号
                    sheet.row_breaks.append(page_break)

                common.write_categories(sheet, row, obj)
                obj.count = 1
            else:
                if row[5] != '√':
                    return
                else:
                    # 标题信息 只做第一次
                    common.write_title(sheet, row, 5)
                    # 写入行数据
                    common.copy_row(sheet, obj, 1)
                    common.copy_row(sheet, obj, 2)

                    obj.rowIndex = obj.rowIndex + 2
                    common.write_row(sheet, row, obj, 5)

