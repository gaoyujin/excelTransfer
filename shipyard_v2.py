import common_v2
import common
from openpyxl.worksheet.pagebreak import Break


# 写入船厂一行数据
def write_sheet_data(wb, row, row_idx, obj):
    if row[0] == '' or common_v2.contains_keywords(row[0]):
        return
    else:
        # 根据索引获取指定sheet
        sheet = wb['船厂']

        # 写入标题信息
        if row_idx == 0:
            common_v2.write_title_v2(sheet, row)

        # 写入内容
        if row_idx > 4:
            # 分段标题
            if row[1] == '' and row[2] == '':
                # 插入一行新的
                common.copy_row(sheet, obj, 1)

                obj.rowIndex = obj.rowIndex + 1

                if obj.rowIndex > 10:
                    page_break = Break(obj.rowIndex - 2)  # 创建分页对象
                    # 其中i或者j为行号或者列号
                    sheet.row_breaks.append(page_break)

                common_v2.write_categories_v2(sheet, row, obj)
                obj.count = 1
            else:
                if row[0] == '':
                    return
                else:
                    # 写入行数据
                    common.copy_row(sheet, obj, 1)
                    common.copy_row(sheet, obj, 2)

                    obj.rowIndex = obj.rowIndex + 2
                    common_v2.write_row_v2(sheet, row, obj, 5)

