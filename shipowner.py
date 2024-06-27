import common


# 写入船东一行数据
def write_sheet_data(wb, row, row_idx, obj):
    if row_idx < 2 or row[0] == '':
        return
    else:
        # 根据索引获取指定sheet
        sheet = wb.worksheets[2]

        if row_idx > 1:
            if row[3] == '' and row[4] == '':
                # 插入一行新的
                common.copy_row(sheet, obj, 1)

                obj.rowIndex = obj.rowIndex + 1
                common.write_categories(sheet, row, obj)
                obj.count = 1
            else:
                if row[6] != '√':
                    return
                else:
                    # 标题信息 只做第一次
                    common.write_title(sheet, row, 6)
                    # 写入行数据
                    common.copy_row(sheet, obj, 1)
                    common.copy_row(sheet, obj, 2)

                    obj.rowIndex = obj.rowIndex + 2
                    common.write_row(sheet, row, obj, 6)
