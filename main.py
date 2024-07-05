import os
import xlrd2
from openpyxl import load_workbook
import common
import shipyard
import shipowner
import inspection



op_flag = True
while op_flag:
    try:
        print(f"请输入需要分析的Excel路径====》")
        filePath = input()

        # 路径不存在
        if not os.path.exists(filePath):
            print(f"输入路径不存在！")
        else:
            # 读取文件夹中的所有文件
            fileList = os.listdir(filePath)

            if len(fileList) > 0:
                # 遍历文件
                for xlsFile in fileList:
                    if (xlsFile.endswith(".xls") or xlsFile.endswith(".XLS") or
                            xlsFile.endswith(".xlsx") or xlsFile.endswith(".XLSX")):
                        xls_path = filePath + '\\' + xlsFile
                        print(f"开始处理文件：{xls_path}")

                        # 创建一个新的工作簿（Workbook）对象
                        wb = load_workbook('./temp/demo.xlsx')

                        file_name = common.get_file_name(xls_path)

                        # 传入Excel文件路径打开文件
                        workbook = xlrd2.open_workbook(xls_path)
                        worksheet = workbook.sheet_by_name('滚动表')

                        if worksheet:
                            # 创建一个对象并设置属性
                            my_shipyard = common.row_object(6, 1)
                            # 创建一个对象并设置属性
                            my_inspection = common.row_object(6, 1)
                            # 创建一个对象并设置属性
                            my_shipowner = common.row_object(6, 1)
                            # 是否开始处理
                            is_start = False
                            for row_idx in range(worksheet.nrows):
                                row = worksheet.row_values(row_idx)

                                if()

                                if not is_start and common.is_chinese_number(row[1]) and row[0].isalpha():
                                    is_start = True

                                # 写入
                                if is_start:
                                    shipyard.write_sheet_data(wb, row, row_idx, my_shipyard)
                                    # 写入船检
                                    inspection.write_sheet_data(wb, row, row_idx, my_inspection)
                                    # 写入船东
                                    shipowner.write_sheet_data(wb, row, row_idx, my_shipowner)

                        # 不存在 output 则创建
                        save_file_path = filePath + "\\output"
                        common.create_directory_if_not_exists(save_file_path)

                        # 存储结果文件
                        save_path = filePath + "\\output\\" + "送审目录-" + file_name + ".xlsx"
                        wb.save(save_path)

                        print(f"处理完成：{xls_path}")

            else:
                print(f"文件夹中没有相关文件！")
    except Exception as err:
        print(f"异常了，错误信息: ${err}")
