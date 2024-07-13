import os
import xlrd2
from openpyxl import load_workbook
import common
import shipyard
import shipowner
import inspection
from colorama import Fore, Style, init

init(autoreset=True)


def init():
    op_flag = True
    first = True
    while op_flag:
        try:

            if first:
                print(
                    f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}老婆大人，把那烦人的文件路径给我，让我给您搞定它：{Style.RESET_ALL}")  # 打印蓝色文字
            else:
                print()
                print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}老婆大人，老公我随时给您效劳：{Style.RESET_ALL}")

            first = False
            filePath = input()

            print()

            if filePath == '结束':
                print(f"{Fore.CYAN}老婆大人，处理程序开始切换，请根据提示操作...{Style.RESET_ALL}")
                op_flag = False
                raise KeyboardInterrupt

            # 路径不存在
            if not os.path.exists(filePath):
                print(
                    f"{Fore.RED}老婆大人，我的能力不够，找不到这个路径，请您惩罚我吧！{Style.RESET_ALL}")
            else:
                # 读取文件夹中的所有文件
                fileList = os.listdir(filePath)

                if len(fileList) > 0:
                    # 遍历文件
                    for xlsFile in fileList:
                        if (xlsFile.endswith(".xls") or xlsFile.endswith(".XLS") or
                                xlsFile.endswith(".xlsx") or xlsFile.endswith(".XLSX")):
                            xls_path = filePath + '\\' + xlsFile
                            print(
                                f"{Fore.LIGHTBLACK_EX}{Style.BRIGHT}老婆大人，老公努力搞定：{xls_path}{Style.RESET_ALL}")

                            # 创建一个新的工作簿（Workbook）对象
                            wb = load_workbook('./temp/demo.xlsx')

                            file_name = common.get_file_name(xls_path)

                            # 传入Excel文件路径打开文件
                            workbook = xlrd2.open_workbook(xls_path)
                            worksheet = workbook.sheet_by_name('滚动表')

                            if worksheet:
                                print()
                                print(f"{Fore.CYAN}老公正在努力劳作，老婆您稍等......{Style.RESET_ALL}")
                                print()
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

                            print(f"{Fore.LIGHTBLACK_EX}{Style.BRIGHT}老婆大人，老公搞定了：{xls_path}{Style.RESET_ALL}")
                            print()
                            print(
                                f"{Fore.WHITE}----------------------------分割线----------------------------{Style.RESET_ALL}")
                            print()
                else:
                    print(
                        f"{Fore.RED}老婆大人，您真好，路径下面是空的，是不是担心老公太累了？{Style.RESET_ALL}")
        except KeyboardInterrupt:
            raise KeyboardInterrupt

        except Exception as err:
            print(
                f"{Fore.RED}老婆大人，不好了，老公的程序异常了，错误信息: ${err}{Style.RESET_ALL}")

