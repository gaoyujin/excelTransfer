import os
from openpyxl import load_workbook
import common
from colorama import Fore, Style, init

init(autoreset=True)


class FileName:
    pass


def init():
    op_flag = True
    first = True
    while op_flag:
        try:

            if first:
                print(
                    f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，把文件集合目录给我，我来收集图名、图号：{Style.RESET_ALL}")  # 打印蓝色文字
            else:
                print()
                print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，奴仆我随时准备收集图名、图号：{Style.RESET_ALL}")

            first = False
            file_path = input()

            if file_path == '切换':
                print()
                print(f"{Fore.CYAN}主人，处理程序开始切换，请根据提示操作...{Style.RESET_ALL}")
                op_flag = False
                raise KeyboardInterrupt

            # 路径不存在
            if not os.path.exists(file_path):
                print()
                print(
                    f"{Fore.RED}主人，我的能力不够，找不到这个路径，请您惩罚我吧！{Style.RESET_ALL}")
            else:
                print()
                print(f"{Fore.CYAN}奴仆正在努力劳作，主人您稍等......{Style.RESET_ALL}")
                print()
                # 读取文件夹中的所有文件
                file_list = os.listdir(file_path)

                if len(file_list) > 0:
                    last_data = []
                    set_data = []
                    count = 1

                    # 遍历文件
                    for xlsFile in file_list:
                        if (xlsFile.endswith(".dwg") or xlsFile.endswith(".DWG") or
                                xlsFile.endswith(".pdf") or xlsFile.endswith(".PDF") or
                                xlsFile.endswith(".xls") or xlsFile.endswith(".XLS") or
                                xlsFile.endswith(".xlsx") or xlsFile.endswith(".XLSX")):

                            file_name = common.get_file_name(xlsFile)
                            names = file_name.split("_")

                            if len(names) < 1:
                                continue

                            if names[0] in set_data:
                                continue
                            else:
                                set_data.append(names[0])
                                my_data = FileName()
                                my_data.count = count
                                my_data.account = names[0]
                                my_data.name = (file_name.replace(names[0] + '_', '').
                                                replace('A_S ', '').replace('A_S_', '').replace('A_S', ''))
                                my_data.allName = xlsFile

                                last_data.append(my_data)
                                count = count + 1

                    if len(last_data) > 0:
                        # 创建一个新的工作簿（Workbook）对象
                        wb = load_workbook('./temp/collect.xlsx')
                        # 根据索引获取指定sheet
                        sheet = wb.worksheets[0]

                        # 创建一个对象并设置属性
                        my_object = common.row_object(2, 1)

                        for item in last_data:
                            if not item or not item.account:
                                break

                            # 插入一行新的
                            common.copy_row(sheet, my_object, 1)

                            # 插入一行新的
                            row_index = item.count + 1
                            sheet['A' + str(row_index)] = item.count
                            sheet['B' + str(row_index)] = item.account
                            sheet['C' + str(row_index)] = item.name
                            sheet['D' + str(row_index)] = item.allName
                            my_object.rowIndex = my_object.rowIndex + 1

                        # 不存在 output 则创建
                        save_file_path = file_path + "\\collect_output"
                        common.create_directory_if_not_exists(save_file_path)

                        # 存储结果文件
                        save_path = file_path + "\\collect_output\\" + "送审文件图名图号集合.xlsx"
                        wb.save(save_path)

                    print(f"{Fore.LIGHTBLACK_EX}{Style.BRIGHT}主人，收集图名、图号完成了：{file_path}{Style.RESET_ALL}")
                else:
                    print(
                        f"{Fore.RED}主人，您真好，路径下面是空的，是不是担心奴仆太累了？{Style.RESET_ALL}")
        except KeyboardInterrupt:
            raise KeyboardInterrupt

        except Exception as err:
            print(
                f"{Fore.RED}主人，不好了，奴仆的程序异常了，错误信息: ${err}{Style.RESET_ALL}")

