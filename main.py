import excel
import collect
import fileName
from colorama import Fore, Style, init


def run(mode):
    if mode == 'excel':
        excel.init()
    if mode == 'collect':
        collect.init()
    if mode == 'fileName':
        fileName.init()


# 变更模式
def change_run(tip):
    if tip == 'init':
        print()
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}# 输入提示：{Style.RESET_ALL}")
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 1 or 图纸目录 or 目录 (送检Excel文件转换){Style.RESET_ALL}")
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 2 or 清单list or 清单 (PDF等文件信息收集){Style.RESET_ALL}")
        print(
            f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 3 or CCS特殊字符 or CCS (去除目录中文件的特殊字符){Style.RESET_ALL}")
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}# 模式切换说明：{Style.RESET_ALL}")
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》'切换'后按回车键{Style.RESET_ALL}")

    is_run = False
    print()
    std = input("请输入要执行的操作的编号或者关键字: ")

    if std == '1' or std == '图纸目录' or std.startswith('目录'):
        try:
            run('excel')
        except KeyboardInterrupt:
            change_run('no')
    if std == '2' or std == '清单list' or std.startswith('清单'):
        try:
            run('collect')
        except KeyboardInterrupt:
            change_run('no')
    if std == '3' or std == 'CCS特殊字符' or std.startswith('CCS'):
        try:
            run('fileName')
        except KeyboardInterrupt:
            change_run('no')
    if std != '1' and std != '2' and std != '3':
        is_run = True
        print(f"{Fore.RED}主人，输入的操作编号不存在{Style.RESET_ALL}")

    while is_run:
        change_run('no')


# 允许默认代码
try:
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}# 输入提示：{Style.RESET_ALL}")
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 1 or 图纸目录 or 目录 (送检Excel文件转换){Style.RESET_ALL}")
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 2 or 清单list or 清单 (PDF等文件信息收集){Style.RESET_ALL}")
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》 3 or CCS特殊字符 or CCS (去除目录中文件的特殊字符){Style.RESET_ALL}")
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}# 模式切换说明：{Style.RESET_ALL}")
    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}输入==》'切换'后按回车键{Style.RESET_ALL}")
    change_run('no')
except KeyboardInterrupt:
    change_run('init')


