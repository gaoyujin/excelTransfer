import excel
import collect
from colorama import Fore, Style, init


def run(mode):
    if mode == 'excel':
        excel.init()
    else:
        collect.init()


# 变更模式
def change_run(tip):
    if tip == 'init':
        print()
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}1：送检Excel文件转换{Style.RESET_ALL}")
        print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}2：PDF等文件信息收集{Style.RESET_ALL}")

    is_run = False
    print()
    std = input("请输入要执行的操作的编号（1 or 2）: ")

    if std == '1':
        try:
            run('excel')
        except KeyboardInterrupt:
            change_run('no')
    if std == '2':
        try:
            run('collect')
        except KeyboardInterrupt:
            change_run('no')
    if std != '1' and std != '2':
        is_run = True
        print(f"{Fore.RED}老婆大人，输入的操作编号不存在{Style.RESET_ALL}")

    while is_run:
        change_run('no')


# 允许默认代码
try:
    run('excel')
except KeyboardInterrupt:
    change_run('init')


