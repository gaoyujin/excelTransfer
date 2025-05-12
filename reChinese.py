import os
import shutil
from datetime import datetime
from colorama import Fore, Style, init

init(autoreset=True)


def is_chinese(char):
    """判断字符是否是中文"""
    return '\u4e00' <= char <= '\u9fff'


def remove_chinese_from_filename(filename):
    """从第一个中文到最后一个中文的这段内容都去除，且这段内容的前面的空格去除"""
    # 找到第一个和最后一个中文字符的位置
    first_chinese_pos = -1
    last_chinese_pos = -1
    
    for i, char in enumerate(filename):
        if is_chinese(char):
            if first_chinese_pos == -1:
                first_chinese_pos = i
            last_chinese_pos = i
    
    # 如果没有找到中文字符，返回原文件名
    if first_chinese_pos == -1:
        return filename
    
    # 获取第一个中文前的内容（去除空格）和最后一个中文后的内容
    prefix = filename[:first_chinese_pos].rstrip()
    suffix = filename[last_chinese_pos + 1:]
    
    # 组合新文件名
    new_filename = prefix + suffix
    
    # 处理多余的空格
    new_filename = ' '.join(new_filename.split())
    
    return new_filename


def init():
    op_flag = True
    first = True
    while op_flag:
        try:

            if first:
                print(
                    f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，把相关文件目录给我，我来去除文件名称中的中文：{Style.RESET_ALL}")  # 打印蓝色文字
            else:
                print()
                print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，奴仆我随时准备去除文件名称中的中文：{Style.RESET_ALL}")

            first = False
            file_path = input().strip()

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

                    # 创建新的目标目录（使用时间戳来避免重名）
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    target_dir = os.path.join(file_path, f"processed_files_{timestamp}")
                    os.makedirs(target_dir, exist_ok=True)

                    # 计数器
                    processed_count = 0

                    # 遍历文件
                    for xlsFile in file_list:

                        # 获取完整的源文件路径
                        source_file = os.path.join(file_path, xlsFile)

                        if (xlsFile.endswith(".pdf") or xlsFile.endswith(".PDF") or
                                xlsFile.endswith(".xls") or xlsFile.endswith(".XLS") or
                                xlsFile.endswith(".xlsx") or xlsFile.endswith(".XLSX")):

                            old_filepath = os.path.join(file_path, xlsFile)

                            # 清理文件名中的特殊字符
                            new_filename = remove_chinese_from_filename(xlsFile)

                            # 构建目标文件路径
                            target_file = os.path.join(target_dir, new_filename)

                            # 复制文件
                            shutil.copy2(old_filepath, target_file)
                            processed_count += 1

                            # 如果文件名发生了变化，打印提示
                            if xlsFile != new_filename:
                                print(f"文件名已更改: '{xlsFile}' -> '{new_filename}'")

                    print(f"{Fore.LIGHTBLACK_EX}{Style.BRIGHT}\n处理完成！{file_path}{Style.RESET_ALL}")
                    print(f"共处理了 {processed_count} 个PDF文件")
                    print(f"处理后的文件保存在: {target_dir}")
                else:
                    print(
                        f"{Fore.RED}主人，您真好，路径下面是空的，是不是担心奴仆太累了？{Style.RESET_ALL}")
        except KeyboardInterrupt:
            raise KeyboardInterrupt

        except Exception as err:
            print(
                f"{Fore.RED}主人，不好了，奴仆的程序异常了，错误信息: ${err}{Style.RESET_ALL}")

