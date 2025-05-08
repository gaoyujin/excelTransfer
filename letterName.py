import re
import os
import shutil
from datetime import datetime
from colorama import Fore, Style, init


init(autoreset=True)


def process_filename(filename):
    # 使用正则表达式匹配文件名中的字母模式
    # 匹配 _A_S_, _A_S, _A_S  等多种格式
    pattern = r'_([A-Z])_([A-Z])_?'
    match = re.search(pattern, filename)
    
    if match:
        # 提取字母
        letters = [match.group(1), match.group(2)]
        
        # 获取不包含字母模式的其他部分
        other_parts = re.split(pattern, filename)
        # 过滤掉空字符串和None
        other_parts = [part for part in other_parts if part and part not in letters]
        # 去除每个字符串两端的空格
        cleaned_parts = [s.strip() for s in other_parts]
        
        # 用逗号连接其他部分，并在末尾添加字母
        new_name = ','.join(cleaned_parts + [match.group(1)])
        
        return new_name.rstrip(',')
    return filename


def init():
    op_flag = True
    first = True
    while op_flag:
        try:

            if first:
                print(
                    f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，把相关文件目录给我，我来重新排列文件名称（_A_S等）：{Style.RESET_ALL}")  # 打印蓝色文字
            else:
                print()
                print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}主人，奴仆我随时准备重新排列文件名称（_A_S等）：{Style.RESET_ALL}")

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

                        if (xlsFile.endswith(".pdf") or
                                xlsFile.endswith(".PDF")):

                            # 处理文件名（不包含扩展名）
                            name_without_ext = os.path.splitext(xlsFile)[0]
                            new_name = process_filename(name_without_ext)

                            if new_name != name_without_ext:
                                # 添加扩展名
                                new_filename = new_name + '.pdf'

                                # 构建目标文件路径
                                target_file = os.path.join(target_dir, new_filename)

                                # 复制文件到新位置
                                shutil.copy2(source_file, target_file)

                                processed_count += 1

                                # 如果文件名发生了变化，打印提示
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
