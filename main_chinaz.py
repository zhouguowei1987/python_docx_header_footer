import re
import time
from docx import Document
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx2pdf import convert
from docx.shared import RGBColor  # 设置字体的颜色
from docx.oxml.ns import qn
from win32com import client as wc
import os
import shutil
import rarfile
import re

rarfile.UNRAR_TOOL = "D:\\Program Files (x86)\\WinRAR\\UnRAR.exe"

def decompress_rar(rar_file_name, dir_name):
    """
    .rar 文件解压
    :param rar_file_name: rar 文件路径
    :param dir_name: 文件解压目录
    :return:
    """
    # 创建 rar 对象
    rar_obj = rarfile.RarFile(rar_file_name)
    # 目录切换
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)
    os.chdir(dir_name)
    # Extract all files into current directory.
    rar_obj.extractall()
    # rar_obj.extractall(dir_name)
    # 关闭
    rar_obj.close()


if __name__ == '__main__':
    # 解压压缩包
    # rar_root_dir = "E:\\workspace\\sc.chinaz.com\\2024-12-08"
    # rar_dirs = sorted(os.listdir(rar_root_dir))
    # rar_files = sorted(os.listdir(rar_root_dir))
    # for rar_file in rar_files:
    #     rar_file_path = rar_root_dir + "\\" + rar_file
    #     dst_file_path = "E:\\workspace\\sc.chinaz.com\\uncompress_2024-12-08"
    #     if os.path.splitext(rar_file)[1] in [".ppt", ".pptx"]:
    #         print("==========" + "开始复制" + "==========")
    #         shutil.copy(rar_file_path, dst_file_path + "\\" + rar_file)
    #         print("==========" + "复制完成" + "==========")
    #     else:
    #         print("==========" + "开始解压" + rar_file_path + "==========")
    #         try:
    #             decompress_rar(rar_file_path, dst_file_path+"\\"+rar_file.replace(".rar", ""))
    #         except Exception as e:
    #             print(e)
    #             continue
    #         print("==========" + "解压完成" + "==========")
    # exit()

    root_dir = "E:\\workspace\\sc.chinaz.com\\uncompress_2024-12-08"
    dir_files = sorted(os.listdir(root_dir))
    for dir_file in dir_files:
        dir_file_path = root_dir + "\\" + dir_file
        print(dir_file_path)

        if os.path.isdir(dir_file_path):
            first_child_dir_files = sorted(os.listdir(dir_file_path))
            for first_child_dir_file in first_child_dir_files:
                first_child_dir_file_path = dir_file_path + "\\" + first_child_dir_file
                print(first_child_dir_file_path)
                if os.path.isdir(first_child_dir_file_path):
                    second_child_dir_files = sorted(os.listdir(first_child_dir_file_path))
                    for second_child_dir_file in second_child_dir_files:
                        print(second_child_dir_file)
                        extension = os.path.splitext(second_child_dir_file)[-1]
                        print(extension)
                        if extension in [".ppt", ".pptx"]:
                            src_child_file_path = first_child_dir_file_path + "\\" + second_child_dir_file
                            dst_child_file_path = root_dir + "\\" + dir_file + extension
                            print(src_child_file_path)
                            print(dst_child_file_path)
                            try:
                                os.rename(src_child_file_path, dst_child_file_path)
                            except WindowsError:
                                os.remove(dst_child_file_path)
                                os.rename(src_child_file_path, dst_child_file_path)

