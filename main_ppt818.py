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
import zipfile

zipfile.UNRAR_TOOL = "D:\\Program Files (x86)\\WinRAR\\UnRAR.exe"


def decompress_zip(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)


if __name__ == '__main__':
    # 解压压缩包
    # zip_root_dir = "E:\\workspace\\www.ppt818.com\\2024-12-20\\www.zip_ppt818.com"
    # zip_dirs = sorted(os.listdir(zip_root_dir))
    # zip_files = sorted(os.listdir(zip_root_dir))
    # for zip_file in zip_files:
    #     zip_file_path = zip_root_dir + "\\" + zip_file
    #     dst_file_path = "E:\\workspace\\www.ppt818.com\\2024-12-20\\www.uncompress_ppt818.com"
    #     print("==========" + "开始解压" + zip_file_path + "==========")
    #     try:
    #         decompress_zip(zip_file_path, dst_file_path+"\\"+zip_file.replace(".zip", ""))
    #     except Exception as e:
    #         print(e)
    #         continue
    #     print("==========" + "解压完成" + "==========")
    # exit()

    # root_dir = "E:\\workspace\\www.ppt818.com\\2024-12-20\\www.uncompress_ppt818.com"
    # files = sorted(os.listdir(root_dir))
    # for file in files:
    #     file_path = root_dir + "\\" + file
    #     print(file_path)
    #
    #     # 查看一下是否是文件夹，如果是文件夹，则将文件移出
    #     if os.path.isdir(file_path):
    #         child_files = sorted(os.listdir(file_path))
    #         for child_file in child_files:
    #             extension = os.path.splitext(child_file)[-1]
    #             if extension not in [".ppt", ".pptx"]:
    #                 continue
    #             src_child_file_path = file_path + "\\" + child_file
    #             dst_child_file_path = root_dir + "\\" + file + extension
    #             try:
    #                 os.rename(src_child_file_path, dst_child_file_path)
    #             except WindowsError:
    #                 os.remove(dst_child_file_path)
    #                 os.rename(src_child_file_path, dst_child_file_path)

    root_dir = "E:\\workspace\\www.ppt818.com\\2024-12-20\\www.uncompress_ppt818.com"
    files = sorted(os.listdir(root_dir))
    for file in files:
        file_path = root_dir + "\\" + file
        print(file_path)

        # 删除文件名中含有“课件”字样文件
        if "课件" in file:
            os.remove(file_path)

        # 删除文件名中不含有“PPT模板”字样文件
        if "PPT模板" not in file:
            os.remove(file_path)
