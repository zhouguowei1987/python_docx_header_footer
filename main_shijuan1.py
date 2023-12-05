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
rarfile.UNRAR_TOOL = "F:\\WinRAR\\UnRAR.exe"


def change_word_font(doc_file):
    try:
        # 打开doc文件
        doc = Document(doc_file)
        doc.styles['Normal'].font.name = u'Times New Roman'  # 设置西文字体
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')  # 设置中文字体使用字体2->宋体
        doc.save(doc_file)
        return True
    except Exception as e:
        print(e)
        return False


def remove_header_footer(doc):
    # doc：需要去页眉页脚的docx 文件
    try:
        document = Document(doc)
        for section in document.sections:
            section.different_first_page_header_footer = False
            section.header.is_linked_to_previous = True
            section.footer.is_linked_to_previous = True
        document.save(doc)
        return True
    except Exception as e:
        print(e)
        return False


def doc2docx(in_file, out_file):
    try:
        word = wc.Dispatch("Word.Application")
        try:
            print(in_file)
            print(out_file)
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, 12, False, "", True, "", False, False, False, False)
            print('转换成功')
            doc.Close()
            word.Quit()
            return True
        except Exception as e:
            print(e)
    except Exception as e:
        print(e)
    return False


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
    # category_rars_arr = ['中考试卷', '高考试卷']
    # category_rars_arr = ['高考试卷']
    # rar_root_dir = "G:\\www.rar_shijuan1.com"
    # rar_dirs = sorted(os.listdir(rar_root_dir))
    # for category_rar in category_rars_arr:
    #     rar_files = sorted(os.listdir(rar_root_dir + "\\" + category_rar))
    #     for rar_file in rar_files:
    #         rar_file_path = rar_root_dir+ "\\" + category_rar + "\\" + rar_file
    #         print("==========" + "开始解压" + rar_file_path + "==========")
    #         try:
    #             decompress_rar(rar_file_path, "G:\\www.shijuan1.com\\www.shijuan1.com\\" + category_rar)
    #         except Exception as e:
    #             print(e)
    #             continue
    #         print("==========" + "解压完成" + "==========")
    # exit()

    # category_dirs_arr = ['中考试卷', '高考试卷']
    category_dirs_arr = ['高考试卷']
    root_dir = "G:\\www.shijuan1.com\\www.shijuan1.com"
    category_dirs = sorted(os.listdir(root_dir))
    for category in category_dirs:
        if category in category_dirs_arr:
            files = sorted(os.listdir(root_dir + "\\" + category))
            for file in files:
                file_path = root_dir + "\\" + category + "\\" + file
                print(file_path)

                # 文件名称小于20，则删除
                # if len(file) < 20:
                #     os.remove(file_path)

                # # 查看一下是否是文件夹，如果是文件夹，则将文件移出
                # if os.path.isdir(file_path):
                #     child_files = sorted(os.listdir(file_path))
                #     for child_file in child_files:
                #         extension = os.path.splitext(child_file)[-1]
                #         if extension not in [".doc", ".docx"]:
                #             continue
                #         src_child_file_path = file_path + "\\" + child_file
                #         dst_child_file_path = root_dir + "\\" + category + "\\" + child_file
                #         try:
                #             os.rename(src_child_file_path, dst_child_file_path)
                #         except WindowsError:
                #             os.remove(dst_child_file_path)
                #             os.rename(src_child_file_path, dst_child_file_path)

                docx_dir = "G:\\www.shijuan1.com\\docx.www.shijuan1.com" + "\\" + category
                if not os.path.exists(docx_dir):
                    os.makedirs(docx_dir)

                sub_file = file
                left_flag_index = file.find("【")
                right_flag_index = file.find("】")
                if left_flag_index == 0 and right_flag_index != -1:
                    # 文档名称以“【”开头，以“】”结尾，则替换名称
                    sub_file = file[right_flag_index + 1:]
                docx_file = docx_dir + "\\" + sub_file.lower().replace(os.path.splitext(sub_file)[1], ".docx")
                docx_file = docx_file.strip()
                docx_file = docx_file.replace("（", "(").replace("）", ")")
                docx_file = docx_file.replace("精品解析：", "")
                docx_file = docx_file.replace("【KS5U+高考】", "")
                docx_file = docx_file.replace("【ks5u+高考】", "")
                docx_file = docx_file.replace("【品优教学】", "")
                docx_file = docx_file.replace("——", "")
                docx_file = docx_file.replace("+Word版", "")
                docx_file = docx_file.replace(" Word版", "")
                docx_file = docx_file.replace("(word答案)", "")
                docx_file = docx_file.replace("(word精校版)", "")
                docx_file = docx_file.replace("(word版)", "")
                docx_file = docx_file.replace("(word答案)", "")
                docx_file = docx_file.replace("(word解析版)", "(含解析)")
                docx_file = docx_file.replace("(word版无答案)", "")
                docx_file = docx_file.replace("(word版回忆版无答案)", "")
                docx_file = docx_file.replace("(word版，含听力原文)", "")
                docx_file = docx_file.replace("(word版含答案)", "")
                docx_file = docx_file.replace("(word版，有答案)", "")
                docx_file = docx_file.replace("(文字版-含答案)", "")
                docx_file = docx_file.replace("word版", "")
                docx_file = docx_file.replace("-", "")
                docx_file = docx_file.replace(",", "")
                docx_file = docx_file.replace("，", "")

                if not os.path.exists(docx_file):
                    # 获取文件后缀
                    file_ext = os.path.splitext(file_path)[-1]
                    if file_ext == ".docx":
                        # 已经是docx文件了，直接复制过去
                        shutil.copy(file_path, docx_file)
                    else:
                        with open(docx_file, 'w') as f:
                            pass
                        print("==========开始转化为docx==============")
                        if not doc2docx(file_path, docx_file):
                            os.remove(docx_file)
                            continue
                        print("==========转化完成==============")

                if os.path.exists(docx_file):
                    # 删除页眉页脚
                    if not remove_header_footer(docx_file):
                        os.remove(docx_file)
                        continue

                    # 改变文档字体
                    if not change_word_font(docx_file):
                        os.remove(docx_file)
                        continue
