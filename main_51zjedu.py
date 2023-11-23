import re
import time
from docx import Document
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx2pdf import convert
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor  # 设置字体的颜色
from docx.oxml.ns import qn
from win32com import client as wc
import os
import shutil


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


def remove_header_footer(doc_file):
    # doc：需要去页眉页脚的docx 文件
    try:
        doc = Document(doc_file)
        for section in doc.sections:
            section.different_first_page_header_footer = False
            section.header.is_linked_to_previous = True
            section.footer.is_linked_to_previous = True
        doc.save(doc_file)
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


if __name__ == '__main__':
    category_dirs_arr = ['一年级', '二年级', '三年级', '四年级', '五年级', '六年级']
    root_dir = "G:\\www.51zjedu.com\\www.51zjedu.com"
    category_dirs = sorted(os.listdir(root_dir))
    for category in category_dirs:
        if category in category_dirs_arr:
            files = sorted(os.listdir(root_dir + "\\" + category))
            for file in files:
                file_path = root_dir + "\\" + category + "\\" + file
                print(file_path)
                docx_dir = "G:\\www.51zjedu.com\\docx.51zjedu.com\\" + category
                if not os.path.exists(docx_dir):
                    os.makedirs(docx_dir)

                docx_file = docx_dir + "\\" + file.lower().replace(os.path.splitext(file)[1], ".docx")

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

                finish_dir = "G:\\finish.51zjedu.com" + "\\" + category
                if not os.path.exists(finish_dir):
                    os.makedirs(finish_dir)
                finish_file = docx_file.replace("docx.", "finish.").replace(".docx", ".pdf")

                if not os.path.exists(finish_file):
                    if os.path.exists(docx_file):
                        # 删除页眉页脚
                        if not remove_header_footer(docx_file):
                            continue

                    if os.path.exists(docx_file):
                        # 改变文档字体
                        if not change_word_font(docx_file):
                            continue

                    # 将docx转化为pdf
                    with open(finish_file, "w") as f:
                        # 将 Word 文档转换为 PDF
                        try:
                            convert(docx_file, finish_file)
                            print("转换成功！")
                        except Exception as e:
                            print("转换失败：", str(e))