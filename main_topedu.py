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


def remove_header_footer(doc):
    # doc：需要去页眉页脚的docx 文件
    document = Document(doc)
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True
    document.save(doc)


def check_only_image(doc_file):
    try:
        doc = Document(doc_file)
        if len(doc.paragraphs) < 5:
            return True
        else:
            i = 0
            for para in doc.paragraphs:
                if i == 4 and para.text == "":
                    doc.save(doc_file)
                    return True
                i += 1
        doc.save(doc_file)
    except Exception as e:
        print(e)
        return True
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
    category_dirs_arr = ['高考真题', '中考真题']
    root_dir = "G:\\topedu.ybep.com.cn"
    category_dirs = sorted(os.listdir(root_dir))
    for category in category_dirs:
        if category in category_dirs_arr:
            files = sorted(os.listdir(root_dir + "\\" + category))
            for file in files:
                file_path = root_dir + "\\" + category + "\\" + file
                print(file_path)

                docx_dir = "G:\\docx.topedu.ybep.com.cn" + "\\" + category
                if not os.path.exists(docx_dir):
                    os.makedirs(docx_dir)

                docx_file = docx_dir + "\\" + file.replace(os.path.splitext(file)[1], ".docx")
                docx_file = docx_file.replace("（", "(").replace("）", ")")
                docx_file = docx_file.replace("[", "").replace("]", "")
                docx_file = docx_file.replace("(Word)", "")

                if not os.path.exists(docx_file):
                    print("==========开始转化为docx==============")
                    if not doc2docx(file_path, docx_file):
                        continue
                    print("==========转化完成==============")
                else:
                    # 已经是docx文件了，直接复制过去
                    shutil.copy(file_path, docx_file)

                # 删除并设置页眉页脚
                remove_header_footer(docx_file)

                finish_dir = "G:\\finish.topedu.ybep.com.cn" + "\\" + category
                if not os.path.exists(finish_dir):
                    os.makedirs(finish_dir)
                finish_file = docx_file.replace("docx.", "finish.").replace(".docx", ".pdf")

                if not os.path.exists(finish_file):
                    # 删除只包含图片
                    if check_only_image(docx_file):
                        continue
                    # 将docx转化为pdf
                    with open(finish_file, "w") as f:
                        # 将 Word 文档转换为 PDF
                        try:
                            convert(docx_file, finish_file)
                            print("转换成功！")
                        except Exception as e:
                            print("转换失败：", str(e))
