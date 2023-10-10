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
from win32com.client import constants
import os
import shutil


def check_header(doc):
    # doc：需要去页眉页脚的docx 文件
    document = Document(doc)
    if not document.sections[0].header.is_linked_to_previous:
        return True
    return False


def remove_header_footer(doc):
    # doc：需要去页眉页脚的docx 文件
    document = Document(doc)
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True
    document.save(doc)


def remove_last_image(doc_file):
    try:
        doc = Document(doc_file)
        paragraphs = doc.paragraphs
        # 最后一个段落
        para = paragraphs[len(paragraphs) - 1]
        p = para._element
        img = p.xpath('.//pic:pic')
        if img:
            # 最后一个段落是图片
            p.getparent().remove(p)
            p._p = p._element = None
        doc.save(doc_file)
    except Exception as e:
        print(e)
        return True
    return False


def delete_blank_page(doc_file):
    try:
        doc = Document(doc_file)
        for page in doc.page_margins:
            doc.remove_page(page)
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
    root_dir = "G:\\lvlin.baidu.com\\lvlin.baidu.com"
    files = sorted(os.listdir(root_dir))
    for file in files:
        file_path = root_dir + "\\" + file
        print(file_path)
        docx_dir = "G:\\lvlin.baidu.com\\docx.lvlin.baidu.com"
        if not os.path.exists(docx_dir):
            os.makedirs(docx_dir)

        docx_file = docx_dir + "\\" + file

        if not os.path.exists(docx_file):
            with open(docx_file, 'w') as f:
                pass
            print("==========开始转化为docx==============")
            if not doc2docx(file_path, docx_file):
                continue
            print("==========转化完成==============")
        else:
            # 已经是docx文件了，直接复制过去
            shutil.copy(file_path, docx_file)

        # 检测是否有页眉
        if check_header(docx_file):
            # 有页眉，检测最后一页是否是图片，如果是图片，则直接删除
            remove_last_image(docx_file)

        if os.path.exists(docx_file):
            # 删除并设置页眉页脚
            remove_header_footer(docx_file)
            # 删除空白页
            delete_blank_page(docx_file)
