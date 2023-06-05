import re
import time
from docx import Document
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor  # 设置字体的颜色
from docx.oxml.ns import qn
import os
import shutil


def remove_header_footer(doc, save_doc):
    # doc：需要去页眉页脚的docx 文件
    # finish_doc： 需要另存为的新文件名
    document = Document(doc)
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True
    document.save(save_doc)


def remove_links(doc_file):
    doc = Document(doc_file)
    for p in doc.paragraphs:
        for run in p.runs:
            if run.hyperlink:
                run.text = run.text.replace(run.text, '')
    doc.save(doc_file)


def docx_remove_content(doc_file):
    # 定义需要去除的内容
    content_to_remove = '''不用注册，免费下载！'''
    # 打开doc文件
    doc = Document(doc_file)
    # 遍历doc文件中的段落
    for para in doc.paragraphs:
        # 如果段落中包含需要去除的内容，使用正则表达式替换为空字符串
        if re.search(content_to_remove, para.text):
            para.text = re.sub(content_to_remove, '', para.text)

    doc.save(doc_file)


def change_word_font(doc_file):
    # 打开doc文件
    doc = Document(doc_file)
    doc.styles['Normal'].font.name = u'Times New Roman'  # 设置西文字体
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')  # 设置中文字体使用字体2->宋体

    # 修改标题字体
    para = doc[1].paragraphs
    for run in para.runs:
        # run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = RGBColor(255, 0, 0)
    doc.save(doc_file)


def change_line_spacing(doc_file):
    doc = Document(doc_file)
    for p in doc.paragraphs:  # 循环处理每个段落
        p.paragraph_format.line_spacing = 1.5  # 行距设置为3
    doc.save(doc_file)


if __name__ == '__main__':
    root_dir = "../www.hi138.com"
    files = sorted(os.listdir(root_dir))
    for file in files:
        if os.path.splitext(file)[1] == ".docx":
            file_path = root_dir + "/" + file
            print(file_path)
            exit()
            file_finish_dir = "finish.hi138.com"
            if not os.path.exists(file_finish_dir):
                os.makedirs(file_finish_dir)
            file_finish_path = file_finish_dir + file

            try:
                # 删除页眉页脚
                remove_header_footer(file_path, file_finish_path)

                # 删除文档中链接
                remove_links(file_finish_path)

                # 改变文档字体
                change_word_font(file_finish_path)

                # 修改行距
                change_line_spacing(file_finish_path)
            except Exception as e:
                print(e)
