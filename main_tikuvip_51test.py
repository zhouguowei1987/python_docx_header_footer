import re
import time
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from win32com import client as wc
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


def get_word_pages(in_file):
    pages = 1
    try:
        word = wc.Dispatch("Word.Application")
        try:
            doc = word.Documents.Open(in_file)
            word.ActiveDocument.Repaginate()
            pages = word.ActiveDocument.ComputeStatistics(2)
            doc.Close()
            word.Quit()
            return pages
        except Exception as e:
            print(e)
        finally:
            return pages
    except Exception as e:
        print(e)
    finally:
        return pages


def docx_remove_content(doc_file):
    # 定义需要去除的内容
    content_to_remove = '''XXXXXX'''
    # 打开doc文件
    doc = Document(doc_file)
    # 遍历doc文件中的段落
    for para in doc.paragraphs:
        # 如果段落中包含需要去除的内容，使用正则表达式替换为空字符串
        if re.search(content_to_remove, para.text):
            para.text = re.sub(content_to_remove, 'OfficePLUS', para.text)

    doc.save(doc_file)


def change_word_font(doc_file):
    # 打开doc文件
    doc = Document(doc_file)
    doc.styles['Normal'].font.name = u'Times New Roman'  # 设置西文字体
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'微软雅黑')  # 设置中文字体使用字体2->宋体
    doc.save(doc_file)


def check_only_image(doc_file):
    doc = Document(doc_file)
    i = 0
    for para in doc.paragraphs:
        if i == 0 and para.text == "":
            doc.save(doc_file)
            return True
        i += 1
    doc.save(doc_file)
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
    category_dirs_arr = ['自考', '专升本考试', '一级建造师考试', '小升初', '考研', '公务员考试', '高中会考', '高考',
                         '二级建造师考试', '成人高考']
    root_dir = "G:\\tikuvip.51test.net"
    category_dirs = sorted(os.listdir(root_dir))
    for category in category_dirs:
        if category in category_dirs_arr:
            files = sorted(os.listdir(root_dir + "\\" + category))
            for file in files:
                if os.path.splitext(file)[1] == ".doc":
                    file_path = root_dir + "\\" + category + "\\" + file
                    print(file_path)
                    if "答案" not in file_path:
                        continue

                    docx_dir = "G:\\docx.tikuvip.51test.net" + "\\" + category
                    finish_dir = "G:\\finish.tikuvip.51test.net" + "\\" + category
                    if not os.path.exists(docx_dir):
                        os.mkdir(docx_dir)

                    if not os.path.exists(finish_dir):
                        os.mkdir(finish_dir)
                    docx_file = docx_dir + "\\" + file.replace(".doc", ".docx")
                    finish_file = finish_dir + "\\" + file.replace(".doc", ".docx")
                    if not os.path.exists(finish_file):
                        if not os.path.exists(docx_file):
                            print("==========开始转化为docx==============")
                            if not doc2docx(file_path, docx_file):
                                continue
                            print("==========转化完成==============")
                    # 删除只包含图片
                    if check_only_image(docx_file):
                        continue
                    # 删除页眉页脚
                    remove_header_footer(docx_file, finish_file)
                    # 改变文档字体
                    change_word_font(finish_file)

