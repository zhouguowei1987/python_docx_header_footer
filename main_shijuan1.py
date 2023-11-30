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
    category_dirs_arr = ['中考试卷', '高考试卷']
    root_dir = "G:\\www.shijuan1.com\\www.shijuan1.com"
    category_dirs = sorted(os.listdir(root_dir))
    for category in category_dirs:
        if category in category_dirs_arr:
            files = sorted(os.listdir(root_dir + "\\" + category))
            for file in files:
                file_path = root_dir + "\\" + category + "\\" + file
                print(file_path)

                # 查看一下是否是文件夹，如果是文件夹，则将文件移出
                # if os.path.isdir(file_path):
                #     child_files = sorted(os.listdir(file_path))
                #     for child_file in child_files:
                #         extension = os.path.splitext(child_file)[-1]
                #         if extension not in [".doc", ".docx"]:
                #             continue
                #         src_child_file_path = file_path + "\\" + child_file
                #         dst_child_file_path = root_dir + "\\" + category + "\\" + child_file
                #         os.rename(src_child_file_path, dst_child_file_path)

                docx_dir = "G:\\www.shijuan1.com\\docx.www.shijuan1.com" + "\\" + category
                if not os.path.exists(docx_dir):
                    os.makedirs(docx_dir)

                docx_file = docx_dir + "\\" + file.lower().replace(os.path.splitext(file)[1], ".docx")
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
                        continue

                    # 改变文档字体
                    if not change_word_font(docx_file):
                        continue
