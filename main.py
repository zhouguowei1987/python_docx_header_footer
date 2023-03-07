import time
from docx import Document
from win32com import client as wc
import os
import shutil


def remove_header_footer(doc, finish_doc):
    # doc：需要去页眉页脚的docx 文件
    # finish_doc： 需要另存为的新文件名
    document = Document(doc)
    for section in document.sections:
        section.different_first_page_header_footer = False
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True
    document.save(finish_doc)


def get_word_pages(in_file):
    pages = 1
    try:
        word = wc.Dispatch("Word.Application")
        try:
            doc = word.Documents.Open(in_file)
            word.ActiveDocument.Repaginate()
            pages = word.ActiveDocument.ComputeStatistics(2)
            print(pages)
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
        except Exception as e:
            print(e)
    except Exception as e:
        print(e)


if __name__ == '__main__':

    subject_dirs_arr = ['化学试卷', '历史试卷', '地理试卷', '数学试卷', '生物试卷', '语文试卷']
    root_dir = "G:\\final-www.shijuan1.com"
    subject_dirs = sorted(os.listdir(root_dir))
    for subject in subject_dirs:
        if subject in subject_dirs_arr:
            shijuan_dirs = sorted(os.listdir(root_dir + "\\" + subject))

            for shijuan in shijuan_dirs:
                word_dir = root_dir + "\\" + subject + "\\" + shijuan
                finish_dir = root_dir + "\\" + subject + "\\" + shijuan + "_finish"
                doc2docx_dir = root_dir + "\\" + subject + "\\" + shijuan + "_doc2docx"

                if not os.path.exists(finish_dir):
                    os.mkdir(finish_dir)

                if not os.path.exists(doc2docx_dir):
                    os.mkdir(doc2docx_dir)

                files = sorted(os.listdir(word_dir))
                for file in files:
                    if os.path.splitext(file)[1] in [".doc", ".docx"]:
                        print(file)
                        if os.path.splitext(file)[1] == ".docx":
                            # 将文件复制到doc2docx_dir目录
                            print("复制文件")
                            shutil.copyfile(word_dir + "\\" + file, doc2docx_dir + "\\" + file)
                        elif os.path.splitext(file)[1] == ".doc":
                            # 将doc文件转化为docx文件
                            print("转化文件")
                            doc2docx(word_dir + "\\" + file, doc2docx_dir + "\\" + os.path.splitext(file)[0] + ".docx")
                        # 去除word页眉和页脚
                        doc2docx_file = doc2docx_dir + "\\" + os.path.splitext(file)[0] + ".docx"
                        finish_doc = finish_dir + "\\" + os.path.splitext(file)[0] + ".docx"
                        if get_word_pages(doc2docx_file) >= 3:
                            remove_header_footer(doc2docx_file, finish_doc)
                            print("=============完成======================")