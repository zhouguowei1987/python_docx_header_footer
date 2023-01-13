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


word_dir = "G:\\www.ychedu.com\\母婴育儿"
finish_dir = "G:\\www.ychedu.com\\母婴育儿_finish"
doc2docx_dir = "G:\\www.ychedu.com\\母婴育儿_doc2docx"

if __name__ == '__main__':
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
                time.sleep(3)
            # 去除word页眉和页脚
            doc2docx_file = doc2docx_dir + "\\" + os.path.splitext(file)[0] + ".docx"
            finish_doc = finish_dir + "\\" + os.path.splitext(file)[0] + ".docx"
            remove_header_footer(doc2docx_file, finish_doc)
