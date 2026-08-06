import re
import time
from docx import Document
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor  # 设置字体的颜色
from docx.oxml.ns import qn
from win32com import client as wc
import os
import shutil


def remove_header_footer(doc):
    time.sleep(10)
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
    time.sleep(10)
    try:
        word = wc.Dispatch("Word.Application")
        doc = word.Documents.Open(in_file,  ReadOnly=1)
        # word.Visible = True  # 显示Word界面
        doc.SaveAs(out_file, 12, False, "", True, "", False, False, False, False)
        print('转换成功')
        doc.Close()
        word.Quit()  # 确保Word被关闭，无论是否发生异常
        return True
    except Exception as e:
        print(e)
        return False
    finally:
        del word  # 释放资源


if __name__ == '__main__':
    root_dir = "F:\\workspace\\shigongdaquan.max.book118.com\\shigongdaquan.max.book118.com"
    files = sorted(os.listdir(root_dir))
    for file in files:
        file_path = root_dir + "\\" + file
        print(file_path)
        docx_dir = "F:\\workspace\\shigongdaquan.max.book118.com\\docx.shigongdaquan.max.book118.com"
        if not os.path.exists(docx_dir):
            os.makedirs(docx_dir)

        docx_file = docx_dir + "\\" + file.replace(os.path.splitext(file)[1], ".docx")
        print(docx_file)

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
                    # 删除原文件
                    try:
                        os.remove(file_path)
                        os.remove(docx_file)
                        continue
                    except Exception as e:
                        print(e)
                print("==========转化完成==============")
            if os.path.exists(docx_file):
                # 删除页眉页脚
                print("==========开始删除页眉页脚==============")
                if not remove_header_footer(docx_file):
                    # 删除原文件
                    try:
                        os.remove(file_path)
                        os.remove(docx_file)
                        continue
                    except Exception as e:
                        print(e)
                print("==========完成删除页眉页脚==============")