from pptx import Presentation
import os


if __name__ == '__main__':
    root_dir = "G:\\aaa"
    files = sorted(os.listdir(root_dir))
    for file in files:
        file_path = root_dir + "\\" + file
        print(file_path)

        prs = Presentation(file_path)

        slide_master = prs.slide_masters[1]
        print(slide_master.name)
        # slide_master.name = "New Name"

        # slide_master = prs.slide_masters[0]
        # print(slide_master.name)
        #
        layout = prs.slide_layouts[0]
        layout.name = "ababababa"
        print(layout.name)
        prs.save(root_dir + "\\" + "电子奖状模板.pptx")
        # print(prs.slide_masters)
        # slide_master = prs.slide_masters[0]
        # print(slide_master.name)
        # for slide_master in prs.slide_masters:
        #     print(slide_master.name)


