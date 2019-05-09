from docx import Document
import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
import tkinter
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH

class fenqu_gui(object):

    def select_da_path(self):
        da_file = tkinter.filedialog.askopenfilename()
        self.da_path.set(da_file)

    def select_xiao_path(self):
        xiao_file = tkinter.filedialog.askopenfilename()
        self.xiao_path.set(xiao_file)

    def select_baogao_path(self):
        baogao_file = tkinter.filedialog.askopenfilename()
        self.baogao_path.set(baogao_file)

    def __init__(self):
        self.window = tkinter.Tk()

        self.da_path = tkinter.StringVar()
        self.xiao_path = tkinter.StringVar()
        self.baogao_path = tkinter.StringVar()

        self.window.title("中科院分区 by 北理工图书馆 不懂如山")

        self.da_label = tkinter.Label(self.window, text="大类文件:")
        self.da_input = tkinter.Entry(self.window, width=50, textvariable=self.da_path)
        self.da_path_button = tkinter.Button(self.window, text="选择文件", command=self.select_da_path)

        self.xiao_label = tkinter.Label(self.window, text="小类文件:")
        self.xiao_input = tkinter.Entry(self.window, width=50, textvariable=self.xiao_path)
        self.xiao_path_button = tkinter.Button(self.window, text="选择文件", command=self.select_xiao_path)

        self.baogao_label = tkinter.Label(self.window, text="报告文件:")
        self.baogao_input = tkinter.Entry(self.window, width=50, textvariable=self.baogao_path)
        self.baogao_path_button = tkinter.Button(self.window, text="选择文件", command=self.select_baogao_path)

        self.bgn_button = tkinter.Button(self.window, command=self.begin_fenqu, text="生成")

    def gui_arrange(self):

        self.da_label.grid(row=1, column=1)
        self.da_input.grid(row=1,column=2)
        self.da_path_button.grid(row=1, column=3)

        self.xiao_label.grid(row=2, column=1)
        self.xiao_input.grid(row=2, column=2)
        self.xiao_path_button.grid(row=2, column=3)

        self.baogao_label.grid(row=3, column=1)
        self.baogao_input.grid(row=3, column=2)
        self.baogao_path_button.grid(row=3, column=3)

        self.bgn_button.grid(row=4, column=2, sticky="e")

    def begin_fenqu(self):

        da_path = str(self.da_path.get()).strip()
        xiao_path = str(self.xiao_path.get()).strip()
        baogao_path = str(self.baogao_path.get()).strip()

        if (len(da_path)) < 1:
            tkinter.messagebox.showinfo("提示", "大类不能为空")
            return
        if (len(xiao_path)) < 1:
            tkinter.messagebox.showinfo("提示", "小类不能为空")
            return
        if (len(baogao_path)) < 1:
            tkinter.messagebox.showinfo("提示", "报告文件不能为空")
            return



        df_da = pd.read_excel(da_path,0)
        df_xiao = pd.read_excel(xiao_path,0)

        i = 0
        document = Document(baogao_path)

        p = document.add_paragraph("")
        run = p.add_run("1. 大类情况")
        run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
        run.bold = True

        table = document.add_table(rows=df_da.shape[0] + 1, cols=5)
        table.style = "Table Grid"
        table.autofit = True

        irows= 0
        icols = 0
        # 表格包括1.序号，2.文章题目 3.收录情况 4,引用情况  5.贡献情况
        hdr_cells = table.rows[irows + 0].cells
        hdr_cells[icols + 0].text = "期刊全称"
        hdr_cells[icols + 0].width = 5486400

        hdr_cells[icols + 1].text = "ISSN"
        hdr_cells[icols + 2].text = "所属大类"
        hdr_cells[icols + 3].text = "大类分区"
        hdr_cells[icols + 4].text = "Top期刊"

        irows += 1
        while i<df_da.shape[0]:

            title = df_da['期刊全称'][i]
            issn = df_da['ISSN'][i]
            dalei = df_da['所属大类'][i]
            dafenqu = df_da['大类分区'][i]
            datop = df_da['Top期刊'][i]

            hdr_cells = table.rows[irows].cells

            hdr_cells[0].text = title
            hdr_cells[1].text = issn
            hdr_cells[2].text = dalei
            hdr_cells[3].text = str(dafenqu)
            hdr_cells[4].text = datop

            i += 1
            irows += 1


        run = p.add_run("2. 小类情况")
        run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
        run.bold = True

        table = document.add_table(rows=df_xiao.shape[0] + 1, cols=5)
        table.style = "Table Grid"
        table.autofit = True

        irows = 0
        icols = 0
        # 表格包括1.序号，2.文章题目 3.收录情况 4,引用情况  5.贡献情况
        hdr_cells = table.rows[irows + 0].cells
        hdr_cells[icols + 0].text = "期刊全称"
        hdr_cells[icols + 0].width = 5486400

        hdr_cells[icols + 1].text = "ISSN"
        hdr_cells[icols + 2].text = "所属小类"
        hdr_cells[icols + 3].text = "小类分区"
        hdr_cells[icols + 4].text = "所属小类(中文)"

        i = 0
        irows += 1
        while i < df_xiao.shape[0]:
            title = df_xiao['期刊全称'][i]
            issn = df_xiao['ISSN'][i]
            xiaolei = df_xiao['所属小类'][i]
            xiaofenqu = df_xiao['小类分区'][i]
            xiaocn = df_xiao['所属小类(中文)'][i]

            hdr_cells = table.rows[irows].cells

            hdr_cells[0].text = title
            hdr_cells[1].text = issn
            hdr_cells[2].text = xiaolei
            hdr_cells[3].text = str(xiaofenqu)
            hdr_cells[4].text = xiaocn

            i += 1
            irows += 1

        document.save(baogao_path)
        tkinter.messagebox.showinfo("提示", "顺利完工！")
if __name__ == '__main__':
    gui = fenqu_gui()
    gui.gui_arrange()
    # 主程序执行
    tkinter.mainloop()