import tkinter
from tkinter import ttk

class Spider_gui(object):

    def select_path(self):
        rpt_dir = tkinter.filedialog.askdirectory()
        self.path.set(rpt_dir)

    def __init__(self):
        self.window = tkinter.Tk()

        self.path = tkinter.StringVar()
        self.jcr_opt = tkinter.BooleanVar()
        self.fenqu_opt = tkinter.BooleanVar()
        self.yinyong_opt = tkinter.BooleanVar()
        self.progress_value = tkinter.IntVar()

        self.jcr_opt.set(False)
        self.yinyong_opt.set(False)
        self.fenqu_opt.set(False)

        self.window.title("检索报告 by 北理工图书馆 不懂如山")
        self.window.iconbitmap("working.ico")

        self.url_label = tkinter.Label(self.window, text="URL:")
        self.url_input = tkinter.Entry(self.window, width=50)

        self.path_label = tkinter.Label(self.window, text="保存路径：")
        self.path_input = tkinter.Entry(self.window, width=50, textvariable=self.path)
        self.path_button = tkinter.Button(self.window, text="路径选择", command=self.select_path)

        self.bianhao_label = tkinter.Label(self.window, text="编号")
        self.bianhao_input = tkinter.Entry(self.window, width=10)

        self.JCR_checkbutton = tkinter.Checkbutton(self.window, text="JCR", onvalue=True, offvalue=False, width=15,
                                                   variable=self.jcr_opt)
        self.fenqu_checkbutton = tkinter.Checkbutton(self.window, text="中科院分区", onvalue=True, offvalue=False,
                                                     width=15, variable=self.fenqu_opt)
        self.yinyong_checkbutton = tkinter.Checkbutton(self.window,text="引用", onvalue=True, offvalue=False, width=15,variable=self.yinyong_opt)

        self.author_label = tkinter.Label(self.window,text="作者英文名：")
        self.author_input = tkinter.Entry(self.window,width=50)
        self.author_tip = tkinter.Label(self.window,text="(示例: Mao, ErKe)")


        self.progress_bar = tkinter.ttk.Progressbar(self.window,orient="horizontal", length=350, mode='determinate',variable=self.progress_value, maximum=100)
        self.processing_info = tkinter.Listbox(self.window, width=50)

        self.img_label = tkinter.Label(self.window)

        self.bgn_button = tkinter.Button(self.window, command=self.begin_crawl, text="开始")

    def gui_arrange(self):
        self.url_label.grid(row=1, column=1)
        self.url_input.grid(row=1, column=2)

        self.path_label.grid(row=2, column=1)
        self.path_input.grid(row=2, column=2)
        self.path_button.grid(row=2, column=3)

        self.bianhao_label.grid(row=3, column=1)
        self.bianhao_input.grid(row=3, column=2, sticky="w")
        self.JCR_checkbutton.grid(row=3, column=2 )
        self.fenqu_checkbutton.grid(row=3, column=2, sticky='e')
        self.yinyong_checkbutton.grid(row=3, column=3, sticky='e')

        self.author_label.grid(row=4, column=1)
        self.author_input.grid(row=4,column=2, stick="w")
        self.author_tip.grid(row=4,column=3,stick="w")

        self.progress_bar.grid(row=5,column=2)
        self.processing_info.grid(row=6, column=2)
        imgobj = tkinter.PhotoImage(file=r"working2.gif")
        self.img_label.config(image=imgobj)
        self.img_label.grid(row=6,column=3)
        self.img_label.image = imgobj
        self.bgn_button.grid(row=7, column=2, sticky="e")

    def begin_crawl(self):
        gui.bgn_button['state'] = 'disabled'




if __name__ == '__main__':
    gui = Spider_gui()
    gui.gui_arrange()
    # 主程序执行
    tkinter.mainloop()