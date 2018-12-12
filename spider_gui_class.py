import tkinter
import tkinter.filedialog
import threading
import scrape_sci


class Spider_gui(object):

    def select_path(self):
        rpt_dir = tkinter.filedialog.askdirectory()
        self.path.set(rpt_dir)


    def __init__(self):


        self.window = tkinter.Tk()

        self.path = tkinter.StringVar()
        self.jcr_opt = tkinter.BooleanVar()
        self.fenqu_opt = tkinter.BooleanVar()

        self.jcr_opt.set(False)
        self.fenqu_opt.set(False)

        self.window.title("检索报告采集")

        self.url_label = tkinter.Label(self.window, text="URL:")
        self.url_input = tkinter.Entry(self.window, width=50)

        self.path_label = tkinter.Label(self.window, text="保存路径：")
        self.path_input = tkinter.Entry(self.window, width=50, textvariable=self.path)
        self.path_button = tkinter.Button(self.window,text="路径选择",command=self.select_path)

        self.bianhao_label = tkinter.Label(self.window, text="编号")
        self.bianhao_input = tkinter.Entry(self.window, width = 10)

        self.JCR_checkbutton = tkinter.Checkbutton(self.window,text="JCR",onvalue=True,  offvalue=False, width=15, variable=self.jcr_opt)
        self.fenqu_checkbutton = tkinter.Checkbutton(self.window,text="中科院分区", onvalue=True, offvalue=False, width=15, variable=self.fenqu_opt)

        self.processing_info = tkinter.Listbox(self.window, width=50)

        self.bgn_button = tkinter.Button(self.window,command=self.begin_crawl, text="开始")


    def gui_arrange(self):

        self.url_label.grid(row=1,column=1)
        self.url_input.grid(row=1,column=2)

        self.path_label.grid(row=2,column=1)
        self.path_input.grid(row=2,column=2)
        self.path_button.grid(row=2, column=3)

        self.bianhao_label.grid(row=3,column=1)
        self.bianhao_input.grid(row=3,column=2,sticky="w")
        self.JCR_checkbutton.grid(row=3,column=2)
        self.fenqu_checkbutton.grid(row=3,column=2,sticky="e")

        self.processing_info.grid(row=4,column=2)
        self.bgn_button.grid(row=5,column=2,sticky="e")


    def begin_crawl(self):

        url = self.url_input.get()
        threading.start_new_thread(scrape_sci.scrape_sci(url))

if __name__=="__main__":

    sp_gui = Spider_gui()
    sp_gui.gui_arrange()

    tkinter.mainloop()



