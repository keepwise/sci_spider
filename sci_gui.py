import tkinter
import scrape_sci

class sci_scrape_gui:

    def __init__(self):

        self.root = tkinter.Tk()
        self.root.title("定制检索报告")
        self.url_label = tkinter.Label(self.root,text="URL地址：")
        self.url_input = tkinter.Entry(self.root, width=100)
        self.rpt_label = tkinter.Label(self.root, text="报告编号：")
        self.rpt_input = tkinter.Entry(self.root, width=100)
        self.author_label = tkinter.Label(self.root, text="作者英文名：")
        self.author_input = tkinter.Entry(self.root, width=100)
        self.directory_label = tkinter.Label(self.root, text="保存地址：")
        self.directory_input = tkinter.Entry(self.root, width=100)

        self.download_info = tkinter.Listbox(self.root,width=110)

        self.confirm_btn = tkinter.Button(self.root, command= self.begin_scrape, text = "开始")

    def gui_arrang(self):
        #self.download_info.pack()
        #self.confirm_btn.pack()
        self.url_label.grid(row=1, column=1, padx=5, pady=5)
        self.url_input.grid(row=1, column=2, padx=5, pady=5)
        self.rpt_label.grid(row=2, column=1, padx=5, pady=5)
        self.rpt_input.grid(row=2, column=2, padx=5, pady=5)
        self.author_label.grid(row=3, column=1, padx=5, pady=5)
        self.author_input.grid(row=3, column=2, padx=5, pady=5)
        self.directory_label.grid(row=4, column=1, padx=5, pady=5)
        self.directory_input.grid(row=4, column=2, padx=5, pady=5)
        self.download_info.grid(row=5, column=2, padx=5, pady=5)
        self.confirm_btn.grid(row=6, column=2, padx=5, pady=5)
    def begin_scrape(self):
        self.seed_url = str(self.url_input.get()).strip()
        self.author = str(self.author_input.get()).strip()
        path = str(self.url_input.get()).strip()
        scrape_sci.scrape_sci(seed_url=self.seed_url)

def main():
    # 初始化对象
    gui = sci_scrape_gui()
    # 进行布局
    gui.gui_arrang()
    # 主程序执行
    tkinter.mainloop()


if __name__ == "__main__":
    main()
