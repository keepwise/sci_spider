import re
import urllib
import time
from datetime import datetime
import queue
from  lxml import etree
import urllib.robotparser
import urllib.parse
import urllib.request
import common
from docx import Document
import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
import tkinter
from tkinter import ttk
import threading
import tkinter.filedialog




headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "content-type": "text/html;charset=UTF-8",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:63.0) Gecko/20100101 Firefox/63.0"
}

url_crawled_num = 0  #爬取的原文数量
num_retries = 2   #爬取链接时尝试的次数
delay = 5   #请求间延迟的时间
proxy = None
wroten_original_num = 0  #已经写入的被引文献数量
document = Document()
document.styles["Normal"].font.name = "Times New Roman"
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
document.styles["Normal"].paragraph_format.space_after = docx.shared.Pt(1)

document.styles["Default Paragraph Font"].font.name = "Times New Roman"
document.styles['Default Paragraph Font']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def get_papers_queue(seed_url):

    global url_crawled_num, num_retries, delay, proxy,headers


    throttle = common.Throttle(delay)

    url_queue = queue.deque()
    html = common.download(url=seed_url,proxy=None,num_retries=num_retries,headers=headers)
    html_emt = etree.HTML(html)
    url_list = html_emt.xpath("//a[contains(@class,'snowplow-full-record')]/@href")
    for url in url_list:
        url_queue.append(url)
    next_page = html_emt.xpath("//a[contains(@class,'snowplow-navigation-nextpage-bottom')]/@href")

    try:
        if len(next_page) != 0:
            next_page = common.normalize(seed_url,next_page[0])
            throttle.wait(next_page)
            url_queue.extend(get_papers_queue(seed_url=next_page))
    except Exception as e:
        print("next_page: %s" % next_page)
        print(str(e))

    return url_queue

def get_paper_record(url):
    global url_crawled_num, num_retries, delay, proxy, headers

    paper = {}
    html = common.download(url=url, proxy=None, num_retries=num_retries,headers=headers)
    html_emt = etree.HTML(html)
    try:

        title = html_emt.xpath("//div[@class='title']/value/text()")[0]
        authors = html_emt.xpath("//div[@class='l-content']//p[@class='FR_field']/span[contains(text(),'By:') or contains(text(),'作者:')]/following-sibling::a/text()")
        fullNames = ""
        print(authors)

        for author in authors:

            fullName = html_emt.xpath("//p[@class='FR_field']/span[contains(text(),'By:') or contains(text(),'作者:')]/following-sibling::a[text()='%s']/following-sibling::text()" % author)[0]
            if fullName.find(";") != -1:

                fullNames = fullNames + author + fullName
            else:
                fullNames = fullNames + author + fullName + ";"

        print("FullNames:" + fullNames)

        SourceTitle = html_emt.xpath("//div[contains(@class,'block-record-info-source')]/p/span/value/text() | //div[contains(@class,'block-record-info-source')]/a/span/value/text() ")[0]

        SourceList = html_emt.xpath(
            "//div[contains(@class,'block-record-info-source')]/p[@class='FR_field']/node()/text() | //div[contains(@class,'block-record-info-source')]/p[@class='FR_field']/text()")

        source = SourceTitle + " "

        i = 0
        for ss in SourceList:
            if i % 2:
                source += str(ss).replace("\n", "") + " "
            else:
                source += str(ss).replace("\n", "")
            i = i + 1

        citing_url = html_emt.xpath("//a[@class='snowplow-citation-network-times-cited-count-link']/@href")
        if len(citing_url):
            citing_url = citing_url[0]
        else:
            citing_url = ""

        reprint_author_lst = html_emt.xpath("//div[@class='block-record-info']//span[contains(text(),'Reprint Address')]/following-sibling::text()")
        reprint_author_lst = filter(lambda x: "\n" not in str(x), reprint_author_lst)

        print(reprint_author_lst)
        reprint_addr_lst = html_emt.xpath("//table[@class='FR_table_noborders']/tr/td[2]/text()")

        rept_addr_author_lst = []
        rept_addr_author_lst = zip(reprint_author_lst,reprint_addr_lst)
        paper['reprint_author'] = ""

        for rpt in rept_addr_author_lst:
            if len(rpt)>1:
                paper['reprint_author'] += str(rpt[0]) + ", "
                paper['reprint_author'] += str(rpt[1]) + "\n"
            else:
                print("通讯作者错误: %s" % rept_addr_author_lst)
        print("reprint author: %s " % paper['reprint_author'])

        wos_cited_num = html_emt.xpath("//a[@class='snowplow-citation-network-times-cited-count-link']/span/text()")
        if len(wos_cited_num):
            wos_cited_num = wos_cited_num[0]
        else:
            wos_cited_num = "0"
        wos_no = html_emt.xpath("//input[@name='recordID']/@value")
        if len(wos_no):
            wos_no = wos_no[0]
        else:
            wos_no = ""
        data = html_emt.xpath("//div[@class='block-record-info']//span[contains(text(),'ISSN:')]/following-sibling::value/text()")
        if len(data):
            if len(data)==2:
                issn = data[0]
                eissn = data[1]
            if len(data)==1:
                issn = data[0]
                eissn = ""
        else:
            issn = ""
            eissn = ""

        paper['title'] = title
        paper['author'] = fullNames
        paper['source'] = source
        paper['citing_url'] = citing_url
        paper['wos_cited_num'] = wos_cited_num
        paper['wos_no'] = wos_no
        paper['issn']  = issn
        paper['eissn'] = eissn
    except Exception as e:

            print("获取文章详细信息失败,url: %s " % url)
            print("tite: %s" %  title)
            print("author: %s " % fullNames)
            print("source: %s" % source)
            print("错误：%s" % str(e))
    return paper


def write_shoulu(original_paper_lst):
    global document, wroten_original_num


    document.add_page_break()
    # 写入文件头
    p = document.add_paragraph("")
    p.add_run("附件二：").bold = True
    p = document.add_paragraph("")
    run = p.add_run("美国《科学引文索引》（SCI-EXPANDED）收录情况")
    run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
    run.bold = True

    style = document.styles.add_style("label",WD_STYLE_TYPE.CHARACTER)
    style.font.bold = True
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    style = document.styles.add_style("indent", WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.left_indent = docx.shared.Cm(0.5)
    style.paragraph_format.space_after = docx.shared.Pt(1)
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    i = 0
    for paper in original_paper_lst:
        i += 1
        p = document.add_paragraph("")
        p.add_run(str(i)+". Title: ",style="label")
        p.add_run(paper['title'])

        p = document.add_paragraph("",style="indent")
        p.add_run("Author(s): ",style="label")
        p.add_run(paper['author'])

        p = document.add_paragraph("", style="indent")
        p.add_run("Source: ", style="label")
        p.add_run(paper['source'])

        p = document.add_paragraph("", style="indent")
        p.add_run("Web of Science Core Collection引用次数: ", style="label")
        p.add_run(paper['wos_cited_num'])

        p = document.add_paragraph("", style="indent")
        p.add_run("入藏号: ", style="label")
        p.add_run(paper['wos_no'])

        p = document.add_paragraph("", style="indent")
        p.add_run("通讯作者: ", style="label")
        p.add_run(paper['reprint_author'])

        p = document.add_paragraph("", style="indent")
        p.add_run("ISSN: ", style="label")
        p.add_run(paper['issn'])

        p = document.add_paragraph("", style="indent")
        p.add_run("eISSN: ", style="label")
        p.add_run(paper['eissn'])

        document.save(str(gui.path_input.get()).strip()+"\\"+str(gui.bianhao_input.get()).strip()+".docx")

def write_word(record, record_type,cite_total=0, cur_cite=0):
    #record_type注明是原文还是引文

    global document, wroten_original_num
    if record_type == "original":
        if wroten_original_num == 0:

            #写入文件头
            p = document.add_paragraph("")
            p.add_run("附件一：").bold = True
            p = document.add_paragraph("")
            run = p.add_run("美国《Web of Science 核心合集：引文索引》（网络版）引用情况")
            run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
            run.bold = True
            # 设置引文段落格式
            style = document.styles.add_style("yinwen", WD_STYLE_TYPE.PARAGRAPH)
            style.paragraph_format.left_indent = docx.shared.Cm(1)
            style.paragraph_format.space_after = docx.shared.Pt(1)

            style.font.name = "Times New Roman"
            style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

            # 设置标题样式
            heading_style = document.styles.add_style("sci_heading", WD_STYLE_TYPE.CHARACTER)
            heading_style.font.bold = True
            heading_style.font.underline = True
            heading_style.font.name = "Times New Roman"
            heading_style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        wroten_original_num += 1
        p = document.add_paragraph("")
        p.add_run(str(wroten_original_num) + ". ")
        p.add_run("被引文献", style="sci_heading")
        p.add_run(":")

        p = document.add_paragraph("")
        p.add_run("Title: ").bold = True
        p.add_run(record['title'])

        p = document.add_paragraph("")
        p.add_run("Author(s): ").bold = True
        p.add_run(record['author'])

        p = document.add_paragraph("")
        p.add_run("Source: ").bold = True
        p.add_run(record['source'])

        p = document.add_paragraph("", style="yinwen")
        p.add_run("引用文献", style="sci_heading")
        p.add_run(": (自引").bold = True
        p.add_run("    ", style="sci_heading")
        p.add_run("篇  ").bold = True

        p.add_run("他引").bold = True
        p.add_run("    ", style="sci_heading")
        p.add_run("篇  )").bold = True
    if record_type == "citation":
        document.add_paragraph("Record %d of %d" % (cur_cite, cite_total), style="yinwen")
        p = document.add_paragraph("", style="yinwen")
        p.add_run("Title: ").bold = True
        p.add_run(record['title'])

        p = document.add_paragraph("", style="yinwen")
        p.add_run("Author(s): ").bold = True
        p.add_run(record['author'])

        p = document.add_paragraph("", style="yinwen")
        p.add_run("Source: ").bold = True
        p.add_run(record['source'])

    #print(str(gui.directory_input.get()).strip())
    document.save(str(gui.path_input.get()).strip() + "\\" + str(gui.bianhao_input.get()).strip()+".docx")

def scrape_sci(seed_url):

    orginal_papers_queue = get_papers_queue(seed_url)
    orginal_total = len(orginal_papers_queue)
    throttle = common.Throttle(delay)
    original_papers_lst = []
    procced_url_num = 0
    processed_origin_num = 0

    while orginal_papers_queue:

        original_url = orginal_papers_queue.popleft()
        paper = {}
        original_url = common.normalize(seed_url,original_url)
        throttle.wait(original_url)
        paper = get_paper_record(original_url)
        procced_url_num += 1
        processed_origin_num += 1
        gui.processing_info.insert(0,"正在处理 %d: %s" % (procced_url_num, original_url))
        gui.progress_value = round((procced_url_num/orginal_total)*100)

        print("progress_value: %d" % gui.progress_value)
        original_papers_lst.append(paper)

        if len(paper['citing_url'])>1:
            write_word(paper, record_type='original')
            citing_url = common.normalize(seed_url, paper['citing_url'])

            citing_papers_queue = get_papers_queue(citing_url)
            cite_total = len(citing_papers_queue)
            cur_cite = 1
            while citing_papers_queue:
                citation_url = citing_papers_queue.popleft()
                citation_url = common.normalize(seed_url, citation_url)

                throttle.wait(citation_url)
                citation = {}
                citation = get_paper_record(citation_url)
                procced_url_num += 1
                gui.processing_info.insert(0, "正在处理 %d: %s" % (procced_url_num, original_url))
                write_word(citation,record_type='citation',cite_total=cite_total, cur_cite=cur_cite)
                cur_cite += 1
    write_shoulu(original_papers_lst)
    gui.processing_info.insert(0, "顺利完成")

def scrape_jcr():
    pass

class Spider_gui(object):

    def select_path(self):
        rpt_dir = tkinter.filedialog.askdirectory()
        self.path.set(rpt_dir)

    def __init__(self):
        self.window = tkinter.Tk()

        self.path = tkinter.StringVar()
        self.jcr_opt = tkinter.BooleanVar()
        self.fenqu_opt = tkinter.BooleanVar()
        self.progress_value = tkinter.IntVar()

        self.jcr_opt.set(False)
        self.fenqu_opt.set(False)

        self.window.title("检索报告采集")

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

        self.progress_bar = tkinter.ttk.Progressbar(self.window,orient="horizontal", length=350, mode='determinate',variable=self.progress_value)
        self.processing_info = tkinter.Listbox(self.window, width=50)

        self.bgn_button = tkinter.Button(self.window, command=self.begin_crawl, text="开始")

    def gui_arrange(self):
        self.url_label.grid(row=1, column=1)
        self.url_input.grid(row=1, column=2)

        self.path_label.grid(row=2, column=1)
        self.path_input.grid(row=2, column=2)
        self.path_button.grid(row=2, column=3)

        self.bianhao_label.grid(row=3, column=1)
        self.bianhao_input.grid(row=3, column=2, sticky="w")
        self.JCR_checkbutton.grid(row=3, column=2)
        self.fenqu_checkbutton.grid(row=3, column=2, sticky="e")
        self.progress_bar.grid(row=4,column=2)
        self.processing_info.grid(row=5, column=2)
        self.bgn_button.grid(row=6, column=2, sticky="e")

    def begin_crawl(self):
        url = self.url_input.get()
        t1 = threading.Thread(target=scrape_sci,args=(url,))
        t1.start()



if __name__ == '__main__':
    gui = Spider_gui()
    gui.gui_arrange()
    # 主程序执行
    tkinter.mainloop()
