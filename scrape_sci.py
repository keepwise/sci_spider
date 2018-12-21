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
import tkinter.messagebox
import pymysql
import pandas as pd



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

original_papers_lst = []
cur_original_paper_no = 0

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

def get_paper_record(url,database):
    global url_crawled_num, num_retries, delay, proxy, headers

    paper = {}
    html = common.download(url=url, proxy=None, num_retries=num_retries,headers=headers)
    html = html.encode(encoding='utf-8')
    html = html.decode("utf-8")

    file = open(r"C:\Users\wangxiaoshan\Desktop\wxs_py\test.html", "w", encoding="utf-8")
    file.write(html)
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

        reprint_author_lst = html_emt.xpath("//div[@class='block-record-info']//span[contains(text(),'Reprint Address')]/following-sibling::text()| //div[@class='block-record-info']//span[contains(text(),'通讯作者地址')]/following-sibling::text()")
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
        #自引次数
        paper['ziyin'] = 0
        #他引次数
        paper['tayin'] = 0
        paper['shoulu'] = database
    except Exception as e:

            print("获取文章详细信息失败,url: %s " % url)
            print("tite: %s" %  title)
            print("author: %s " % fullNames)
            print("source: %s" % source)
            print("错误：%s" % str(e))
    return paper


def write_shoulu(original_paper_lst):

    global document, wroten_original_num

    # 写入文件头
    p = document.add_paragraph("")
    p.add_run("附件一：").bold = True
    p = document.add_paragraph("")
    #run = p.add_run("美国《科学引文索引》（SCI-EXPANDED）收录情况")
    run = p.add_run("收录详细情况")
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

        if paper.get('wos_no','not') != 'not':
            p = document.add_paragraph("", style="indent")
            p.add_run("入藏号: ", style="label")
            p.add_run(paper['wos_no'])

        if paper.get('accession number','not') !='not':
            p = document.add_paragraph("", style="indent")
            p.add_run("Accession number: ", style="label")
            p.add_run(paper['accession number'])

        p = document.add_paragraph("", style="indent")
        p.add_run("通讯作者: ", style="label")
        p.add_run(paper['reprint_author'])

        p = document.add_paragraph("", style="indent")
        p.add_run("ISSN: ", style="label")
        p.add_run(paper['issn'])

        p = document.add_paragraph("", style="indent")
        p.add_run("eISSN: ", style="label")
        p.add_run(paper['eissn'])

        p = document.add_paragraph("", style="indent")
        p.add_run("收录情况: ", style="label")
        p.add_run(paper['shoulu'])

        document.save(str(gui.path_input.get()).strip()+"\\"+str(gui.bianhao_input.get()).strip()+"_shoulu.docx")

def write_word(record, record_type,cite_total=0, cur_cite=0):
    #record_type注明是原文还是引文

    global document, wroten_original_num, original_papers_lst, cur_original_paper_no

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
        yinyong_paragraph_position = len(document.paragraphs)-1

        p.add_run("引用文献", style="sci_heading")
        p.add_run(": (自引").bold = True
        p.add_run("    ", style="sci_heading")
        p.add_run("篇  ").bold = True
        p.add_run("他引").bold = True
        p.add_run("    ", style="sci_heading")
        p.add_run("篇  )").bold = True

        original_papers_lst[cur_original_paper_no]['yinyong_paragraph_position'] = yinyong_paragraph_position
    if record_type == "citation":
        ziyin_tayin(record,cite_total,cur_cite)

    #print(str(gui.directory_input.get()).strip())
    document.save(str(gui.path_input.get()).strip() + "\\" + str(gui.bianhao_input.get()).strip()+".docx")

def ziyin_tayin(citation,cite_total=0, cur_cite=0):
    '''区分自引他引，并写入word文档'''
    global original_papers_lst,cur_original_paper_no

    original_authors = original_papers_lst[cur_original_paper_no]['author']
    original_authors_lst = original_authors.split(";")
    original_authors_lst = [str(x).strip() for x in original_authors_lst]
    #生成类似下述列表 [0]="Wang, XY (Wang, Xinyu)"   [1]="Fu, MY (Fu, Mengyin)" [2] = "Ma, HB (Ma, Hongbin)"

    citation_authors = citation['author']
    original_paper = original_papers_lst[cur_original_paper_no]

    p = document.add_paragraph("", style="yinwen")
    p.add_run("Record %d of %d" % (cur_cite, cite_total)).bold = True
    p = document.add_paragraph("", style="yinwen")
    p.add_run("Title: ").bold = True
    p.add_run(citation['title'])
    p = document.add_paragraph("", style="yinwen")
    p.add_run("Author(s): ").bold = True

    for full_author in original_authors_lst:
        # 如果[0]="Wang, XY (Wang, Xinyu)"   [1]="Fu, MY (Fu, Mengyin)" [2] = "Ma, HB (Ma, Hongbin)"匹配
        if len(full_author)>1 and str(citation_authors).find(full_author) != -1:
            author_len = len(full_author)
            author_start = str(citation_authors).find(full_author)
            p.add_run(citation['author'][0:author_start])
            p.add_run(citation['author'][author_start:(author_start+author_len)]).font.highlight_color = WD_COLOR_INDEX.RED
            p.add_run(citation['author'][(author_start + author_len):])

            original_papers_lst[cur_original_paper_no]['ziyin'] += 1
            break
    else:
        #如果[0]="Wang, XY (Wang, Xinyu)"   [1]="Fu, MY (Fu, Mengyin)" [2] = "Ma, HB (Ma, Hongbin)"不匹配
        # 就用 (Wang, Xinyu) ， (Fu, Mengyin)，(Ma, Hongbin)尝试匹配，如果匹配 就算自引一次
        for full_author in original_authors_lst:
            if str(full_author).find("(") != -1:
                full_name = str(full_author).split("(")[1]
                #full_name = (Wang, Xinyu)
                full_name = "(" + full_name.strip()
            else:
                full_name = str(full_author).strip()
            if len(full_name)>1 and str(citation_authors).find(full_name) != -1:
                author_len = len(full_name)
                author_start = str(citation_authors).find(full_name)
                p.add_run(citation['author'][0:author_start])
                p.add_run(citation['author'][author_start:(author_start + author_len)]).font.highlight_color = WD_COLOR_INDEX.RED
                p.add_run(citation['author'][(author_start + author_len):])

                original_papers_lst[cur_original_paper_no]['ziyin'] += 1
                break
            else:
                #如果full_name中含有",",如 Wang, Xinyu
                if full_name.find(",") != -1:
                    full_name = full_name.replace(",", "")
                    #对 Wang Xinyu进行检索
                    if str(citation_authors).find(full_name) != -1:
                        author_len = len(full_name)
                        author_start = str(citation_authors).find(full_name)
                        p.add_run(citation['author'][0:author_start])
                        p.add_run(citation['author'][author_start:(author_start + author_len)]).font.highlight_color = WD_COLOR_INDEX.RED
                        p.add_run(citation['author'][(author_start + author_len):])

                        original_papers_lst[cur_original_paper_no]['ziyin'] += 1
                        break
        else:
            # 如果[0]="Wang, XY (Wang, Xinyu)"   [1]="Fu, MY (Fu, Mengyin)" [2] = "Ma, HB (Ma, Hongbin)"不匹配
            # 如果 (Wang, Xinyu) ， (Fu, Mengyin)，(Ma, Hongbin)也不匹配
            #使用 Wang, XY,    Fu, MY, Ma, HB 进行匹配，两次匹配以上可以确认一次自引
            basic_matched_num = 0
            for full_author in original_authors_lst:
                if str(full_author).find("(") != -1:
                    basic_name = str(full_author).split("(")[0]
                    basic_name = basic_name.strip()
                    if str(citation_authors).find(basic_name) != -1:
                        author_len = len(basic_name)
                        author_start = str(citation_authors).find(basic_name)
                        p.add_run(citation['author'][0:author_start])
                        p.add_run(citation['author'][author_start:(author_start + author_len)]).font.highlight_color = WD_COLOR_INDEX.GREEN
                        p.add_run(citation['author'][(author_start + author_len):])

                        basic_matched_num += 1

                        if basic_matched_num>1:
                            original_papers_lst[cur_original_paper_no]['ziyin'] += 1
                            break
            else:
                p.add_run(citation['author'])

    p = document.add_paragraph("", style="yinwen")
    p.add_run("Source: ").bold = True
    p.add_run(citation['source'])

    original_papers_lst[cur_original_paper_no]['tayin'] = cite_total - original_papers_lst[cur_original_paper_no]['ziyin']
    yinyong_paragraph_position = original_papers_lst[cur_original_paper_no]['yinyong_paragraph_position']

    p = document.paragraphs[yinyong_paragraph_position]
    p = p.clear()
    p.add_run("引用文献", style="sci_heading")
    p.add_run(": (自引").bold = True
    p.add_run(" %d " % original_papers_lst[cur_original_paper_no]['ziyin'], style="sci_heading")
    p.add_run("篇  ").bold = True
    p.add_run("他引").bold = True
    p.add_run(" %d " % original_papers_lst[cur_original_paper_no]['tayin'], style="sci_heading")
    p.add_run("篇  )").bold = True




def scrape_sci(seed_url):

    global original_papers_lst
    orginal_papers_queue = get_papers_queue(seed_url)
    orginal_total = len(orginal_papers_queue)
    throttle = common.Throttle(delay)

    procced_url_num = 0
    processed_origin_num = 0

    while orginal_papers_queue:

        global cur_original_paper_no
        original_url = orginal_papers_queue.popleft()
        paper = {}
        original_url = common.normalize(seed_url,original_url)
        throttle.wait(original_url)
        paper = get_paper_record(original_url,"SCIE")
        procced_url_num += 1
        processed_origin_num += 1
        gui.processing_info.insert(0,"正在处理 %d: %s" % (procced_url_num, original_url))

        #gui.progress_bar.update()
        print("progress_value: %d" % gui.progress_value.get())
        original_papers_lst.append(paper)

        if len(paper['citing_url'])>1 and gui.yinyong_opt.get():
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
                citation = get_paper_record(citation_url, "SCIE")
                procced_url_num += 1
                gui.processing_info.insert(0, "正在处理 %d: %s" % (procced_url_num, original_url))
                write_word(citation,record_type='citation',cite_total=cite_total, cur_cite=cur_cite)
                cur_cite += 1
        cur_original_paper_no += 1
        gui.progress_value.set((procced_url_num / orginal_total) * 100)


    shoulu_document = Document()
    shoulu_document.styles["Normal"].font.name = "Times New Roman"
    shoulu_document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    shoulu_document.styles["Normal"].paragraph_format.space_after = docx.shared.Pt(1)

    shoulu_document.styles["Default Paragraph Font"].font.name = "Times New Roman"
    shoulu_document.styles['Default Paragraph Font']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    author_contribution(shoulu_document)
    write_shoulu(original_papers_lst)
    ei_file = gui.ei_file_path.get()
    if(len(ei_file.strip())>1):
        write_ei_shoulu(shoulu_document)

    if gui.jcr_opt.get() == True:
        scrape_jcr(shoulu_document)
    if gui.fenqu_opt.get() == True:
        scrape_fenqu(shoulu_document)
    gui.processing_info.insert(0, "顺利完成")
    gui.bgn_button['state']= 'normal'
    tkinter.messagebox.showinfo("提示","主人，活儿干完啦！")

def ei_shoulu(document):

    global original_papers_lst
    ei_file = gui.ei_file_path.get()
    ei_file = ei_file.strip()
    ei_papers = pd.read_excel(ei_file,0)
    i = 0
    while i<ei_papers.shape[0]:
        title = ei_papers['Title'][i]
        ei_paper = {}
        j = 0
        while j<len(original_papers_lst):
            if title == original_papers_lst[j]['title']:
                original_papers_lst[j]['accession number'] =  ei_papers['Accession number'][i]
                original_papers_lst[j]['shoulu'] += ", EI"
                break
            j += 1
        else:
            ei_paper['title'] = title
            ei_paper['author'] = ei_papers['Author'][i]
            ei_paper['source'] = ei_papers['Source'][i]+" 卷:"+ei_papers['Volume'][i]+"     页:"+ei_papers['Pages'][i] +" 出版年:"+ei_papers['Publication year'][i]+" 文献类型:"+ei_papers['Document type'][i]
            ei_paper['issn'] = ei_papers['ISSN']
            ei_paper['accession number'] = ei_papers['Accession number'][i]
            ei_paper['shoulu'] = "EI"
            original_papers_lst.append(ei_paper)

        i += 1

def scrape_jcr(document):
    global original_papers_lst

    document.add_page_break()
    # 写入文件头
    p = document.add_paragraph("")
    p.add_run("附件二：").bold = True
    p = document.add_paragraph("")
    run = p.add_run("美国《期刊引用报告》（JCR）收录情况")
    run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
    run.bold = True

    #期刊的不重复列表
    issn_lst = []
    for paper in original_papers_lst:
        if paper['issn'] not in issn_lst:
            issn_lst.append(paper['issn'])
    try:
        db = pymysql.connect(host="118.31.57.201", db="chaxinbu", user="db_kiaora", password="Kiaora477(")
        cursor = db.cursor()
        i = 1
        for issn in issn_lst:
            cursor.execute("select * from jcr where issn='%s'" % issn)

            result = cursor.fetchone()
            if result is not None:
                title = result[1]
                jif = result[2]
                document.add_paragraph("%d. 刊名：%s" %(i,title))
                document.add_paragraph("ISSN: %s" % issn)
                document.add_paragraph("2017年影响因子: %s" % jif)
                i += 1

    except Exception as e:
        print("JCR收录出错:%s" % str(e))
    finally:
        document.save(str(gui.path_input.get()).strip() + "\\" + str(gui.bianhao_input.get()).strip() + "_shoulu.docx")

def scrape_fenqu(document):
    global original_papers_lst

    document.add_page_break()
    # 写入文件头
    p = document.add_paragraph("")
    run = p.add_run("中科院SCI分区表（网络版）收录情况")
    run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
    run.bold = True

    # 期刊的不重复列表
    issn_lst = []
    for paper in original_papers_lst:
        if paper['issn'] not in issn_lst:
            issn_lst.append(paper['issn'])

    i = 1
    for issn in issn_lst:

        document.add_paragraph("%d. 刊名：" % i)
        document.add_paragraph("  ISSN：%s" % issn)
        document.add_paragraph("  2017版分区情况")

        table = document.add_table(rows=3, cols=4)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = ""
        hdr_cells[1].text = "学科名称"
        hdr_cells[2].text = "分区"
        hdr_cells[3].text = "TOP期刊"


        table.rows[1].cells[0].text = "小类"
        table.rows[2].cells[0].text = "大类"

        i += 1
    document.save(str(gui.path_input.get()).strip() + "\\" + str(gui.bianhao_input.get()).strip() + "_shoulu.docx")

def author_contribution(document):
    '''查询作者是否第一作者、通讯作者'''
    global original_papers_lst

    if len(document.paragraphs)>1:
        del document.paragraphs[:]
    author = gui.author_input.get()
    name_lst = author.split(",")
    if len(name_lst)<1:
        tkinter.messagebox.showinfo("提示","作者格式有误")
        return
    else:
        p = document.add_paragraph("")
        run = p.add_run("作者论文情况概览")
        run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
        run.bold = True

        table = document.add_table(rows=len(original_papers_lst)+1, cols=5)
        table.style = "Table Grid"
        table.autofit = True

        #表格包括1.序号，2.文章题目 3.收录情况 4,引用情况  5.贡献情况
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "序号"
        hdr_cells[1].text = "标题"
        hdr_cells[2].text = "收录情况"
        hdr_cells[3].text = "引用情况"
        hdr_cells[4].text = "贡献情况"

        capital_pattern = re.compile("[A-Z]")

        capital_lst = capital_pattern.findall(str(name_lst[1]))
        capital_name = ""
        for capital in capital_lst:
            capital_name += capital+"[\.-]?"
        #生成正则"Mao[\s,]+E[\.-]?K[\.-]?"
        pattern_str = str(name_lst[0]).strip() + "[\s,]+" + capital_name
        paper_num = 1
        for paper in original_papers_lst:
            #匹配是否第一作者
            matchObj = re.match(pattern_str,paper['author'], re.M)
            if matchObj:
                paper['bool_first_author'] = True
            else:
                paper['bool_first_author'] = False

            #匹配是否通讯作者
            matchObj = re.search(pattern_str,paper['reprint_author'],re.M)
            if matchObj:
                paper['bool_reprint_author'] = True
            else:
                paper['bool_reprint_author'] = False

            paper_cells = table.rows[paper_num].cells
            paper_cells[0].text = str(paper_num)
            paper_cells[1].text = paper['title']
            if paper.get("wos",'not') != 'not':
                paper_cells[1].text += "\n WOS:" + paper['wos']
            if paper.get('accession number','not') != 'not':
                paper_cells[1].text += "\n Accession Number:" + paper['accession number']

            paper_cells[2].text = paper['shoulu']
            paper_cells[3].text = "自引%d次，他引%d次" %(paper['ziyin'],paper['tayin'])
            paper_cells[4].text = ""
            if paper['bool_first_author']:
                paper_cells[4].text += "第一作者  "
            if paper['bool_reprint_author']:
                paper_cells[4].text +="通讯作者"
            paper_num += 1

    document.save(str(gui.path_input.get()).strip() + "\\" + str(gui.bianhao_input.get()).strip() + "_shoulu.docx")

class Spider_gui(object):

    def select_path(self):
        rpt_dir = tkinter.filedialog.askdirectory()
        self.path.set(rpt_dir)

    def select_ei_path(self):
        ei_file = tkinter.filedialog.askopenfilename()
        self.path.set(ei_file)

    def __init__(self):
        self.window = tkinter.Tk()

        self.path = tkinter.StringVar()
        self.jcr_opt = tkinter.BooleanVar()
        self.fenqu_opt = tkinter.BooleanVar()
        self.yinyong_opt = tkinter.BooleanVar()
        self.progress_value = tkinter.IntVar()
        self.ei_file_path = tkinter.StringVar()

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

        self.ei_label = tkinter.Label(self.window, text="EI文件：")
        self.ei_path_input = tkinter.Entry(self.window, width=50, textvariable=self.ei_file_path)
        self.ei_path_button = tkinter.Button(self.window, text="选择文件", command=self.select_ei_path)

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

        self.bgn_button = tkinter.Button(self.window, command=self.begin_crawl, text="开始")

    def gui_arrange(self):
        self.url_label.grid(row=1, column=1)
        self.url_input.grid(row=1, column=2)

        self.path_label.grid(row=2, column=1)
        self.path_input.grid(row=2, column=2)
        self.path_button.grid(row=2, column=3)

        self.ei_path_label.grid(row=3, column=1)
        self.ei_path_input.grid(row=3, column=2)
        self.ei_path_button.grid(row=3, column=3)

        self.bianhao_label.grid(row=4, column=1)
        self.bianhao_input.grid(row=4, column=2, sticky="w")
        self.JCR_checkbutton.grid(row=4, column=2 )
        self.fenqu_checkbutton.grid(row=4, column=2, sticky='e')
        self.yinyong_checkbutton.grid(row=4, column=3, sticky='e')

        self.author_label.grid(row=5, column=1)
        self.author_input.grid(row=5,column=2, stick="w")
        self.author_tip.grid(row=5,column=3,stick="w")

        self.progress_bar.grid(row=6,column=2)
        self.processing_info.grid(row=7, column=2)
        self.bgn_button.grid(row=8, column=2, sticky="e")

    def begin_crawl(self):
        gui.bgn_button['state'] = 'disabled'
        url = self.url_input.get()
        t1 = threading.Thread(target=scrape_sci,args=(url,))
        t1.start()



if __name__ == '__main__':
    gui = Spider_gui()
    gui.gui_arrange()
    # 主程序执行
    tkinter.mainloop()
