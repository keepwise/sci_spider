import re
import urllib
import time
from datetime import datetime
import queue
from lxml import etree
import urllib.robotparser
import urllib.parse
import urllib.request
import sys
import io
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
#sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')

url = "http://apps.webofknowledge.com/full_record.do?product=UA&search_mode=MarkedList&qid=1&SID=8CYq2BiRX3op1mTZVOi&page=1&doc=1&colName=WOS"
request = urllib.request.Request(url=url, headers=headers)
opener = urllib.request.build_opener()

try:
    response = opener.open(request)
    #response.encoding = "utf-8"
    html = response.read().decode("UTF-8")
    # print(html)

    file = open(r"C:\Users\wxs\Desktop\py\test.htm","w",encoding='utf-8')
    file.write(html)
    file.close()

    code = response.code

except urllib.error.URLError as e:
    print('Download error: %s' % e.reason)
    html = ''
    if hasattr(e, 'code'):
        code = e.code

    else:
        code = None
except urllib.error.HTTPError as e:

    print("Download error: %s" % e.reason)