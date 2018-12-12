import re
import urllib
import time
from datetime import datetime
import queue
from  lxml import etree
import urllib.robotparser
import urllib.parse
import urllib.request

url_crawled_num = 0
def link_crawler(seed_url, link_regex=None, delay=5, max_depth=-1, max_urls=-1, headers=None, user_agent='wswp',
                 proxy=None, num_retries=1):
    """Crawl from the given seed URL following links matched by link_regex
    """
    # the queue of URL's that still need to be crawled

    orginal_papers_queue = queue.deque([seed_url])
    # the URL's that have been seen and at what depth
    seen = {seed_url: 0}
    # track how many URL's have been downloaded
    num_urls = 0
    rp = get_robots(seed_url)
    throttle = Throttle(delay)

    headers = {
        "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding":"gzip, deflate",
        "Accept-Language":"zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
        "Cache-Control":"max-age=0",
        "Connection":"keep-alive",
        "Upgrade-Insecure-Requests":"1",
        "User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:63.0) Gecko/20100101 Firefox/63.0"
        }
   

    while orginal_papers_queue:
        url = orginal_papers_queue.pop()
        # check url passes robots.txt restrictions
        if rp.can_fetch(user_agent, url):
            throttle.wait(url)
            html = download(url, headers, proxy=proxy, num_retries=num_retries)
            links = []

            depth = seen[url]
            if depth != max_depth:
                # can still crawl further
                if link_regex:
                    # filter for links matching our regular expression
                    links.extend(link for link in get_full_links(html,headers=headers,num_retries=num_retries,seed_url=seed_url))

                for link in links:
                    link = normalize(seed_url, link)
                    # check whether already crawled this link
                    if link not in seen:
                        seen[link] = depth + 1
                        # check link is within same domain
                        if same_domain(seed_url, link):
                            # success! add this new link to queue
                            orginal_papers_queue.append(link)

            # check whether have reached downloaded maximum
            num_urls += 1
            if num_urls == max_urls:
                break
        else:
            print('Blocked by robots.txt: %s'% url)


class Throttle:
    """Throttle downloading by sleeping between requests to same domain
    """

    def __init__(self, delay):
        # amount of delay between downloads for each domain
        self.delay = delay
        # timestamp of when a domain was last accessed
        self.domains = {}

    def wait(self, url):
        domain = urllib.parse.urlparse(url).netloc
        last_accessed = self.domains.get(domain)

        if self.delay > 0 and last_accessed is not None:
            sleep_secs = self.delay - (datetime.now() - last_accessed).seconds
            if sleep_secs > 0:
                time.sleep(sleep_secs)
        self.domains[domain] = datetime.now()


def download(url, headers, proxy, num_retries, data=None):

    global url_crawled_num

    print('Downloading: %s' % url)
    request = urllib.request.Request(url, data, headers)
    opener = urllib.request.build_opener()
    if proxy:
        proxy_params = {urllib.parse.urlparse(url).scheme: proxy}
        opener.add_handler(urllib.ProxyHandler(proxy_params))
    try:
        response = opener.open(request)
        html = response.read()
        code = response.code
        if url.find("full_record.do?product=WOS") != -1 :
            url_crawled_num += 1
            html_emt = etree.HTML(html)
            title = html_emt.xpath("//div[@class='title']/value/text()")
            print("Title %d:  %s" % (url_crawled_num, str(title[0])))
        
    except urllib.error.URLError as e:
        print('Download error: %s' % e.reason)
        html = ''
        if hasattr(e, 'code'):
            code = e.code
            if num_retries > 0 and 500 <= code < 600:
                # retry 5XX HTTP errors
                return download(url, headers, proxy, num_retries - 1, data)
        else:
            code = None
    except urllib.error.HTTPError as e:
        
        print("Download error: %s" % e.reason)
        
    return html


def normalize(seed_url, link):
    """Normalize this URL by removing hash and adding domain
    """
    link, _ = urllib.parse.urldefrag(link)  # remove hash to avoid duplicates
    return urllib.parse.urljoin(seed_url, link)


def same_domain(url1, url2):
    """Return True if both URL's belong to same domain
    """
    return urllib.parse.urlparse(url1).netloc == urllib.parse.urlparse(url2).netloc
 

def get_robots(url):
    """Initialize robots parser for this domain
    """
    rp = urllib.robotparser.RobotFileParser()
    rp.set_url(urllib.parse.urljoin(url, '/robots.txt'))
    rp.read()
    return rp

def get_full_links(html,headers,num_retries,seed_url):
    """Return a list of links from html
    """
    html_emt = etree.HTML(html)
    url_list = html_emt.xpath("//a[contains(@class,'snowplow-full-record')]/@href")

    next_page = html_emt.xpath("//a[contains(@class,'snowplow-navigation-nextpage-bottom')]/@href")

    try:
        if len(next_page) != 0:
            html = download(url=next_page[0],headers=headers,num_retries=num_retries, proxy=None)
            next_url_list = get_full_links(html,headers,num_retries,seed_url)
            url_list.extend(next_url_list)
    except Exception as e:
        print("next_page: %s" % next_page)
        print(str(e))
    return url_list



if __name__ == '__main__':
    
    link_crawler('http://apps.webofknowledge.com/Search.do?product=WOS&SID=7D8ohfK6fXMgCHFGq8Q&search_mode=GeneralSearch&prID=b07521dc-41dc-4a36-8ab9-cc4f697e15e9', '/(index|view)', delay=10, num_retries=2, max_depth=1,
                 user_agent='GoodCrawler')
