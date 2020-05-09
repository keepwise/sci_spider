

title = html_emt.xpath("//div[@class='title']/value/text()")[0]
        #authors = html_emt.xpath( "//div[@class='l-content']//p[@class='FR_field']/span[contains(text(),'By:') or contains(text(),'作者:')]/following-sibling::a/text()|//div[@class='l-content']//p[@class='FR_field']/span[@id='more_authors_authors_txt_label']/a/text()")
        authors = html_emt.xpath("//div[@class='l-content']//div[@class='block-record-info']//p[@class='FR_field']/span[text()='By:' or text()='作者:']/following-sibling::a/text()|//div[@class='l-content']//div[@class='block-record-info']//p[@class='FR_field']/span[@id='more_authors_authors_txt_label']/a/text()")