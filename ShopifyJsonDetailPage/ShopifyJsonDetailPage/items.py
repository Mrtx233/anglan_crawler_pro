import scrapy


class ShopifyJsonItem(scrapy.Item):
    title = scrapy.Field()
    name = scrapy.Field()
    price1 = scrapy.Field()
    price2 = scrapy.Field()
    styles1 = scrapy.Field()
    styles2 = scrapy.Field()
    styles3 = scrapy.Field()
    src_links = scrapy.Field()
    details = scrapy.Field()


# 动态注册含连字符的字段
ShopifyJsonItem.fields["link-href"] = scrapy.Field()
