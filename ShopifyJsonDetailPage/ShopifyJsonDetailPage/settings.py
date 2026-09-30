BOT_NAME = "ShopifyJsonDetailPage"

SPIDER_MODULES = ["ShopifyJsonDetailPage.spiders"]
NEWSPIDER_MODULE = "ShopifyJsonDetailPage.spiders"

ADDONS = {}

ROBOTSTXT_OBEY = False

SPIDER_LOADER_CLASS = "ShopifyJsonDetailPage.utils.spider_loader.RecursiveSpiderLoader"

CONCURRENT_REQUESTS_PER_DOMAIN = 1
DOWNLOAD_DELAY = 1

ITEM_PIPELINES = {
    "ShopifyJsonDetailPage.pipelines.ShopifyJsonPipeline": 300,
}

FEED_EXPORT_ENCODING = "utf-8"

LOG_LEVEL = "INFO"
