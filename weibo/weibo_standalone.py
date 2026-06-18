# -*- coding: utf-8 -*-

import os
import re
import sys
import warnings
from datetime import datetime, timedelta
from urllib.parse import unquote

warnings.filterwarnings("ignore", category=UserWarning)

import openpyxl
import requests
import scrapy
from scrapy.crawler import CrawlerProcess
from scrapy.exceptions import CloseSpider, DropItem

CONFIG = {
    "KEYWORD_LIST": ["大学生"],
    "WEIBO_TYPE": 1,
    "START_DATE": "2026-04-01",
    "END_DATE": "2026-04-03",
    "FURTHER_THRESHOLD": 46,
    "DOWNLOAD_DELAY": 10,
    "LOG_LEVEL": "INFO",
    # True=断点续爬  需要全新爬取时请设 False，或删除 crawls/search 目录后再开 True。
    "USE_JOBDIR": False,
    "DEFAULT_REQUEST_HEADERS": {
        "Accept":
        "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-US;q=0.7",
        "cookie": "SCF=AiBz0bA-2MlX09p_dHEcGD6ct6gRtmP1aO27LfRxWs0_rMyJAIM1LE7hPtAybBBef235X0wX3BlXgdMwO6acfaQ.; SUB=_2A25E0DwBDeRhGeFH4lEY8SrEwj2IHXVnrDHJrDV8PUNbmtAbLVHNkW9NehU3MyzaeajfMz93iBTwOvXoASooXLYn; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9Wha_sIOzlrjnuhsoBhickxo5NHD95QN1K.01K2X1h.pWs4Dqcjci--fiK.fiKyWi--fiK.fiKyWi--fi-i8iKnRi--fi-2XiKLWi--fiK.pi-zfi--fiK.pi-zf; ALF=02_1778112849; XSRF-TOKEN=7KAdI18Yq9wlikDu2hqsRD3A; WBPSESS=7uojwwkb7JXoN6Scbv5ln5ACYhf31_mt9f8v-8xUOVP6QN2NHhhkt8LXz4LIhEArev8WVavwCVsibQfEUyo0IQpwI_cjQ14sNXBTe5zKAfj6WDsD7U1bpAoQMI92boqQPdb1PupHtrp1EdDNWiMk9g=="
    },
}

def convert_weibo_type(weibo_type):
    if weibo_type == 0:
        return "&typeall=1"
    elif weibo_type == 1:
        return "&scope=ori"
    elif weibo_type == 2:
        return "&xsort=hot"
    elif weibo_type == 3:
        return "&atten=1"
    elif weibo_type == 4:
        return "&vip=1"
    elif weibo_type == 5:
        return "&category=4"
    elif weibo_type == 6:
        return "&viewpoint=1"
    return "&scope=ori"

SUBALL_SUFFIX = "&suball=1"


def get_keyword_list(file_name):
    with open(file_name, "rb") as f:
        try:
            lines = f.read().splitlines()
            lines = [line.decode("utf-8-sig") for line in lines]
        except UnicodeDecodeError:
            print("%s文件应为utf-8", file_name)
            sys.exit()
        keyword_list = []
        for line in lines:
            if line:
                keyword_list.append(line)
    return keyword_list


def standardize_date(created_at):
    if "刚刚" in created_at:
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    elif "秒" in created_at:
        second = created_at[: created_at.find("秒")]
        second = timedelta(seconds=int(second))
        created_at = (datetime.now() - second).strftime("%Y-%m-%d %H:%M")
    elif "分钟" in created_at:
        minute = created_at[: created_at.find("分钟")]
        minute = timedelta(minutes=int(minute))
        created_at = (datetime.now() - minute).strftime("%Y-%m-%d %H:%M")
    elif "小时" in created_at:
        hour = created_at[: created_at.find("小时")]
        hour = timedelta(hours=int(hour))
        created_at = (datetime.now() - hour).strftime("%Y-%m-%d %H:%M")
    elif "今天" in created_at:
        today = datetime.now().strftime("%Y-%m-%d")
        created_at = today + " " + created_at[2:]
    elif "年" not in created_at:
        year = datetime.now().strftime("%Y")
        month = created_at[:2]
        day = created_at[3:5]
        time = created_at[6:]
        created_at = year + "-" + month + "-" + day + " " + time
    else:
        year = created_at[:4]
        month = created_at[5:7]
        day = created_at[8:10]
        time = created_at[11:]
        created_at = year + "-" + month + "-" + day + " " + time
    return created_at


def str_to_time(text):
    return datetime.strptime(text, "%Y-%m-%d")


class WeiboItem(scrapy.Item):
    id = scrapy.Field()
    bid = scrapy.Field()
    user_id = scrapy.Field()
    screen_name = scrapy.Field()
    text = scrapy.Field()
    article_url = scrapy.Field()
    location = scrapy.Field()
    at_users = scrapy.Field()
    topics = scrapy.Field()
    reposts_count = scrapy.Field()
    comments_count = scrapy.Field()
    attitudes_count = scrapy.Field()
    created_at = scrapy.Field()
    source = scrapy.Field()
    pics = scrapy.Field()
    video_url = scrapy.Field()
    retweet_id = scrapy.Field()
    ip = scrapy.Field()
    user_authentication = scrapy.Field()


class DuplicatesPipeline(object):
    def __init__(self):
        self.ids_seen = set()

    def process_item(self, item, spider):
        if item["weibo"]["id"] in self.ids_seen:
            raise DropItem("过滤重复微博: %s" % item)
        self.ids_seen.add(item["weibo"]["id"])
        return item


class ExcelPipeline(object):
    """仅写入 .xlsx，需: pip install openpyxl"""

    HEADER = [
        "id",
        "bid",
        "user_id",
        "用户昵称",
        "微博正文",
        "头条文章url",
        "发布位置",
        "艾特用户",
        "话题",
        "转发数",
        "评论数",
        "点赞数",
        "发布时间",
        "发布工具",
        "微博图片url",
        "微博视频url",
        "retweet_id",
        "ip",
        "user_authentication",
    ]

    def process_item(self, item, spider):
        base_dir = "结果文件" + os.sep + item["keyword"]
        if not os.path.isdir(base_dir):
            os.makedirs(base_dir)
        file_path = base_dir + os.sep + item["keyword"] + ".xlsx"
        row = [
            item["weibo"].get("id", ""),
            item["weibo"].get("bid", ""),
            item["weibo"].get("user_id", ""),
            item["weibo"].get("screen_name", ""),
            item["weibo"].get("text", ""),
            item["weibo"].get("article_url", ""),
            item["weibo"].get("location", ""),
            item["weibo"].get("at_users", ""),
            item["weibo"].get("topics", ""),
            item["weibo"].get("reposts_count", ""),
            item["weibo"].get("comments_count", ""),
            item["weibo"].get("attitudes_count", ""),
            item["weibo"].get("created_at", ""),
            item["weibo"].get("source", ""),
            ",".join(item["weibo"].get("pics", []) or []),
            item["weibo"].get("video_url", ""),
            item["weibo"].get("retweet_id", ""),
            item["weibo"].get("ip", ""),
            item["weibo"].get("user_authentication", ""),
        ]
        if not os.path.isfile(file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(self.HEADER)
            wb.save(file_path)
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        ws.append(row)
        wb.save(file_path)
        return item


class SearchSpider(scrapy.Spider):
    name = 'search'
    allowed_domains = ['weibo.com']
    keyword_list = CONFIG['KEYWORD_LIST']
    if not isinstance(keyword_list, list):
        if not os.path.isabs(keyword_list):
            keyword_list = os.getcwd() + os.sep + keyword_list
        if not os.path.isfile(keyword_list):
            sys.exit('不存在%s文件' % keyword_list)
        keyword_list = get_keyword_list(keyword_list)

    for i, keyword in enumerate(keyword_list):
        if len(keyword) > 2 and keyword[0] == '#' and keyword[-1] == '#':
            keyword_list[i] = '%23' + keyword[1:-1] + '%23'
    weibo_type = convert_weibo_type(CONFIG['WEIBO_TYPE'])
    suball_suffix = SUBALL_SUFFIX
    base_url = 'https://s.weibo.com'
    start_date = CONFIG.get('START_DATE', datetime.now().strftime('%Y-%m-%d'))
    end_date = CONFIG.get('END_DATE', datetime.now().strftime('%Y-%m-%d'))
    if str_to_time(start_date) > str_to_time(end_date):
        sys.exit('CONFIG 配置错误：START_DATE 应早于或等于 END_DATE')
    further_threshold = CONFIG.get('FURTHER_THRESHOLD', 46)

    def start_requests(self):
        start_date = datetime.strptime(self.start_date, '%Y-%m-%d')
        end_date = datetime.strptime(self.end_date,
                                     '%Y-%m-%d') + timedelta(days=1)
        start_str = start_date.strftime('%Y-%m-%d') + '-0'
        end_str = end_date.strftime('%Y-%m-%d') + '-0'
        for keyword in self.keyword_list:
            base_url = 'https://s.weibo.com/weibo?q=%s' % keyword
            url = base_url + self.weibo_type
            url += self.suball_suffix
            url += '&timescope=custom:{}:{}'.format(start_str, end_str)
            yield scrapy.Request(url=url,
                                 callback=self.parse,
                                 meta={
                                     'base_url': base_url,
                                     'keyword': keyword
                                 })

    def parse(self, response):
        base_url = response.meta.get('base_url')
        keyword = response.meta.get('keyword')
        is_empty = response.xpath(
            '//div[@class="card card-no-result s-pt20b40"]')
        page_count = len(response.xpath('//ul[@class="s-scroll"]/li'))
        if is_empty:
            print('当前页面搜索结果为空')
        elif page_count < self.further_threshold:
            # 解析当前页面
            for weibo in self.parse_weibo(response):
                yield weibo
            next_url = response.xpath(
                '//a[@class="next"]/@href').extract_first()
            if next_url:
                next_url = self.base_url + next_url
                yield scrapy.Request(url=next_url,
                                     callback=self.parse_page,
                                     meta={'keyword': keyword})
        else:
            start_date = datetime.strptime(self.start_date, '%Y-%m-%d')
            end_date = datetime.strptime(self.end_date, '%Y-%m-%d')
            while start_date <= end_date:
                start_str = start_date.strftime('%Y-%m-%d') + '-0'
                start_date = start_date + timedelta(days=1)
                end_str = start_date.strftime('%Y-%m-%d') + '-0'
                url = base_url + self.weibo_type
                url += self.suball_suffix
                url += '&timescope=custom:{}:{}&page=1'.format(
                    start_str, end_str)
                yield scrapy.Request(url=url,
                                     callback=self.parse_by_day,
                                     meta={
                                         'base_url': base_url,
                                         'keyword': keyword,
                                         'date': start_str[:-2]
                                     })

    def parse_by_day(self, response):
        base_url = response.meta.get('base_url')
        keyword = response.meta.get('keyword')
        is_empty = response.xpath(
            '//div[@class="card card-no-result s-pt20b40"]')
        date = response.meta.get('date')
        page_count = len(response.xpath('//ul[@class="s-scroll"]/li'))
        if is_empty:
            print('页面为空')
        elif page_count < self.further_threshold:
            for weibo in self.parse_weibo(response):
                yield weibo
            next_url = response.xpath(
                '//a[@class="next"]/@href').extract_first()
            if next_url:
                next_url = self.base_url + next_url
                yield scrapy.Request(url=next_url,
                                     callback=self.parse_page,
                                     meta={'keyword': keyword})
        else:
            start_date_str = date + '-0'
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d-%H')
            for i in range(1, 25):
                start_str = start_date.strftime('%Y-%m-%d-X%H').replace(
                    'X0', 'X').replace('X', '')
                start_date = start_date + timedelta(hours=1)
                end_str = start_date.strftime('%Y-%m-%d-X%H').replace(
                    'X0', 'X').replace('X', '')
                url = base_url + self.weibo_type
                url += self.suball_suffix
                url += '&timescope=custom:{}:{}&page=1'.format(
                    start_str, end_str)
                # 获取一小时的搜索结果
                yield scrapy.Request(url=url,
                                     callback=self.parse_by_hour,
                                     meta={
                                         'base_url': base_url,
                                         'keyword': keyword,
                                         'start_time': start_str,
                                         'end_time': end_str
                                     })

    def parse_by_hour(self, response):
        keyword = response.meta.get('keyword')
        is_empty = response.xpath(
            '//div[@class="card card-no-result s-pt20b40"]')
        start_time = response.meta.get('start_time')
        end_time = response.meta.get('end_time')
        page_count = len(response.xpath('//ul[@class="s-scroll"]/li'))
        if is_empty:
            print('页面为空')
        elif page_count < self.further_threshold:
            # 解析当前页面
            for weibo in self.parse_weibo(response):
                yield weibo
            next_url = response.xpath(
                '//a[@class="next"]/@href').extract_first()
            if next_url:
                next_url = self.base_url + next_url
                yield scrapy.Request(url=next_url,
                                     callback=self.parse_page,
                                     meta={'keyword': keyword})
        else:
            print(
                '单小时结果过多；从该小时首页分页抓取。'
            )
            base_kw = 'https://s.weibo.com/weibo?q=%s' % keyword
            url = base_kw + self.weibo_type + self.suball_suffix
            url += '&timescope=custom:{}:{}&page=1'.format(start_time, end_time)
            yield scrapy.Request(
                url=url,
                callback=self.parse_page,
                meta={'keyword': keyword},
            )

    def parse_page(self, response):
        keyword = response.meta.get('keyword')
        is_empty = response.xpath(
            '//div[@class="card card-no-result s-pt20b40"]')
        if is_empty:
            print('页面为空')
        else:
            for weibo in self.parse_weibo(response):
                yield weibo
            next_url = response.xpath(
                '//a[@class="next"]/@href').extract_first()
            if next_url:
                next_url = self.base_url + next_url
                yield scrapy.Request(url=next_url,
                                     callback=self.parse_page,
                                     meta={'keyword': keyword})

    def get_ip(self, bid):
        url = f"https://weibo.com/ajax/statuses/show?id={bid}&locale=zh-CN"
        response = requests.get(url, headers=self.settings.get('DEFAULT_REQUEST_HEADERS'))
        if response.status_code != 200:
            return ""
        try:
            data = response.json()
        except requests.exceptions.JSONDecodeError:
            return ""
        ip_str = data.get("region_name", "")
        if ip_str:
            ip_str = ip_str.split()[-1]
        return ip_str

    def get_article_url(self, selector):
        article_url = ''
        text = selector.xpath('string(.)').extract_first().replace(
            '\u200b', '').replace('\ue627', '').replace('\n',
                                                        '').replace(' ', '')
        if text.startswith('发布了头条文章'):
            urls = selector.xpath('.//a')
            for url in urls:
                if url.xpath(
                        'i[@class="wbicon"]/text()').extract_first() == 'O':
                    if url.xpath('@href').extract_first() and url.xpath(
                            '@href').extract_first().startswith('http://t.cn'):
                        article_url = url.xpath('@href').extract_first()
                    break
        return article_url

    def get_location(self, selector):
        a_list = selector.xpath('.//a')
        location = ''
        for a in a_list:
            if a.xpath('./i[@class="wbicon"]') and a.xpath(
                    './i[@class="wbicon"]/text()').extract_first() == '2':
                location = a.xpath('string(.)').extract_first()[1:]
                break
        return location

    def get_at_users(self, selector):
        a_list = selector.xpath('.//a')
        at_users = ''
        at_list = []
        for a in a_list:
            if len(unquote(a.xpath('@href').extract_first())) > 14 and len(
                    a.xpath('string(.)').extract_first()) > 1:
                if unquote(a.xpath('@href').extract_first())[14:] == a.xpath(
                        'string(.)').extract_first()[1:]:
                    at_user = a.xpath('string(.)').extract_first()[1:]
                    if at_user not in at_list:
                        at_list.append(at_user)
        if at_list:
            at_users = ','.join(at_list)
        return at_users

    def get_topics(self, selector):
        a_list = selector.xpath('.//a')
        topics = ''
        topic_list = []
        for a in a_list:
            text = a.xpath('string(.)').extract_first()
            if len(text) > 2 and text[0] == '#' and text[-1] == '#':
                if text[1:-1] not in topic_list:
                    topic_list.append(text[1:-1])
        if topic_list:
            topics = ','.join(topic_list)
        return topics

    def parse_weibo(self, response):
        keyword = response.meta.get('keyword')
        for sel in response.xpath("//div[@class='card-wrap']"):
            info = sel.xpath(
                "div[@class='card']/div[@class='card-feed']/div[@class='content']/div[@class='info']"
            )
            if info:
                weibo = WeiboItem()
                weibo['id'] = sel.xpath('@mid').extract_first()
                bid = sel.xpath(
                    './/div[@class="from"]/a[1]/@href').extract_first(
                ).split('/')[-1].split('?')[0]
                weibo['bid'] = bid
                weibo['user_id'] = info[0].xpath(
                    'div[2]/a/@href').extract_first().split('?')[0].split(
                    '/')[-1]
                weibo['screen_name'] = info[0].xpath(
                    'div[2]/a/@nick-name').extract_first()
                txt_sel = sel.xpath('.//p[@class="txt"]')[0]
                retweet_sel = sel.xpath('.//div[@class="card-comment"]')
                retweet_txt_sel = ''
                if retweet_sel and retweet_sel[0].xpath('.//p[@class="txt"]'):
                    retweet_txt_sel = retweet_sel[0].xpath(
                        './/p[@class="txt"]')[0]
                content_full = sel.xpath(
                    './/p[@node-type="feed_list_content_full"]')
                is_long_weibo = False
                is_long_retweet = False
                if content_full:
                    if not retweet_sel:
                        txt_sel = content_full[0]
                        is_long_weibo = True
                    elif len(content_full) == 2:
                        txt_sel = content_full[0]
                        retweet_txt_sel = content_full[1]
                        is_long_weibo = True
                        is_long_retweet = True
                    elif retweet_sel[0].xpath(
                            './/p[@node-type="feed_list_content_full"]'):
                        retweet_txt_sel = retweet_sel[0].xpath(
                            './/p[@node-type="feed_list_content_full"]')[0]
                        is_long_retweet = True
                    else:
                        txt_sel = content_full[0]
                        is_long_weibo = True
                weibo['text'] = txt_sel.xpath(
                    'string(.)').extract_first().replace('\u200b', '').replace(
                    '\ue627', '')
                weibo['article_url'] = self.get_article_url(txt_sel)
                weibo['location'] = self.get_location(txt_sel)
                if weibo['location']:
                    weibo['text'] = weibo['text'].replace(
                        '2' + weibo['location'], '')
                weibo['text'] = weibo['text'][2:].replace(' ', '')
                if is_long_weibo:
                    weibo['text'] = weibo['text'][:-4]
                weibo['at_users'] = self.get_at_users(txt_sel)
                weibo['topics'] = self.get_topics(txt_sel)
                reposts_count = sel.xpath(
                    './/a[@action-type="feed_list_forward"]/text()').extract()
                reposts_count = "".join(reposts_count)
                try:
                    reposts_count = re.findall(r'\d+.*', reposts_count)
                except TypeError:
                    print(
                        "无法解析转发按钮，可能是 1) 网页布局有改动 2) cookie无效或已过期。\n"
                        )
                    raise CloseSpider()
                weibo['reposts_count'] = reposts_count[
                    0] if reposts_count else '0'
                comments_count = sel.xpath(
                    './/a[@action-type="feed_list_comment"]/text()'
                ).extract_first()
                comments_count = re.findall(r'\d+.*', comments_count)
                weibo['comments_count'] = comments_count[
                    0] if comments_count else '0'
                attitudes_count = sel.xpath(
                    './/a[@action-type="feed_list_like"]/button/span[2]/text()').extract_first()
                attitudes_count = re.findall(r'\d+.*', attitudes_count)
                weibo['attitudes_count'] = attitudes_count[
                    0] if attitudes_count else '0'
                created_at = sel.xpath(
                    './/div[@class="from"]/a[1]/text()').extract_first(
                ).replace(' ', '').replace('\n', '').split('前')[0]
                weibo['created_at'] = standardize_date(created_at)
                source = sel.xpath('.//div[@class="from"]/a[2]/text()'
                                   ).extract_first()
                weibo['source'] = source if source else ''
                pics = ''
                is_exist_pic = sel.xpath(
                    './/div[@class="media media-piclist"]')
                if is_exist_pic:
                    pics = is_exist_pic[0].xpath('ul[1]/li/img/@src').extract()
                    pics = [pic[8:] for pic in pics]
                    pics = [
                        re.sub(r'/.*?/', '/large/', pic, 1) for pic in pics
                    ]
                    pics = ['https://' + pic for pic in pics]
                video_url = ''
                is_exist_video = sel.xpath(
                    './/div[@class="thumbnail"]//video-player').extract_first()
                if is_exist_video:
                    video_url = re.findall(r'src:\'(.*?)\'', is_exist_video)[0]
                    video_url = video_url.replace('&amp;', '&')
                    video_url = 'http:' + video_url
                if not retweet_sel:
                    weibo['pics'] = pics
                    weibo['video_url'] = video_url
                else:
                    weibo['pics'] = ''
                    weibo['video_url'] = ''
                weibo['retweet_id'] = ''
                if retweet_sel and retweet_sel[0].xpath(
                        './/div[@node-type="feed_list_forwardContent"]/a[1]'):
                    retweet = WeiboItem()
                    retweet['id'] = retweet_sel[0].xpath(
                        './/a[@action-type="feed_list_like"]/@action-data'
                    ).extract_first()[4:]
                    retweet['bid'] = retweet_sel[0].xpath(
                        './/p[@class="from"]/a/@href').extract_first().split(
                        '/')[-1].split('?')[0]
                    info = retweet_sel[0].xpath(
                        './/div[@node-type="feed_list_forwardContent"]/a[1]'
                    )[0]
                    retweet['user_id'] = info.xpath(
                        '@href').extract_first().split('/')[-1]
                    retweet['screen_name'] = info.xpath(
                        '@nick-name').extract_first()
                    retweet['text'] = retweet_txt_sel.xpath(
                        'string(.)').extract_first().replace('\u200b',
                                                             '').replace(
                        '\ue627', '')
                    retweet['article_url'] = self.get_article_url(
                        retweet_txt_sel)
                    retweet['location'] = self.get_location(retweet_txt_sel)
                    if retweet['location']:
                        retweet['text'] = retweet['text'].replace(
                            '2' + retweet['location'], '')
                    retweet['text'] = retweet['text'][2:].replace(' ', '')
                    if is_long_retweet:
                        retweet['text'] = retweet['text'][:-4]
                    retweet['at_users'] = self.get_at_users(retweet_txt_sel)
                    retweet['topics'] = self.get_topics(retweet_txt_sel)
                    reposts_count = retweet_sel[0].xpath(
                        './/ul[@class="act s-fr"]/li[1]/a[1]/text()'
                    ).extract_first()
                    reposts_count = re.findall(r'\d+.*', reposts_count)
                    retweet['reposts_count'] = reposts_count[
                        0] if reposts_count else '0'
                    comments_count = retweet_sel[0].xpath(
                        './/ul[@class="act s-fr"]/li[2]/a[1]/text()'
                    ).extract_first()
                    comments_count = re.findall(r'\d+.*', comments_count)
                    retweet['comments_count'] = comments_count[
                        0] if comments_count else '0'
                    attitudes_count = retweet_sel[0].xpath(
                        './/a[@class="woo-box-flex woo-box-alignCenter woo-box-justifyCenter"]//span[@class="woo-like-count"]/text()'
                    ).extract_first()
                    attitudes_count = re.findall(r'\d+.*', attitudes_count)
                    retweet['attitudes_count'] = attitudes_count[
                        0] if attitudes_count else '0'
                    created_at = retweet_sel[0].xpath(
                        './/p[@class="from"]/a[1]/text()').extract_first(
                    ).replace(' ', '').replace('\n', '').split('前')[0]
                    retweet['created_at'] = standardize_date(created_at)
                    source = retweet_sel[0].xpath(
                        './/p[@class="from"]/a[2]/text()').extract_first()
                    retweet['source'] = source if source else ''
                    retweet['pics'] = pics
                    retweet['video_url'] = video_url
                    retweet['retweet_id'] = ''
                    yield {'weibo': retweet, 'keyword': keyword}
                    weibo['retweet_id'] = retweet['id']
                weibo["ip"] = self.get_ip(bid)

                avator = sel.xpath(
                    "div[@class='card']/div[@class='card-feed']/div[@class='avator']"
                )
                if avator:
                    user_auth = avator.xpath('.//svg/@id').extract_first()
                    print(user_auth)
                    if user_auth == 'woo_svg_vblue':
                        weibo['user_authentication'] = '蓝V'
                    elif user_auth == 'woo_svg_vyellow':
                        weibo['user_authentication'] = '黄V'
                    elif user_auth == 'woo_svg_vorange':
                        weibo['user_authentication'] = '红V'
                    elif user_auth == 'woo_svg_vgold':
                        weibo['user_authentication'] = '金V'
                    else:
                        weibo['user_authentication'] = '普通用户'
                print(weibo)
                yield {'weibo': weibo, 'keyword': keyword}

def _scrapy_settings_dict():
    m = sys.modules[__name__]
    pname = m.__name__ + ".DuplicatesPipeline"
    cname = m.__name__ + ".ExcelPipeline"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    settings = {
        "BOT_NAME": "weibo_standalone",
        "COOKIES_ENABLED": False,
        "TELNETCONSOLE_ENABLED": False,
        "LOG_LEVEL": CONFIG["LOG_LEVEL"],
        "DOWNLOAD_DELAY": CONFIG["DOWNLOAD_DELAY"],
        "DEFAULT_REQUEST_HEADERS": CONFIG["DEFAULT_REQUEST_HEADERS"],
        "ITEM_PIPELINES": {
            pname: 300,
            cname: 301,
        },
    }
    if CONFIG.get("USE_JOBDIR"):
        settings["JOBDIR"] = os.path.join(script_dir, "crawls", "search")
    return settings


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(
        "weibo_standalone 启动，工作目录:",
        script_dir,
        "| USE_JOBDIR=",
        CONFIG.get("USE_JOBDIR"),
        flush=True,
    )
    if CONFIG.get("USE_JOBDIR"):
        print(
            "提示: 已启用 JOBDIR。若任务与之前完全相同，可能因去重立即结束；"
            "可删除 crawls/search 后重试，或将 CONFIG['USE_JOBDIR'] 设为 False。",
            flush=True,
        )
    process = CrawlerProcess(_scrapy_settings_dict())
    process.crawl(SearchSpider)
    process.start()

if __name__ == "__main__":
    main()
