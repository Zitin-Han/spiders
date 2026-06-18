# 微博搜索爬虫

基于 Scrapy + Requests 的微博关键词搜索爬虫，支持日期范围、原创筛选，导出 Excel。

## 功能

- 关键词批量搜索（列表/文件两种方式）
- 日期范围筛选，支持窄日期分天展开避免翻页截断
- 原创/热门/全部/关注人/认证用户/媒体/观点 7 种搜索类型
- Excel 输出（openpyxl），断点续爬支持

## 安装

```bash
pip install scrapy requests openpyxl
```

## 配置

编辑 `weibo_standalone.py` 顶部 `CONFIG` 字典：

| 参数 | 说明 | 示例 |
|------|------|------|
| `KEYWORD_LIST` | 关键词列表 | `["大学生", "字节跳动"]` |
| `WEIBO_TYPE` | 0=全部 1=原创 2=热门 | `1` |
| `START_DATE` | 开始日期 | `"2026-04-01"` |
| `END_DATE` | 结束日期 | `"2026-04-03"` |
| `DOWNLOAD_DELAY` | 请求间隔(秒) | `10` |
| `USE_JOBDIR` | 断点续爬 | `True`/`False` |

> ⚠️ `DEFAULT_REQUEST_HEADERS.cookie` 需要更新为有效 Cookie（浏览器登录 weibo.com 后 F12 复制）。

## 运行

```bash
python weibo_standalone.py
```

结果保存在 `结果文件/<关键词>/<关键词>.xlsx`，包含 19 个字段：

id, bid, user_id, 用户昵称, 微博正文, 头条文章url, 发布位置, 艾特用户, 话题,
转发数, 评论数, 点赞数, 发布时间, 发布工具, 微博图片url, 微博视频url, retweet_id, ip, 认证信息
