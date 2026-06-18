# 链家租房爬虫

基于 Requests + BeautifulSoup 的链家（lianjia.com）北京租房数据爬虫，遇到人机验证时自动启动 Selenium 浏览器供手动通过。

## 功能

- 请求链家租房列表页，解析房源信息
- 自动翻页爬取（默认100页）
- 遇到反爬/人机验证时，打开 Chrome 浏览器供用户手动完成验证，验证后继续
- 提取字段：标题、链接、区域、商圈、小区、面积、朝向、户型、楼层、标签、价格、房屋编码
- 导出为 Excel

## 安装

```bash
pip install requests beautifulsoup4 pandas selenium openpyxl
```

需要 Chrome 浏览器及对应版本 ChromeDriver。

## 使用

```bash
jupyter notebook lianjia.ipynb
```

## 版本说明

| 文件 | 目标区域 | 输出文件 |
|------|----------|----------|
| `lianjia.ipynb` | 东城区 | `北京租房东城.xlsx` |
| `lianjia_miyun.ipynb` | 密云区 | `北京租房密云.xlsx` |

修改 `base_url` 即可切换任意北京区域，例如：
- 海淀：`https://bj.lianjia.com/zufang/haidian`
- 朝阳：`https://bj.lianjia.com/zufang/chaoyang`

## 运行流程

1. 输入链家网站 Cookie（手动登录后从浏览器复制）
2. 脚本逐页爬取，每页间隔 5~8 秒
3. 遇验证 → Selenium 打开浏览器 → 用户手动完成 → 按回车继续
4. 数据保存至 `data/` 目录
