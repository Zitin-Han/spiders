# Spiders

个人爬虫代码集合，按项目分目录存放。

## 项目列表

| 目录 | 说明 | 状态 |
|------|------|------|
| [tmall](./tmall/) | 天猫商品评论爬虫（Selenium + 批量 Excel 任务） | ✅ 可用 |

## 环境要求

- Python 3.8+
- Chrome 浏览器

## 通用依赖

```bash
pip install selenium webdriver-manager pandas openpyxl
```

## tmall / 天猫商品评论爬虫

### 文件说明

| 文件 | 说明 |
|------|------|
| `天猫.ipynb` | 爬虫与分析 Notebook |
| `血糖仪.xlsx` | 批量任务配置（`url` 列填入商品链接） |

### 快速开始

1. 进入 `tmall` 目录，打开 `天猫.ipynb`
2. 修改 `血糖仪.xlsx` 中 `url` 列，填入天猫商品链接
3. 运行批量爬虫单元格，浏览器扫码登录淘宝
4. 评论保存为 `result_{商品ID}.csv`

### 注意事项

- 需淘宝账号扫码登录
- 商品间自动休息 5–10 秒，降低反爬风险
- 页面结构变更时可能需要更新选择器
