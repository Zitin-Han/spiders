# 蓝鲸BI项目

基于 Python + Selenium 的电商运营数据自动化管道，对接 **灵动（Lingdong）BI 平台**，实现数据导出 → 本地处理 → 云端同步的全流程自动化。

## 项目背景

针对 Mercado Libre（墨西哥站）店铺的日常运营需求，将灵动 BI 平台上的商品分析、交易分析、广告流量等数据自动导出，合并写入本地运营报表，并同步至金山文档供团队协作查看。

## 架构概览

```
┌─────────────────────────────────────────────────────────────┐
│                        main.py（调度入口）                     │
├──────────────┬──────────────────┬───────────────────────────┤
│  step1_export │  step2_process   │     step3_sync            │
│  数据导出      │  数据处理与合并    │   同步至金山文档            │
│  (Selenium)   │  (pandas+openpyxl)│   (pyautogui+xlwings)    │
└──────────────┴──────────────────┴───────────────────────────┘
```

### 三阶段流水线

| 阶段 | 模块 | 功能 | 核心技术 |
|------|------|------|----------|
| 1️⃣ 导出 | `step1_export.py` | 登录灵动平台，自动导出商品分析、交易分析（近7天 SKU 数据）、广告流量（7天+30天）报表 | Selenium WebDriver |
| 2️⃣ 处理 | `step2_process.py` | 读取下载的 Excel，清洗数值（含百分比识别），写入本地运营模板对应列 | pandas / openpyxl |
| 3️⃣ 同步 | `step3_sync.py` | 将更新后的 Excel 保留格式复制，粘贴到金山文档在线表格 | xlwings / pyautogui |

## 项目结构

```
蓝鲸BI项目/
├── main.py                 # 调度入口，串联三阶段
├── config.py               # 公共配置与工具函数（从 .env 加载凭据）
├── step1_export.py         # 阶段1：Selenium 自动化导出
├── step2_process.py        # 阶段2：数据清洗与本地写入
├── step3_sync.py           # 阶段3：同步至金山文档
├── code.ipynb              # 开发调试用 Jupyter Notebook
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量模板
├── .env                    # 实际凭据（已 gitignore，不上传）
├── 店铺每日数据更新.xlsx     # 运营报表模板
└── 新运营表_XQ5.xlsx        # 参考报表
```

## 快速开始

### 环境要求

- Python 3.10+
- Google Chrome 浏览器
- ChromeDriver（与 Chrome 版本匹配，放入 PATH）

### 安装

```bash
# 克隆仓库
git clone https://github.com/Zitin-Han/spiders.git
cd spiders/蓝鲸BI项目

# 安装依赖
pip install -r requirements.txt

# 配置环境变量（参考 .env.example）
cp .env.example .env
# 编辑 .env 填入你的灵动平台账号密码
```

### 运行

```bash
python main.py
```

程序将依次执行：
1. 打开 Chrome，登录灵动平台，导出各类报表
2. 关闭浏览器，读取下载文件，写入本地 Excel
3. 打开金山文档，粘贴更新后的数据

## 运维说明

- **定时运行**：建议通过 Windows 任务计划程序（Task Scheduler）设置每日定时触发 `python main.py`
- **ChromeDriver 版本**：确保 ChromeDriver 与本地 Chrome 版本一致
- **下载目录**：默认读取 `%USERPROFILE%\Downloads`，可在 `.env` 中自定义
- **金山文档登录**：step3 依赖金山文档网页端已登录状态

## 安全提示

⚠️ `.env` 文件包含灵动平台登录凭据，已通过 `.gitignore` 排除，**切勿提交到仓库**。  
如有需要，仅分享 `.env.example` 模板文件。

## 数据字段说明

运营报表（`店铺每日数据更新.xlsx` → `数据更新表`）包含以下数据列：

| 列范围 | 来源 | 内容 |
|--------|------|------|
| E-G 列 | 历史分析 | SKU、近7日销量、昨日销量 |
| J-N 列 | 商品分析 | 商品ID、购物体验分、折扣价(MXN)、原价(MXN)、近7天访问量 |
| Z-AJ 列 | 广告(7天) | 展现量、点击数、ACoS、广告花费、ROAS 等 |
| AL-AM 列 | 广告(30天) | 商品ID、TACOS(ACOAS) |

> 货币单位：MXN（墨西哥比索）
