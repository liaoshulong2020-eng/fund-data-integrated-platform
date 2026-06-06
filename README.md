# 基金数据集成平台

一个基于 Python/Tkinter 的基金数据采集、评分、可视化和持仓分析工具。项目把基金数据爬取、四级分类、策略评分、专题筛选、同花顺风格展示、基金详情、净值走势、基金对比、市场指数、资金流向和我的持仓整合在一个桌面程序中。

## 核心功能

- 一键运行：智能清洗、数据爬取、模块预计算。
- 迭代爬取：保留历史净值，日常只补最新净值并合并去重。
- 完整走势：首次可选择补全完整历史净值，后续增量更新。
- 模块评分：收益表现、风险回撤、风险效率、位置估值、趋势择时、基金经理、交易成本、收益归因，每个模块按算法特点展示不同核心指标。
- 综合策略：长期综合、回撤震荡、趋势突破、低波稳健、超跌反弹，支持软件内结果看板、详情和对比。
- 专题筛选：宽基、行业、海外、债券、商品、REITs/FOF 等主题。
- 内置可视化：结果在软件内部展示，不跳转浏览器，表格、图表、详情联动。
- 基金详情：基金基础信息、今日涨幅、阶段收益、净值走势。
- 交互走势图：支持成立来、今年、5年、3年、1年、6月、3月、1月等区间，支持拖拽十字线查看日期和净值。
- 回撤分析：净值走势图中直接标记最大回撤、回撤起点、低点、修复完成、回撤进度和当前回撤。
- 基金对比：选择多只基金进行走势对比。
- 市场首页：全球主要指数、行业/概念资金流。
- 我的持仓：导入持仓，计算市值、收益、今日收益。
- 低内存计算：收益表现和风险回撤支持流式处理，减少大数据文件导致界面卡顿。

## 目录结构

```text
fund-data-integrated-platform/
├── fund_platform.py        # 主程序，完整功能入口
├── requirements.txt        # Python 依赖
├── pyproject.toml          # 项目元信息
├── LICENSE                 # MIT License
├── README.md               # 项目说明
├── .gitignore              # 忽略本地数据、缓存、日志
├── docs/
│   ├── ARCHITECTURE.md     # 系统结构说明
│   └── DATA_WORKFLOW.md    # 数据爬取与迭代更新说明
├── data/                   # 运行后生成数据，不上传真实数据
├── outputs/                # 运行后生成 Excel/报告，不上传真实结果
└── cache/                  # 运行后生成缓存，不上传真实缓存
```

## 环境要求

- Windows 10/11
- Python 3.9+
- Tkinter，通常随 Python 官方安装包自带

当前开发和测试主要使用：

```text
Python 3.9
```

## 安装依赖

```powershell
python -m pip install -r requirements.txt
```

如果电脑上有多个 Python 版本，建议明确使用 Python 3.9：

```powershell
C:\Users\Administrator\AppData\Local\Programs\Python\Python39\python.exe -m pip install -r requirements.txt
```

## 运行

```powershell
C:\Users\Administrator\AppData\Local\Programs\Python\Python39\python.exe fund_platform.py
```

或者在项目目录中运行：

```powershell
python fund_platform.py
```

## 一键运行建议

首次使用：

1. 运行程序。
2. 点击“一键运行”。
3. 选择是否重新智能清洗。
4. 选择是否重新爬取。
5. 首次想要完整走势图时，选择“补全完整历史净值”。

日常更新：

1. 点击“一键运行”。
2. 选择重新爬取。
3. 历史净值补全选择“否”。
4. 系统会读取上一份完整有效数据，只抓最新净值并合并旧历史。

## 数据说明

本仓库不包含真实基金数据、缓存、Excel 结果或日志。运行后会在本地生成：

- `fund_data/`：基金 JSON 数据。
- `fund_excel/`：评分和筛选结果。
- `fund_cache/`：模块缓存、资金流缓存。
- `fund_visual_reports/`：历史报告。
- `target_funds.json`：清洗后的基金池。

这些文件默认被 `.gitignore` 排除，不建议上传到 GitHub。

## 数据来源

项目使用公开页面或公开接口获取数据，主要来源包括：

- 东方财富基金页面和净值接口
- AkShare 公开数据
- 新浪行情接口
- 同花顺数据中心公开页面

不同来源可能存在延迟、字段口径差异或访问限制。数据仅供学习研究和个人分析使用。

## 免责声明

本项目不构成任何投资建议。基金投资存在风险，任何投资决策需自行判断并承担风险。

## License

MIT License
