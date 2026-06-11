# -*- coding: utf-8 -*-
# Auto-generated unified system script
import os
import sys
import glob
import threading
import webbrowser
from flask import Flask, jsonify, request
import types
from datetime import datetime as _dt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

EMBEDDED_SOURCES = {
    'jijin_system.py': '# -*- coding: utf-8 -*-\n"""\n基金数据工具集成平台 v3.2 - 四级分类 + 标签系统\n================================================\n\n四级分类：资产大类 → 子类 → 主题 → 策略\n    · 数据中心       全量导出 / 智能清洗 / 开始爬取 / 转Excel / 市场分析\n    · 单模块评分     收益 / 风险 / 效率 / 归因 / 位置 / 趋势 / 经理 / 成本\n    · 综合策略       长期综合 / 回撤震荡 / 趋势突破 / 低波稳健 / 超跌反弹\n    · 权益类         A股/港股/美股 宽基 + 行业 + 红利 + 主动\n    · 债券类         短债 / 纯债 / 可转债 / 一级二级 / 美债\n    · 海外类         美股 / 港股 / 日经 / 欧洲\n    · 商品类         黄金 / 原油\n    · REITs / FOF\n\n基金标签系统：build_fund_tags() 为每只基金派生多维度标签（资产大类/细分/主题/状态）\n工具：⚡ 一键运行 —— 清洗→爬取→长期综合 全自动流水线\n"""\n\nimport tkinter as tk\nfrom tkinter import ttk, messagebox, scrolledtext\nimport threading\nimport json\nimport sys\nimport os\nimport glob\nimport re\nimport time\nimport datetime\nimport requests\nimport schedule\nfrom collections import Counter\nfrom concurrent.futures import ThreadPoolExecutor, as_completed\nfrom datetime import datetime as dt\nfrom typing import Any, Optional\nfrom copy import copy\n\nimport numpy as np\nimport pandas as pd\nfrom openpyxl import load_workbook\nfrom openpyxl.utils import get_column_letter\nfrom openpyxl.styles import Alignment, Font, PatternFill\nfrom openpyxl.comments import Comment\n\n\n# ========================================================\n# 全局爬取控制器\n# ========================================================\nclass CrawlController:\n    """线程安全的爬取控制器：支持暂停 / 继续 / 停止"""\n    def __init__(self):\n        self._pause_event = threading.Event()\n        self._stop_flag = False\n        self.results = []\n        self._lock = threading.Lock()\n        self._pause_event.set()\n\n    def pause(self):\n        self._pause_event.clear()\n\n    def resume(self):\n        self._pause_event.set()\n\n    def stop(self):\n        self._stop_flag = True\n        self._pause_event.set()\n\n    def reset(self):\n        self._pause_event.set()\n        self._stop_flag = False\n        self.results = []\n\n    def wait_if_paused(self):\n        self._pause_event.wait()\n        return not self._stop_flag\n\n    def is_stopped(self):\n        return self._stop_flag\n\n    def add_result(self, data):\n        with self._lock:\n            self.results.append(data)\n\n    def get_results(self):\n        with self._lock:\n            return list(self.results)\n\n\ncontroller = CrawlController()\n\n\n# ========================================================\n# ① 全量导出\n# ========================================================\ndef run_export_all(log):\n    try:\n        import akshare as ak\n        log("=" * 60)\n        log("正在获取市面上所有基金...")\n        log("=" * 60)\n\n        all_funds = ak.fund_name_em()\n        log(f"\\n成功获取 {len(all_funds):,} 只基金")\n\n        today = dt.now().strftime("%Y%m%d")\n\n        filename_codes = f"所有基金代码_{today}.txt"\n        with open(filename_codes, \'w\', encoding=\'utf-8\') as f:\n            for code in all_funds[\'基金代码\']:\n                f.write(f"{code}\\n")\n        log(f"\\n已保存纯代码列表: {filename_codes}")\n\n        filename_detailed = f"所有基金列表_{today}.txt"\n        with open(filename_detailed, \'w\', encoding=\'utf-8\') as f:\n            f.write("# 市面上所有基金列表\\n")\n            f.write(f"# 总数: {len(all_funds):,} 只\\n")\n            f.write(f"# 生成日期: {today}\\n\\n")\n            for _, row in all_funds.iterrows():\n                f.write(f"{row[\'基金代码\']:<10}  # {row[\'基金简称\']}\\n")\n        log(f"已保存详细列表: {filename_detailed}")\n\n        filename_csv = f"所有基金列表_{today}.csv"\n        all_funds[[\'基金代码\', \'基金简称\']].to_csv(filename_csv, index=False, encoding=\'utf-8-sig\')\n        log(f"已保存 CSV 格式: {filename_csv}")\n\n        log("\\n全部完成!")\n        log("=" * 60)\n\n    except Exception as e:\n        log(f"发生错误: {e}")\n\n\n# ========================================================\n# ② 智能清洗\n# ========================================================\ndef run_clean_list(log):\n    try:\n        import akshare as ak\n        log("正在获取全市场基金列表...")\n        df = ak.fund_name_em()\n        log(f"原始总数: {len(df):,} 只")\n\n        # 基础池保留所有可被策略使用的类别（含债券/FOF/REITs/商品）\n        TARGET_TYPES = [\'股票型\', \'混合型\', \'指数型\', \'QDII\', \'ETF\', \'联接\',\n                        \'债券型\', \'债券\', \'FOF\', \'REITs\', \'商品\']\n        # 仅剔除真正不参与评分的类别\n        DROP_TYPES = [\'货币\', \'理财\']\n\n        if \'基金类型\' in df.columns:\n            pattern = \'|\'.join(TARGET_TYPES)\n            mask_keep = df[\'基金类型\'].str.contains(pattern, na=False)\n            mask_drop = df[\'基金类型\'].str.contains(\'|\'.join(DROP_TYPES), na=False)\n            df_clean = df[mask_keep & ~mask_drop]\n        else:\n            df_clean = df\n\n        mask_backend = df_clean[\'基金简称\'].str.contains(\'后端\', na=False)\n        df_clean = df_clean[~mask_backend]\n\n        log(f"清洗后剩余: {len(df_clean):,} 只（含债券/FOF/REITs/商品，策略评分时再按需剔除）")\n\n        target_codes = df_clean[\'基金代码\'].tolist()\n        with open("target_funds.json", \'w\', encoding=\'utf-8\') as f:\n            json.dump(target_codes, f)\n\n        log(f"\\n已生成清洗后的列表: target_funds.json")\n        log("可直接点击【数据爬取】使用此列表")\n\n    except Exception as e:\n        log(f"发生错误: {e}")\n\n\n\n# ========================================================\n# ③ 数据爬取核心\n# ========================================================\ndef now_str():\n    return dt.now().strftime("%Y-%m-%d %H:%M:%S")\n\n\ndef norm_html_space(s):\n    if not s: return s\n    return s.replace("\\xa0", " ").replace("&nbsp;", " ")\n\n\ndef parse_jsonp(text):\n    text = text.strip()\n    l = text.find("(")\n    r = text.rfind(")")\n    if l != -1 and r != -1 and r > l:\n        return text[l + 1: r].strip()\n    return text\n\n\nclass FundDecompiler:\n    def __init__(self, html_source: str, code: str):\n        self.html = norm_html_space(html_source)\n        self.code = code\n        self.data = {\n            "fund_code": code,\n            "extract_time": now_str(),\n            "fund_name": "--",\n            "base_info": {},\n            "performance": {},\n            "status": {}\n        }\n\n    def _regex(self, pattern, default="--"):\n        match = re.search(pattern, self.html, re.S)\n        return match.group(1).strip() if match else default\n\n    def parse_base(self):\n        self.data["fund_name"] = self._regex(r"<div style=\\"float: left\\">(.*?)<span>")\n        self.data["base_info"]["fund_type"] = self._regex(r"类型：<a href=[^>]*>(.*?)</a>")\n        self.data["base_info"]["risk_level"] = self._regex(r"类型：.*?\\|&nbsp;&nbsp;(.*?)</td>")\n        self.data["base_info"]["assets_size"] = self._regex(r"规模</a>：(.*?)（")\n        self.data["base_info"]["assets_date"] = self._regex(r"规模</a>：.*?（(.*?)）")\n        self.data["base_info"]["manager"] = self._regex(r"基金经理：<a href=[^>]*>(.*?)</a>")\n        self.data["base_info"]["company"] = self._regex(r"管 理 人</span>：<a href=[^>]*>(.*?)</a>")\n        self.data["base_info"]["setup_date"] = self._regex(r"成 立 日</span>：(\\d{4}-\\d{2}-\\d{2})")\n\n    def parse_stage_returns(self):\n        perf_patterns = {\n            "1m": r"<span>近1月：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>",\n            "3m": r"<span>近3月：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>",\n            "6m": r"<span>近6月：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>",\n            "1y": r"<span>近1年：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>",\n            "3y": r"<span>近3年：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>",\n            "since": r"<span>成立来：</span><span class=\\"[^\\"]*\\">([-+]?\\d+\\.\\d+%)</span>"\n        }\n        for key, p in perf_patterns.items():\n            val = re.search(p, self.html, re.S)\n            self.data["performance"][key] = val.group(1).strip() if val else "--"\n\n    def parse_trade_status(self):\n        self.data["status"]["buy_status"] = self._regex(r"交易状态：.*?(开放申购|限大额|暂停申购|限量|关闭)")\n        self.data["status"]["sell_status"] = self._regex(r"交易状态：.*?(开放赎回|暂停赎回)")\n        self.data["status"]["buy_fee"] = self._regex(r"购买手续费：.*?<span class=\\"nowPrice\\">(\\d+\\.\\d+%)</span>")\n\n        trade_text = self._regex(r"交易状态：(.*?)(开放赎回|暂停赎回|$)", default="")\n        if trade_text and trade_text != "--":\n            trade_text = re.sub(r\'\\s+\', \' \', trade_text.strip())\n            self.data["status"]["buy_limit_full"] = trade_text\n        else:\n            self.data["status"]["buy_limit_full"] = "--"\n\n        limit_match = re.search(r\'单日累计购买上限(\\d+\\.?\\d*)元\', trade_text if trade_text else "")\n        if limit_match:\n            self.data["status"]["buy_status"] = "限" + limit_match.group(1) + "元"\n\n        if "该基金暂不开放购买" in self.html:\n            self.data["status"]["buy_status"] = "暂停申购"\n\n    def try_parse_nav_from_html(self):\n        nav = self._regex(r"单位净值</a></span>.*?ui-num\\">([\\d.]+)</span>")\n        nav_date = self._regex(r"单位净值</a></span>\\s*\\((\\d{4}-\\d{2}-\\d{2})\\)")\n        daily = self._regex(r"单位净值</a></span>.*?ui-num\\">([-+]?\\d+\\.\\d+%)</span>")\n        return nav_date, nav, daily\n\n    def parse_all(self, headers):\n        self.parse_base()\n        nav_date, nav, daily = self.try_parse_nav_from_html()\n        if nav == "--" or nav_date == "--":\n            url = f"https://api.fund.eastmoney.com/f10/lsjz?fundCode={self.code}&pageIndex=1&pageSize=1"\n            try:\n                h = dict(headers)\n                h["Referer"] = f"https://fund.eastmoney.com/{self.code}.html"\n                resp = requests.get(url, headers=h, timeout=12)\n                resp.encoding = "utf-8"\n                j = json.loads(parse_jsonp(resp.text))\n                lst = (j.get("Data") or {}).get("LSJZList") or []\n                if lst:\n                    row = lst[0]\n                    nav_date = row.get("FSRQ", "--")\n                    nav = row.get("DWJZ", "--")\n                    daily = row.get("JZZZL", "--")\n                    if daily != "--" and not str(daily).endswith("%"):\n                        daily += "%"\n            except:\n                pass\n\n        self.data["performance"]["nav"] = nav\n        self.data["performance"]["nav_date"] = nav_date\n        self.data["performance"]["daily_growth_rate"] = daily\n        self.parse_stage_returns()\n        self.parse_trade_status()\n        return self.data\n\n\ndef fetch_nav_history(code, headers, session=None, max_pages=10, page_size=20):\n    """抓取单只基金的历史净值（分页），用于后续风险/回撤分析。\n    返回 list[{"date": "YYYY-MM-DD", "val": float}]，按日期升序。\n    支持传入 requests.Session 复用 TCP 连接以加速。"""\n    sess = session or requests\n    all_data = []\n    h = dict(headers)\n    h["Referer"] = f"https://fund.eastmoney.com/{code}.html"\n    for page in range(1, max_pages + 1):\n        url = f"https://api.fund.eastmoney.com/f10/lsjz?fundCode={code}&pageIndex={page}&pageSize={page_size}"\n        try:\n            resp = sess.get(url, headers=h, timeout=12)\n            resp.encoding = "utf-8"\n            j = json.loads(parse_jsonp(resp.text))\n            rows = (j.get("Data") or {}).get("LSJZList") or []\n            if not rows:\n                break\n            for r in rows:\n                v = r.get("LJJZ") or r.get("DWJZ")\n                if v in (None, "", "--"):\n                    continue\n                try:\n                    all_data.append({"date": r.get("FSRQ"), "val": float(v)})\n                except Exception:\n                    continue\n            if len(rows) < page_size:\n                break\n        except Exception:\n            break\n    all_data.sort(key=lambda x: x["date"])\n    return all_data\n\n\n# 每个线程维护一个 Session，TCP / TLS 复用，减少握手开销\n_thread_local = threading.local()\n\n\ndef _get_session(headers):\n    sess = getattr(_thread_local, "session", None)\n    if sess is None:\n        sess = requests.Session()\n        sess.headers.update(headers)\n        adapter = requests.adapters.HTTPAdapter(pool_connections=16, pool_maxsize=16, max_retries=1)\n        sess.mount("https://", adapter)\n        sess.mount("http://", adapter)\n        _thread_local.session = sess\n    return sess\n\n\ndef fetch_one(args):\n    """单个基金任务：抓 profile + nav_history。使用线程本地 Session。"""\n    idx, total, code, headers = args\n    code = str(code).zfill(6)\n    try:\n        session = _get_session(headers)\n        url = f"https://fund.eastmoney.com/{code}.html"\n        resp = session.get(url, timeout=12)\n        resp.encoding = "utf-8"\n        if resp.status_code != 200:\n            return False, None, code\n        fund_data = FundDecompiler(resp.text, code).parse_all(headers)\n        fund_data["nav_history"] = fetch_nav_history(code, headers, session=session)\n        return True, fund_data, code\n    except Exception:\n        return False, None, code\n\n\ndef results_to_dataframe(results):\n    """把爬取结果列表转为格式化 DataFrame"""\n    df = pd.json_normalize(results)\n\n    COLUMN_MAPPING = {\n        "fund_code": "基金代码", "fund_name": "基金名称",\n        "performance.nav_date": "净值日期", "performance.nav": "单位净值",\n        "performance.daily_growth_rate": "日增长率",\n        "performance.1m": "近1月", "performance.3m": "近3月",\n        "performance.6m": "近6月", "performance.1y": "近1年",\n        "performance.3y": "近3年", "performance.since": "成立来",\n        "base_info.fund_type": "基金类型", "base_info.risk_level": "风险等级",\n        "base_info.assets_size": "基金规模", "base_info.manager": "基金经理",\n        "base_info.company": "基金公司", "base_info.setup_date": "成立日期",\n        "status.buy_status": "申购状态", "status.sell_status": "赎回状态",\n        "status.buy_fee": "购买手续费"\n    }\n\n    COLUMN_ORDER = [\n        "基金代码", "基金名称", "净值日期", "单位净值", "日增长率",\n        "近1月", "近3月", "近6月", "近1年", "近3年", "成立来",\n        "基金规模", "基金经理", "基金类型", "风险等级", "基金公司",\n        "成立日期", "申购状态", "赎回状态", "购买手续费"\n    ]\n\n    df.rename(columns=COLUMN_MAPPING, inplace=True)\n    final_cols = [c for c in COLUMN_ORDER if c in df.columns]\n    return df[final_cols]\n\n\ndef save_results_to_excel(results, output_path, log=None):\n    """将结果保存为 Excel，返回是否成功"""\n    if not results:\n        if log: log("暂无可导出数据，请先爬取至少一条。")\n        return False\n    try:\n        df = results_to_dataframe(results)\n        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)\n        with pd.ExcelWriter(output_path, engine=\'openpyxl\') as writer:\n            df.to_excel(writer, index=False, sheet_name=\'基金数据\')\n            ws = writer.sheets[\'基金数据\']\n            ws.freeze_panes = \'A2\'\n        if log: log(f"已导出 {len(df)} 条数据 -> {output_path}")\n        return True\n    except Exception as e:\n        if log: log(f"导出失败: {e}")\n        return False\n\n\ndef run_crawler(log, on_progress=None, on_done=None):\n    """爬取主逻辑，支持暂停/继续/停止。"""\n    try:\n        log("开始数据爬取...")\n        FUND_CODES_FILE = "target_funds.json"\n        OUTPUT_DIR = "fund_data"\n\n        if not os.path.exists(FUND_CODES_FILE):\n            log("未找到 target_funds.json，请先运行【智能清洗】")\n            if on_done: on_done()\n            return\n\n        with open(FUND_CODES_FILE, "r", encoding="utf-8") as f:\n            tasks = json.load(f)\n        if isinstance(tasks, dict) and "funds" in tasks:\n            tasks = tasks["funds"]\n        tasks = [str(x).zfill(6) for x in tasks]\n        total = len(tasks)\n\n        log(f"共 {total} 只基金，开始并发爬取...")\n        log("提示：可随时点击【暂停】，暂停后可点击【导出当前】检查数据格式")\n\n        headers = {\n            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",\n            "Accept-Language": "zh-CN,zh;q=0.9",\n            "Connection": "keep-alive",\n            "Referer": "https://fund.eastmoney.com/",\n        }\n\n        MAX_WORKERS = 8   # 并发线程数，过大容易触发风控\n        t0 = time.time()\n        ok = 0\n        done = 0\n\n        stopped = False\n        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:\n            # 提交所有任务\n            future_map = {}\n            for idx, code in enumerate(tasks, 1):\n                fut = pool.submit(fetch_one, (idx, total, code, headers))\n                future_map[fut] = (idx, code)\n\n            for future in as_completed(future_map):\n                # 暂停/停止处理\n                if not controller.wait_if_paused():\n                    stopped = True\n                    # 尝试取消剩余未开始任务\n                    for f in future_map:\n                        if not f.done():\n                            f.cancel()\n                    break\n\n                idx, code = future_map[future]\n                done += 1\n                try:\n                    success, data, ret_code = future.result()\n                except Exception as e:\n                    success, data, ret_code = False, None, code\n                    log(f"[{done}/{total}] {code} 异常: {e}")\n\n                if success and data:\n                    controller.add_result(data)\n                    ok += 1\n                    log(f"[{done}/{total}] {ret_code} {data[\'fund_name\']} | "\n                        f"净值:{data[\'performance\'].get(\'nav\', \'--\')}")\n                else:\n                    log(f"[{done}/{total}] {ret_code} 失败")\n\n                if on_progress:\n                    on_progress(ok, total)\n\n        elapsed = time.time() - t0\n        if stopped:\n            log(f"\\n已停止爬取，已完成 {done}/{total}，成功 {ok} 条，用时 {elapsed:.1f}s")\n        else:\n            log(f"\\n爬取完成！成功 {ok}/{total} 条，用时 {elapsed:.1f}s，平均 {elapsed / max(total, 1):.2f}s/只")\n\n        results = controller.get_results()\n        if results:\n            # 按原始顺序排序\n            code_order = {code: i for i, code in enumerate(tasks)}\n            results.sort(key=lambda d: code_order.get(str(d.get("fund_code", "")).zfill(6), 9999))\n            os.makedirs(OUTPUT_DIR, exist_ok=True)\n            ts = dt.now().strftime("%Y%m%d_%H%M%S")\n            outfile = os.path.join(OUTPUT_DIR, f"fund_profile_{ts}.json")\n            with open(outfile, "w", encoding="utf-8") as f:\n                json.dump(results, f, ensure_ascii=False, indent=4)\n            log(f"JSON 已保存: {outfile}")\n\n    except Exception as e:\n        log(f"爬取过程出错: {e}")\n    finally:\n        if on_done: on_done()\n\n\n\n# ========================================================\n# ④ JSON 转 Excel（从文件）\n# ========================================================\ndef run_to_excel(log):\n    try:\n        INPUT_DIR = "fund_data"\n        OUTPUT_DIR = "fund_excel"\n\n        files = glob.glob(os.path.join(INPUT_DIR, "*.json"))\n        if not files:\n            log("fund_data 目录下没有找到JSON文件")\n            return None\n        latest_file = max(files, key=os.path.getmtime)\n        log(f"正在处理: {latest_file}")\n\n        with open(latest_file, \'r\', encoding=\'utf-8\') as f:\n            data = json.load(f)\n\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        output_path = os.path.join(OUTPUT_DIR, os.path.basename(latest_file).replace(\'.json\', \'.xlsx\'))\n        ok = save_results_to_excel(data, output_path, log)\n        return os.path.abspath(output_path) if ok else None\n\n    except Exception as e:\n        log(f"转Excel失败: {e}")\n        return None\n\n\n# ========================================================\n# ⑤ 收益表现评分（新集成功能）\n# ========================================================\n\n# 指标含义定义（用于 Excel 悬停批注）\nMETRIC_EXPLAIN = {\n    "收益表现评分": "含义：综合近1月至5年各维度表现的加权得分。\\n判断：分数越高（100为满分）代表历史综合排位越靠前。",\n    "近1月": "判断：数值越大越好（红色代表上涨）。",\n    "近3月": "判断：数值越大越好（红色代表上涨）。",\n    "近6月": "判断：数值越大越好（红色代表上涨）。",\n    "近1年": "判断：数值越大越好（红色代表上涨）。",\n    "近3年": "判断：数值越大越好（红色代表上涨）。",\n    "近5年": "判断：数值越大越好（红色代表上涨）。",\n    "成立以来": "判断：数值越大越好。",\n    "规模": "提示：规模过大可能导致调仓困难，过小可能面临清盘风险。",\n    "成立年限": "提示：年化指标需要参考成立年限，年限越长数据参考价值越高。"\n}\n\nRETURN_WEIGHTS = {\n    "return_1m": 8, "return_3m": 12, "return_6m": 12,\n    "return_1y": 18, "return_3y": 15, "return_5y": 15,\n    "annualized_return": 10,\n}\n\n\ndef parse_pct_to_float(val: Any) -> Optional[float]:\n    if val is None: return None\n    if isinstance(val, (int, float)): return float(val)\n    s = str(val).strip().replace("%", "")\n    if s in ("", "--"): return None\n    try:\n        return float(s)\n    except:\n        return None\n\n\ndef fmt_pct(val):\n    v = parse_pct_to_float(val)\n    return f"{v:.2f}%" if v is not None else "--"\n\n\ndef parse_date(s):\n    try:\n        return datetime.datetime.strptime(str(s)[:10], "%Y-%m-%d").date()\n    except:\n        return None\n\n\ndef calc_age(setup, nav):\n    d0, d1 = parse_date(setup), parse_date(nav)\n    if not d0 or not d1: return None\n    return (d1 - d0).days / 365.25\n\n\ndef winsorize(s):\n    return s.clip(s.quantile(0.01), s.quantile(0.99))\n\n\ndef percentile_rank(s):\n    return s.rank(pct=True)\n\n\ndef allowed(m, age):\n    if age is None: return False\n    if m == "return_1y": return age >= 1\n    if m == "return_3y": return age >= 3\n    if m == "return_5y": return age >= 5\n    return True\n\n\ndef calc_score(df, age_series):\n    pct_df = pd.DataFrame(index=df.index)\n    for col in df.columns:\n        pct_df[col] = percentile_rank(winsorize(df[col]))\n    scores = {}\n    for code in df.index:\n        age = age_series.get(code)\n        total_w, s = 0, 0\n        for m, w in RETURN_WEIGHTS.items():\n            if m not in pct_df: continue\n            if not allowed(m, age): continue\n            v = pct_df.at[code, m]\n            if pd.isna(v): continue\n            total_w += w\n            s += v * w\n        if total_w == 0:\n            scores[code] = None\n        else:\n            score = s / total_w * 100\n            if total_w < 40: score *= 0.9\n            scores[code] = round(score, 2)\n    return pd.Series(scores)\n\n\ndef autosize_excel(path):\n    """美化 Excel：冻结首行、红绿配色、列宽自适应、批注"""\n    wb = load_workbook(path)\n    ws = wb.active\n    ws.freeze_panes = "A2"\n\n    red_font = Font(color="FF0000")\n    green_font = Font(color="00B050")\n    header_font = Font(bold=True)\n\n    headers = [str(cell.value) for cell in ws[1]]\n\n    for row in ws.iter_rows(min_row=2):\n        for cell in row:\n            col_name = headers[cell.column - 1]\n            cell.alignment = Alignment(wrap_text=True, vertical=\'center\', horizontal=\'left\')\n\n            color_target_cols = ["近1月", "近3月", "近6月", "近1年", "近3年", "近5年", "成立以来", "收益表现评分"]\n            if col_name in color_target_cols and cell.value:\n                try:\n                    val_str = str(cell.value).replace(\'%\', \'\')\n                    num_val = float(val_str)\n                    if num_val > 0:\n                        cell.font = red_font\n                    elif num_val < 0:\n                        cell.font = green_font\n                except:\n                    pass\n\n            if col_name == "收益表现评分":\n                cell.font = Font(bold=True, color=cell.font.color if cell.font else None)\n\n    for cell in ws[1]:\n        cell.font = header_font\n        cell.alignment = Alignment(horizontal=\'center\', vertical=\'center\')\n        col_txt = str(cell.value)\n        if col_txt in METRIC_EXPLAIN:\n            comment = Comment(METRIC_EXPLAIN[col_txt], "FundAnalyzer")\n            comment.width = 260\n            comment.height = 80\n            cell.comment = comment\n\n    for col in ws.columns:\n        max_len = 0\n        col_letter = get_column_letter(col[0].column)\n        for cell in col:\n            if cell.value:\n                try:\n                    curr_len = len(str(cell.value).encode(\'gbk\'))\n                except:\n                    curr_len = len(str(cell.value))\n                if curr_len > max_len:\n                    max_len = curr_len\n        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)\n\n    wb.save(path)\n\n\ndef run_performance_score(log):\n    """收益表现评分：自动读取最新 JSON，计算评分，输出美化 Excel"""\n    try:\n        INPUT_DIR = "fund_data"\n        OUTPUT_DIR = "fund_excel"\n\n        files = glob.glob(os.path.join(INPUT_DIR, "*.json"))\n        if not files:\n            log("未找到 fund_data 目录下的 JSON 文件，请先爬取数据。")\n            return None\n\n        latest_file = max(files, key=os.path.getmtime)\n        log(f"正在处理: {latest_file}")\n\n        with open(latest_file, \'r\', encoding=\'utf-8\') as f:\n            data = json.load(f)\n\n        results = data if isinstance(data, list) else data.get("results", [])\n        if not results:\n            log("JSON 文件中无有效数据。")\n            return None\n\n        log(f"共加载 {len(results)} 只基金，开始计算收益表现评分...")\n\n        metric_rows = []\n        full_rows = []\n\n        for item in results:\n            perf = item.get("performance", {})\n            base = item.get("base_info", {})\n            code = item.get("fund_code")\n            name = item.get("fund_name")\n            age = calc_age(base.get("setup_date"), perf.get("nav_date"))\n\n            metric_rows.append({\n                "fund_code": code,\n                "return_1m": parse_pct_to_float(perf.get("1m")),\n                "return_3m": parse_pct_to_float(perf.get("3m")),\n                "return_6m": parse_pct_to_float(perf.get("6m")),\n                "return_1y": parse_pct_to_float(perf.get("1y")),\n                "return_3y": parse_pct_to_float(perf.get("3y")),\n                "return_5y": parse_pct_to_float(perf.get("5y")),\n            })\n\n            full_rows.append({\n                "fund_code": code,\n                "基金名称": name,\n                "基金代码": code,\n                "最新净值": perf.get("nav"),\n                "日增长率": perf.get("daily_growth_rate"),\n                "近1月": fmt_pct(perf.get("1m")),\n                "近3月": fmt_pct(perf.get("3m")),\n                "近6月": fmt_pct(perf.get("6m")),\n                "近1年": fmt_pct(perf.get("1y")),\n                "近3年": fmt_pct(perf.get("3y")),\n                "近5年": fmt_pct(perf.get("5y")),\n                "成立以来": fmt_pct(perf.get("since")),\n                "基金类型": base.get("fund_type"),\n                "风险等级": base.get("risk_level"),\n                "规模": base.get("assets_size"),\n                "基金经理": base.get("manager"),\n                "成立日期": base.get("setup_date"),\n                "净值日期": perf.get("nav_date"),\n                "成立年限": age,\n            })\n\n        metric_df = pd.DataFrame(metric_rows).set_index("fund_code")\n        full_df = pd.DataFrame(full_rows).set_index("fund_code")\n\n        scores = calc_score(metric_df, full_df["成立年限"])\n        full_df["收益表现评分"] = scores\n\n        out = full_df.reset_index(drop=True)\n        out = out.sort_values(by="收益表现评分", ascending=False)\n\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"收益表现评分_{ts}.xlsx")\n        out.to_excel(out_path, index=False)\n\n        # 美化 Excel\n        autosize_excel(out_path)\n\n        log(f"评分计算完成！共 {len(out)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n\n    except Exception as e:\n        log(f"收益表现评分出错: {e}")\n        return None\n\n\n\n# ========================================================\n# ⑥ 风险与回撤（离线：基于爬取的 JSON 数据）\n# ========================================================\n\nRISK_METRIC_EXPLAIN = {\n    "夏普比率":   "含义：单位总风险下的超额收益。\\n判断：越大越好。>1表示优秀。",\n    "卡玛比率":   "含义：年化收益与最大回撤的比值。\\n判断：越大越好。衡量抗风险性价比。",\n    "索提诺比率": "含义：单位下行风险下的超额收益。\\n判断：越大越好。剔除了上涨波动的影响。",\n    "年化收益":   "含义：将持有期收益换算为年度收益。\\n判断：越大越好。",\n    "最大回撤":   "含义：历史最大亏损幅度的纪录。\\n判断：越接近0%越好（绝对值越小越好）。",\n    "当前回撤":   "含义：最新净值距离前期高点的距离。\\n判断：0%代表正处于历史最高点。",\n    "回撤状态":   "含义：根据当前回撤深度划分的状态（高位/回调/深度）。",\n    "回撤进度":   "含义：当前回撤占历史最大回撤的比例。\\n判断：接近100%说明已到底部极限。",\n    "波动率":     "含义：净值波动的剧烈程度。\\n判断：越小说明持有过程越平稳。",\n    "下行波动":   "含义：仅统计亏损时的波动程度。\\n判断：越小越好。",\n    "溃疡指数":   "含义：结合回撤深度和持续时间的痛苦指数。\\n判断：越小越好，越低持有越舒服。",\n    "决策建议":   "含义：结合风险与收益给出的操作提示。"\n}\n\nRISK_TRADING_DAYS = 252\nRISK_FREE_RATE = 0.02\n\n\ndef _risk_semantics(mdd: float, sharpe: float, curr_dd: float):\n    mdd_abs = abs(mdd)\n    curr_dd_abs = abs(curr_dd)\n    if curr_dd_abs < 0.02:\n        status = "高位运行"\n    elif curr_dd_abs < 0.10:\n        status = "正常回调"\n    else:\n        status = "深度回撤"\n    progress = f"{curr_dd_abs / mdd_abs * 100:.1f}%" if mdd_abs > 0.001 else "0.0%"\n    if curr_dd_abs > mdd_abs * 0.8 and mdd_abs > 0.15:\n        advice = "极度超跌(博反弹)"\n    elif sharpe > 1.0 and curr_dd_abs < 0.03:\n        advice = "强者恒强(建议持有)"\n    else:\n        advice = "中性观望"\n    return {"回撤状态": status, "回撤进度": progress, "决策建议": advice}\n\n\ndef _calc_risk_metrics_from_history(nav_history):\n    """基于 [{\'date\':..., \'val\':...}, ...] 计算风险指标。"""\n    if not nav_history or len(nav_history) < 20:\n        return {}, None\n    df = pd.DataFrame(nav_history)\n    if "date" not in df.columns or "val" not in df.columns:\n        return {}, None\n    df[\'date\'] = pd.to_datetime(df[\'date\'], errors=\'coerce\')\n    df = df.dropna(subset=[\'date\']).sort_values(\'date\').reset_index(drop=True)\n    if len(df) < 20:\n        return {}, df\n\n    nav = df[\'val\'].ffill().dropna()\n    returns = nav.pct_change().dropna()\n    if returns.empty:\n        return {}, df\n\n    vol = returns.std() * np.sqrt(RISK_TRADING_DAYS)\n    drawdowns = nav / nav.cummax() - 1.0\n    mdd = drawdowns.min()\n    curr_dd = drawdowns.iloc[-1]\n    downside_returns = returns[returns < 0]\n    downside_dev = downside_returns.std() * np.sqrt(RISK_TRADING_DAYS) if not downside_returns.empty else 0\n    ulcer_index = np.sqrt((drawdowns ** 2).mean())\n\n    total_ret = (nav.iloc[-1] / nav.iloc[0]) - 1\n    days = (df[\'date\'].iloc[-1] - df[\'date\'].iloc[0]).days\n    ann_ret = (1 + total_ret) ** (365.25 / max(days, 1)) - 1\n    sharpe = (ann_ret - RISK_FREE_RATE) / vol if vol > 0 else 0\n    calmar = ann_ret / abs(mdd) if abs(mdd) > 0.001 else 0\n    sortino = (ann_ret - RISK_FREE_RATE) / downside_dev if downside_dev > 0 else 0\n\n    semantics = _risk_semantics(mdd, sharpe, curr_dd)\n    fmt_p = lambda x: f"{x * 100:.2f}%"\n    metrics = {\n        "年化收益": fmt_p(ann_ret),\n        "最大回撤": fmt_p(mdd),\n        "当前回撤": fmt_p(curr_dd),\n        "夏普比率": round(sharpe, 3),\n        "卡玛比率": round(calmar, 3),\n        "索提诺比率": round(sortino, 3),\n        "波动率": fmt_p(vol),\n        "下行波动": fmt_p(downside_dev),\n        "溃疡指数": round(ulcer_index, 4),\n        **semantics,\n    }\n    return metrics, df\n\n\ndef _beautify_risk_excel(path: str):\n    wb = load_workbook(path)\n    ws = wb["决策看板"] if "决策看板" in wb.sheetnames else wb.active\n    ws.freeze_panes = "A2"\n\n    red_font = Font(color="FF0000")\n    green_font = Font(color="00B050")\n    headers = [str(cell.value) for cell in ws[1]]\n\n    color_cols = {"年化收益", "最大回撤", "当前回撤", "夏普比率", "卡玛比率", "索提诺比率"}\n    bold_cols = {"决策建议", "夏普比率"}\n\n    for row in ws.iter_rows(min_row=2):\n        for cell in row:\n            col_name = headers[cell.column - 1]\n            cell.alignment = Alignment(wrap_text=True, vertical=\'center\', horizontal=\'left\')\n            if col_name in color_cols and cell.value not in (None, ""):\n                try:\n                    num_val = float(str(cell.value).replace(\'%\', \'\'))\n                    if num_val > 0:\n                        cell.font = red_font\n                    elif num_val < 0:\n                        cell.font = green_font\n                except Exception:\n                    pass\n            if col_name in bold_cols:\n                old_color = cell.font.color if cell.font else None\n                cell.font = Font(bold=True, color=old_color)\n\n    for cell in ws[1]:\n        new_font = copy(cell.font)\n        new_font.bold = True\n        cell.font = new_font\n        cell.alignment = Alignment(horizontal=\'center\', vertical=\'center\')\n        col_txt = str(cell.value)\n        if col_txt in RISK_METRIC_EXPLAIN:\n            comment = Comment(RISK_METRIC_EXPLAIN[col_txt], "RiskEngine")\n            comment.width = 250\n            comment.height = 80\n            cell.comment = comment\n\n    for col in ws.columns:\n        max_len = 0\n        col_letter = get_column_letter(col[0].column)\n        for cell in col:\n            if cell.value:\n                try:\n                    l = len(str(cell.value).encode(\'gbk\'))\n                except Exception:\n                    l = len(str(cell.value))\n                if l > max_len:\n                    max_len = l\n        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)\n\n    wb.save(path)\n\n\ndef run_risk_drawdown(log):\n    """风险与回撤：完全离线，从最新 fund_data/*.json 的 nav_history 计算。"""\n    try:\n        INPUT_DIR = "fund_data"\n        OUTPUT_DIR = "fund_excel"\n\n        files = glob.glob(os.path.join(INPUT_DIR, "*.json"))\n        if not files:\n            log("未找到 fund_data 目录下的 JSON 文件，请先运行【开始爬取】。")\n            return None\n\n        latest_file = max(files, key=os.path.getmtime)\n        log(f"正在处理: {latest_file}")\n\n        with open(latest_file, \'r\', encoding=\'utf-8\') as f:\n            data = json.load(f)\n        results = data if isinstance(data, list) else data.get("results", [])\n        if not results:\n            log("JSON 文件中无有效数据。")\n            return None\n\n        # 校验是否带有 nav_history\n        has_hist = sum(1 for r in results if r.get("nav_history"))\n        if has_hist == 0:\n            log("当前 JSON 中未包含历史净值（nav_history）。")\n            log("请重新点击【开始爬取】以生成含历史净值的数据文件，再运行本功能。")\n            return None\n        log(f"共加载 {len(results)} 只基金，其中 {has_hist} 只带有历史净值，开始计算风险指标...")\n\n        summary_list, history_map = [], {}\n        total = len(results)\n        for idx, item in enumerate(results, 1):\n            code = str(item.get("fund_code", "")).zfill(6)\n            name = item.get("fund_name", "--")\n            base = item.get("base_info", {}) or {}\n            hist = item.get("nav_history") or []\n\n            row = {\n                "代码": code,\n                "名称": name,\n                "类型": base.get("fund_type", "--"),\n                "规模": base.get("assets_size", "--"),\n                "经理": base.get("manager", "--"),\n            }\n            metrics, df_hist = _calc_risk_metrics_from_history(hist)\n            if metrics:\n                row.update(metrics)\n                if df_hist is not None and not df_hist.empty:\n                    history_map[code] = df_hist\n                log(f"[{idx}/{total}] {code} {name} 计算完毕")\n            else:\n                log(f"[{idx}/{total}] {code} {name} 数据不足，已跳过指标计算")\n            summary_list.append(row)\n\n        df_raw = pd.DataFrame(summary_list)\n        display_cols = [\n            "代码", "名称", "夏普比率", "卡玛比率", "索提诺比率",\n            "年化收益", "最大回撤", "当前回撤", "回撤状态", "回撤进度",\n            "决策建议", "波动率", "下行波动", "溃疡指数",\n            "类型", "规模", "经理",\n        ]\n        df_final = df_raw[[c for c in display_cols if c in df_raw.columns]].copy()\n        if "夏普比率" in df_final.columns:\n            df_final = df_final.sort_values("夏普比率", ascending=False, na_position="last")\n\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%m%d_%H%M")\n        out_path = os.path.join(OUTPUT_DIR, f"基金风险决策看板_{ts}.xlsx")\n\n        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:\n            df_final.to_excel(writer, sheet_name="决策看板", index=False)\n            for code, df_h in history_map.items():\n                df_h.to_excel(writer, sheet_name=f"{code}_历史", index=False)\n\n        _beautify_risk_excel(out_path)\n\n        log(f"风险与回撤分析完成！共 {len(df_final)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n\n    except Exception as e:\n        log(f"风险与回撤分析出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑦ 风险效率比率（Efficiency Ratios）\n# ========================================================\nEFF_RISK_FREE_RATE = 0.02\nEFF_MIN_AGE_MONTHS = 6\nEFF_YOUNG_AGE_MONTHS = 12\nEFF_YOUNG_PENALTY = 0.85\nEFF_WEIGHTS = {"sharpe": 40, "calmar": 35, "sortino": 25}\nEFF_METRIC_EXPLAIN = {\n    "风险效率评分": "含义：综合 Sharpe / Calmar / Sortino 三维效率的加权得分。\\n判断：分数越高（100为满分）代表风险回报性价比越高。",\n    "Sharpe估算": "含义：(年化收益 - 无风险利率) / 年化波动率估算。\\n判断：越大越好，>1 为良好。",\n    "Calmar估算": "含义：年化收益 / 最大回撤估算（绝对值）。\\n判断：越大越好，>2 为优秀。",\n    "Sortino估算": "含义：(年化收益 - 无风险利率) / 下行波动率估算。\\n判断：越大越好，>1.5 为良好。",\n    "成立年限": "提示：< 6 个月不评分，6~12 个月降权 × 0.85。",\n    "规模": "提示：规模过大可能影响流动性，过小存在清盘风险。",\n}\n\n\ndef _eff_parse_pct(val):\n    if val is None:\n        return None\n    if isinstance(val, (int, float)):\n        return float(val) / 100.0\n    s = str(val).strip().replace("%", "")\n    if s in ("", "--"):\n        return None\n    try:\n        return float(s) / 100.0\n    except ValueError:\n        return None\n\n\ndef _eff_parse_date(s):\n    try:\n        return datetime.datetime.strptime(str(s)[:10], "%Y-%m-%d").date()\n    except Exception:\n        return None\n\n\ndef _eff_age(setup, nav_date):\n    d0, d1 = _eff_parse_date(setup), _eff_parse_date(nav_date)\n    if not d0 or not d1:\n        return None, None\n    months = (d1 - d0).days / 30.44\n    return months, months / 12.0\n\n\ndef _eff_ann_return(perf, age_years):\n    r_1y = _eff_parse_pct(perf.get("1y"))\n    if r_1y is not None:\n        return r_1y\n    r_since = _eff_parse_pct(perf.get("since"))\n    if r_since is not None and age_years and age_years > 0:\n        try:\n            return (1 + r_since) ** (1.0 / age_years) - 1\n        except Exception:\n            pass\n    r_6m = _eff_parse_pct(perf.get("6m"))\n    if r_6m is not None:\n        return (1 + r_6m) ** 2 - 1\n    return None\n\n\ndef _eff_volatility(returns):\n    valid = [r for r in returns if r is not None]\n    if len(valid) < 2:\n        return None\n    import math as _m\n    n = len(valid)\n    mean = sum(valid) / n\n    variance = sum((r - mean) ** 2 for r in valid) / (n - 1)\n    return _m.sqrt(variance) * _m.sqrt(12)\n\n\ndef _eff_downside(returns):\n    import math as _m\n    valid = [r for r in returns if r is not None]\n    negatives = [r for r in valid if r < 0]\n    if len(negatives) >= 2:\n        n = len(negatives)\n        mean_neg = sum(negatives) / n\n        variance = sum((r - mean_neg) ** 2 for r in negatives) / (n - 1)\n        return _m.sqrt(variance) * _m.sqrt(12)\n    vol = _eff_volatility(valid)\n    return vol * 0.7 if vol else None\n\n\ndef _eff_max_dd(returns):\n    valid = [r for r in returns if r is not None]\n    if not valid:\n        return None\n    min_r = min(valid)\n    return abs(min_r) if min_r < 0 else min(valid) * 0.2\n\n\ndef _eff_metrics(perf, age_years):\n    period_returns = [_eff_parse_pct(perf.get(k)) for k in ("1m", "3m", "6m", "1y")]\n    r_ann = _eff_ann_return(perf, age_years)\n    sigma = _eff_volatility(period_returns)\n    dd = _eff_downside(period_returns)\n    mdd = _eff_max_dd(period_returns)\n    sharpe = (r_ann - EFF_RISK_FREE_RATE) / sigma if (r_ann is not None and sigma and sigma > 0) else None\n    calmar = r_ann / mdd if (r_ann is not None and mdd and mdd > 0) else None\n    sortino = (r_ann - EFF_RISK_FREE_RATE) / dd if (r_ann is not None and dd and dd > 0) else None\n    return {"sharpe": sharpe, "calmar": calmar, "sortino": sortino}\n\n\ndef _winsorize_series(s, q=0.01):\n    return s.clip(s.quantile(q), s.quantile(1 - q))\n\n\ndef _percentile_rank(s):\n    return s.rank(pct=True)\n\n\ndef _eff_score(metric_df, age_months_series):\n    import math as _m\n    pct_df = pd.DataFrame(index=metric_df.index)\n    for col in ["sharpe", "calmar", "sortino"]:\n        if col in metric_df.columns and metric_df[col].dropna().shape[0] > 1:\n            pct_df[col] = _percentile_rank(_winsorize_series(metric_df[col]))\n        elif col in metric_df.columns:\n            pct_df[col] = metric_df[col]\n    scores = {}\n    for code in metric_df.index:\n        age_m = age_months_series.get(code)\n        if age_m is None or age_m < EFF_MIN_AGE_MONTHS:\n            scores[code] = None\n            continue\n        total_w, s = 0.0, 0.0\n        for metric, w in EFF_WEIGHTS.items():\n            if metric not in pct_df.columns:\n                continue\n            v = pct_df.at[code, metric] if code in pct_df.index else None\n            if v is None or (isinstance(v, float) and _m.isnan(v)):\n                continue\n            total_w += w\n            s += v * w\n        if total_w == 0:\n            scores[code] = None\n        else:\n            score = s / total_w * 100\n            if age_m < EFF_YOUNG_AGE_MONTHS:\n                score *= EFF_YOUNG_PENALTY\n            scores[code] = round(score, 2)\n    return pd.Series(scores)\n\n\ndef _score_fill_color(score):\n    try:\n        v = float(score)\n    except (TypeError, ValueError):\n        return None\n    if v >= 80:\n        return "00B050"\n    if v >= 60:\n        return "92D050"\n    if v >= 40:\n        return "FFFF00"\n    return "FF4B4B"\n\n\ndef _beautify_with_explain(path, explain_map, score_col=None, colored_ratio_cols=None,\n                           header_fill_hex=None):\n    wb = load_workbook(path)\n    ws = wb.active\n    ws.freeze_panes = "A2"\n    headers = [str(cell.value) for cell in ws[1]]\n\n    header_font = Font(bold=True, color="FFFFFF" if header_fill_hex else None)\n    for cell in ws[1]:\n        cell.font = header_font if header_fill_hex else Font(bold=True)\n        if header_fill_hex:\n            cell.fill = PatternFill("solid", fgColor=header_fill_hex)\n        cell.alignment = Alignment(horizontal="center", vertical="center")\n        col_txt = str(cell.value)\n        if col_txt in explain_map:\n            text = explain_map[col_txt]\n            cmt = Comment(text, "FundSystem")\n            cmt.width = 320\n            cmt.height = 50 + (text.count(\'\\n\') * 25)\n            cell.comment = cmt\n\n    colored_ratio_cols = set(colored_ratio_cols or [])\n    for row in ws.iter_rows(min_row=2):\n        for cell in row:\n            col_name = headers[cell.column - 1]\n            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")\n            if score_col and col_name == score_col and cell.value not in (None, "", "--"):\n                color = _score_fill_color(cell.value)\n                if color:\n                    cell.fill = PatternFill("solid", fgColor=color)\n                cell.font = Font(bold=True)\n            if col_name in colored_ratio_cols and cell.value not in (None, "", "--"):\n                try:\n                    v = float(str(cell.value).replace("%", ""))\n                    cell.font = Font(color="FF0000" if v > 0 else "00B050")\n                except Exception:\n                    pass\n\n    for col in ws.columns:\n        max_len = 0\n        col_letter = get_column_letter(col[0].column)\n        for cell in col:\n            if cell.value:\n                try:\n                    curr = len(str(cell.value).encode("gbk"))\n                except Exception:\n                    curr = len(str(cell.value))\n                if curr > max_len:\n                    max_len = curr\n        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)\n    wb.save(path)\n\n\ndef _load_latest_json(log):\n    INPUT_DIR = "fund_data"\n    files = glob.glob(os.path.join(INPUT_DIR, "*.json"))\n    if not files:\n        log("未找到 fund_data 目录下的 JSON 文件，请先运行【开始爬取】。")\n        return None, None\n    latest = max(files, key=os.path.getmtime)\n    log(f"正在处理: {latest}")\n    try:\n        with open(latest, "r", encoding="utf-8") as f:\n            data = json.load(f)\n    except Exception as e:\n        log(f"读取 JSON 失败: {e}")\n        return None, None\n    results = data if isinstance(data, list) else data.get("results", [])\n    if not results:\n        log("JSON 文件中无有效数据。")\n        return None, None\n    return latest, results\n\n\ndef run_efficiency_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算风险效率评分...")\n\n        metric_rows, full_rows = [], []\n        for item in results:\n            perf = item.get("performance", {}) or {}\n            base = item.get("base_info", {}) or {}\n            code = item.get("fund_code", "")\n            age_m, age_y = _eff_age(base.get("setup_date"), perf.get("nav_date"))\n            metrics = _eff_metrics(perf, age_y)\n            metric_rows.append({"fund_code": code, **metrics, "age_months": age_m})\n            full_rows.append({\n                "fund_code": code,\n                "基金名称": item.get("fund_name", ""),\n                "基金代码": code,\n                "基金类型": base.get("fund_type"),\n                "风险等级": base.get("risk_level"),\n                "规模": base.get("assets_size"),\n                "基金经理": base.get("manager"),\n                "成立日期": base.get("setup_date"),\n                "净值日期": perf.get("nav_date"),\n                "成立年限": round(age_y, 2) if age_y is not None else None,\n                "Sharpe估算": f"{metrics[\'sharpe\']:.2f}" if metrics["sharpe"] is not None else "--",\n                "Calmar估算": f"{metrics[\'calmar\']:.2f}" if metrics["calmar"] is not None else "--",\n                "Sortino估算": f"{metrics[\'sortino\']:.2f}" if metrics["sortino"] is not None else "--",\n            })\n\n        metric_df = pd.DataFrame(metric_rows).set_index("fund_code")\n        full_df = pd.DataFrame(full_rows).set_index("fund_code")\n        scores = _eff_score(metric_df[["sharpe", "calmar", "sortino"]], metric_df["age_months"])\n        full_df["风险效率评分"] = scores\n\n        ordered = ["基金名称", "风险效率评分", "基金代码", "基金类型", "风险等级",\n                   "规模", "基金经理", "成立日期", "净值日期", "成立年限",\n                   "Sharpe估算", "Calmar估算", "Sortino估算"]\n        out = full_df.reset_index(drop=True)\n        out = out[[c for c in ordered if c in out.columns]]\n        out = out.sort_values(by="风险效率评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"风险效率评分_{ts}.xlsx")\n        out.to_excel(out_path, index=False)\n        _beautify_with_explain(out_path, EFF_METRIC_EXPLAIN, score_col="风险效率评分",\n                               colored_ratio_cols=["Sharpe估算", "Calmar估算", "Sortino估算"])\n        log(f"风险效率评分计算完成！共 {len(out)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"风险效率评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑧ 位置 / 估值 / 情绪（Position）\n# ========================================================\nPOS_WEIGHTS = {"nav_pos_score": 0.5, "dd_score": 0.3, "since_pos_score": 0.2}\nPOS_MIN_AGE_MONTHS = 1\nPOS_PERIOD_FIELDS = ["1m", "3m", "6m", "1y", "3y", "5y"]\nPOS_METRIC_EXPLAIN = {\n    "位置评分": "含义：综合净值位置/回撤/成立以来涨幅的加权得分。\\n判断：分数越高代表当前买入赔率越高（低位），≥80 为极低位安全区。",\n    "净值位置分位": "含义：当前净值在历史推算节点中的百分比位置。\\n判断：越小代表越接近历史底部。",\n    "回撤估算": "含义：基于多期收益推算出的最大可能回撤深度。\\n判断：数值越大代表跌得越深，反弹空间理论上越大。",\n    "成立以来收益": "含义：自成立以来的累计回报。\\n判断：数值越小（甚至负数）代表在同类中处于更低的位置。",\n    "成立年限": "提示：年限过短（<6个月）的基金评分会进行折算降权。"\n}\n\n\ndef _pos_safe_float(val, default=0.0):\n    if pd.isna(val) or val in ("--", "", "None", "null"):\n        return default\n    try:\n        return float(str(val).replace(",", "").strip())\n    except (ValueError, TypeError):\n        return default\n\n\ndef _pos_parse_pct(val):\n    if pd.isna(val) or val in ("--", "", "None", "null"):\n        return np.nan\n    if isinstance(val, (int, float)):\n        return float(val) / 100.0\n    try:\n        return float(str(val).replace("%", "").strip()) / 100.0\n    except (ValueError, TypeError):\n        return np.nan\n\n\ndef _pos_calc_nav_stats(row):\n    nav_now = row.get("nav_raw")\n    if not nav_now or nav_now <= 0 or pd.isna(nav_now):\n        return pd.Series([np.nan, np.nan])\n    hist_navs = [nav_now]\n    returns = []\n    age = row.get("age_months", 0) or 0\n    limits = {"1m": 1, "3m": 3, "6m": 6, "1y": 12, "3y": 36, "5y": 60}\n    for f in POS_PERIOD_FIELDS:\n        if age >= limits.get(f, 0):\n            r = row.get(f"r_{f}")\n            if not pd.isna(r) and r > -0.99:\n                hist_navs.append(nav_now / (1 + r))\n                returns.append(r)\n    if len(hist_navs) < 2:\n        return pd.Series([np.nan, np.nan])\n    pos_val = (sorted(hist_navs).index(nav_now)) / (len(hist_navs) - 1)\n    min_r = min(returns) if returns else 0\n    dd_val = abs(min_r) if min_r < 0 else 0\n    return pd.Series([pos_val, dd_val])\n\n\ndef run_position_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算位置/估值/情绪评分...")\n\n        rows = []\n        for item in results:\n            p = item.get("performance", {}) or {}\n            b = item.get("base_info", {}) or {}\n            try:\n                d0 = pd.to_datetime(b.get("setup_date"))\n                d1 = pd.to_datetime(p.get("nav_date"))\n                age = (d1 - d0).days / 30.44\n            except Exception:\n                age = 0\n            row = {\n                "code": item.get("fund_code"),\n                "name": item.get("fund_name"),\n                "type": b.get("fund_type", "未知"),\n                "age_months": age,\n                "nav_raw": _pos_safe_float(p.get("nav")),\n                "r_since": _pos_parse_pct(p.get("since")),\n                "manager": b.get("manager"),\n                "size": b.get("assets_size"),\n            }\n            for f in POS_PERIOD_FIELDS:\n                row[f"r_{f}"] = _pos_parse_pct(p.get(f))\n            rows.append(row)\n\n        df = pd.DataFrame(rows)\n        if df.empty:\n            log("无有效数据，已终止。")\n            return None\n\n        df[["nav_pos", "dd_proxy"]] = df.apply(_pos_calc_nav_stats, axis=1)\n        df["nav_score_rank"] = df.groupby("type")["nav_pos"].rank(pct=True, ascending=False)\n        df["since_score_rank"] = df.groupby("type")["r_since"].rank(pct=True, ascending=False)\n        df["dd_score_rank"] = df.groupby("type")["dd_proxy"].rank(pct=True, ascending=True)\n\n        def final_score(r):\n            if (r["age_months"] or 0) < POS_MIN_AGE_MONTHS:\n                return np.nan\n            if pd.isna(r["nav_score_rank"]):\n                return np.nan\n            raw_s = (\n                r["nav_score_rank"] * POS_WEIGHTS["nav_pos_score"]\n                + r["dd_score_rank"] * POS_WEIGHTS["dd_score"]\n                + r["since_score_rank"] * POS_WEIGHTS["since_pos_score"]\n            ) * 100\n            if r["age_months"] < 3:\n                raw_s *= 0.8\n            elif r["age_months"] < 6:\n                raw_s *= 0.9\n            return round(raw_s, 2)\n\n        df["位置评分"] = df.apply(final_score, axis=1)\n\n        out_map = {\n            "name": "基金名称", "位置评分": "位置评分", "code": "基金代码",\n            "type": "基金类型", "manager": "基金经理", "size": "规模",\n            "nav_raw": "当前净值", "nav_pos": "净值位置分位",\n            "dd_proxy": "回撤估算", "r_since": "成立以来收益"\n        }\n        final_df = df[list(out_map.keys())].rename(columns=out_map)\n        for col in ["净值位置分位", "回撤估算", "成立以来收益"]:\n            final_df[col] = final_df[col].apply(lambda x: f"{round(x * 100, 2)}%" if pd.notnull(x) else "--")\n        final_df = final_df.sort_values("位置评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"位置评分_{ts}.xlsx")\n        final_df.to_excel(out_path, index=False)\n        _beautify_with_explain(out_path, POS_METRIC_EXPLAIN, score_col="位置评分",\n                               colored_ratio_cols=["成立以来收益"], header_fill_hex="203764")\n        log(f"位置/估值/情绪评分完成！共 {len(final_df)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"位置/估值/情绪评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑨ 趋势与择时（Timing）\n# ========================================================\nTIMING_WEIGHTS = {"m1_rank": 25, "m3_rank": 20, "m6_rank": 15, "dir_score_rank": 20}\nTIMING_OVERHEAT_PENALTY = 15.0\nTIMING_RECOVERY_BONUS = 20.0\nTIMING_MIN_AGE_MONTHS = 1\nTIMING_METRIC_EXPLAIN = {\n    "趋势评分": "含义：综合动量/方向/信号的加权得分。\\n判断：分数越高代表趋势越强。需结合\'位置评分\'使用（高趋势+低位置=最优）。",\n    "趋势方向": "含义：综合1m/3m/6m收益方向判定。\\n状态：强势上行/温和上行/短期回调/中期震荡/持续下行。",\n    "过热信号": "逻辑：1m>15% 且 3m>30%。\\n提示：⚠️代表短期涨幅过载，存在回调风险，评分已扣除15分。",\n    "修复信号": "逻辑：6m<-10% 且 1m&3m反转向上。\\n提示：✅代表超跌后的反弹起步，评分已加20分奖励。"\n}\n\n\ndef _timing_parse_pct(val):\n    if pd.isna(val) or val in ("--", ""):\n        return np.nan\n    if isinstance(val, (int, float)):\n        return float(val) / 100.0\n    try:\n        return float(str(val).replace("%", "").strip()) / 100.0\n    except Exception:\n        return np.nan\n\n\ndef _timing_trend_details(row):\n    r1, r3, r6 = row["m1"], row["m3"], row["m6"]\n    score, label = 0.5, "震荡 ~"\n    if pd.isna(r1) or pd.isna(r3):\n        return pd.Series([score, "数据不足"])\n    if r1 > 0 and r3 > 0:\n        if not pd.isna(r6) and r6 > 0:\n            score, label = 1.0, "强势上行 ↑↑↑"\n        else:\n            score, label = 0.75, "温和上行 ↑↑"\n    elif r1 > 0 and r3 <= 0:\n        score, label = 0.5, "短期反弹 ↑"\n    elif r1 <= 0 and r3 > 0:\n        score, label = 0.6, "短期回调 ↓↑"\n    elif r1 < 0 and r3 < 0:\n        if not pd.isna(r6) and r6 < 0:\n            score, label = 0.1, "持续下行 ↓↓↓"\n        else:\n            score, label = 0.2, "下行趋势 ↓↓"\n    return pd.Series([score, label])\n\n\ndef _timing_signals(row):\n    r1, r3, r6 = row["m1"], row["m3"], row["m6"]\n    overheat = (not pd.isna(r1) and not pd.isna(r3) and r1 > 0.15 and r3 > 0.30)\n    recovery = (not pd.isna(r6) and not pd.isna(r1) and not pd.isna(r3)\n                and r6 < -0.10 and r1 > 0 and r3 > 0)\n    return pd.Series([overheat, recovery])\n\n\ndef _beautify_timing_excel(path):\n    wb = load_workbook(path)\n    ws = wb.active\n    ws.freeze_panes = "A2"\n    headers = [str(cell.value) for cell in ws[1]]\n    header_fill = PatternFill("solid", fgColor="44546A")\n    header_font = Font(bold=True, color="FFFFFF")\n\n    for cell in ws[1]:\n        cell.fill, cell.font = header_fill, header_font\n        cell.alignment = Alignment(horizontal="center")\n        if cell.value in TIMING_METRIC_EXPLAIN:\n            text = TIMING_METRIC_EXPLAIN[cell.value]\n            cmt = Comment(text, "FundSystem")\n            cmt.width = 320\n            cmt.height = 50 + (text.count("\\n") * 25)\n            cell.comment = cmt\n\n    for row in ws.iter_rows(min_row=2):\n        row_cells = {headers[i]: cell for i, cell in enumerate(row)}\n        score_cell = row_cells.get("趋势评分")\n        if score_cell and score_cell.value not in (None, "", "--"):\n            try:\n                v = float(score_cell.value)\n                color = _score_fill_color(v)\n                if color:\n                    score_cell.fill = PatternFill("solid", fgColor=color)\n                score_cell.font = Font(bold=True)\n            except Exception:\n                pass\n        oh = row_cells.get("过热信号")\n        if oh and "过热" in str(oh.value):\n            oh.font = Font(color="FFFFFF", bold=True)\n            oh.fill = PatternFill("solid", fgColor="C00000")\n        rec = row_cells.get("修复信号")\n        if rec and "修复" in str(rec.value):\n            rec.font = Font(color="006100", bold=True)\n            rec.fill = PatternFill("solid", fgColor="C6EFCE")\n        for m in ["近1月", "近3月", "近6月"]:\n            cell = row_cells.get(m)\n            if cell and cell.value not in (None, "--"):\n                try:\n                    val = float(str(cell.value).replace("%", ""))\n                    cell.font = Font(color="FF0000" if val > 0 else "00B050")\n                except Exception:\n                    pass\n\n    for col in ws.columns:\n        ws.column_dimensions[get_column_letter(col[0].column)].width = 15\n    wb.save(path)\n\n\ndef run_timing_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算趋势与择时评分...")\n\n        rows = []\n        for item in results:\n            p = item.get("performance", {}) or {}\n            b = item.get("base_info", {}) or {}\n            try:\n                age = (pd.to_datetime(p.get("nav_date")) - pd.to_datetime(b.get("setup_date"))).days / 30.44\n            except Exception:\n                age = 0\n            rows.append({\n                "code": item.get("fund_code"), "name": item.get("fund_name"),\n                "type": b.get("fund_type", "未知"), "age": age,\n                "m1": _timing_parse_pct(p.get("1m")),\n                "m3": _timing_parse_pct(p.get("3m")),\n                "m6": _timing_parse_pct(p.get("6m")),\n                "manager": b.get("manager"), "size": b.get("assets_size"),\n            })\n        df = pd.DataFrame(rows)\n        if df.empty:\n            log("无有效数据。")\n            return None\n\n        df[["dir_score", "dir_label"]] = df.apply(_timing_trend_details, axis=1)\n        df[["is_oh", "is_rec"]] = df.apply(_timing_signals, axis=1)\n        for col in ["m1", "m3", "m6", "dir_score"]:\n            df[f"{col}_rank"] = df.groupby("type")[col].rank(pct=True)\n\n        def calc_final(r):\n            if (r["age"] or 0) < TIMING_MIN_AGE_MONTHS:\n                return np.nan\n            weights = dict(TIMING_WEIGHTS)\n            if r["age"] < 3:\n                for k in ["m3_rank", "m6_rank", "dir_score_rank"]:\n                    weights.pop(k, None)\n            elif r["age"] < 6:\n                weights.pop("m6_rank", None)\n            try:\n                score_sum = sum((r[k] if not pd.isna(r[k]) else 0) * w for k, w in weights.items())\n                weight_sum = sum(weights.values())\n                base = (score_sum / weight_sum) * 100 if weight_sum else 0\n            except KeyError:\n                return 0.0\n            if r["is_oh"]:\n                base -= TIMING_OVERHEAT_PENALTY\n            if r["is_rec"]:\n                base += TIMING_RECOVERY_BONUS\n            if r["age"] < 6:\n                base *= 0.9\n            return round(max(0, min(100, base)), 2)\n\n        df["趋势评分"] = df.apply(calc_final, axis=1)\n\n        out_cols = {\n            "name": "基金名称", "趋势评分": "趋势评分", "code": "基金代码", "type": "基金类型",\n            "m1": "近1月", "m3": "近3月", "m6": "近6月", "dir_label": "趋势方向",\n            "is_oh": "过热信号", "is_rec": "修复信号", "manager": "基金经理", "age": "成立年限",\n        }\n        f_df = df[list(out_cols.keys())].rename(columns=out_cols)\n        for c in ["近1月", "近3月", "近6月"]:\n            f_df[c] = f_df[c].apply(lambda x: f"{x * 100:.2f}%" if not pd.isna(x) else "--")\n        f_df["过热信号"] = f_df["过热信号"].map({True: "⚠️ 过热", False: "正常"})\n        f_df["修复信号"] = f_df["修复信号"].map({True: "✅ 修复", False: "--"})\n        f_df["成立年限"] = (f_df["成立年限"] / 12).round(2)\n        f_df = f_df.sort_values("趋势评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"趋势评分_{ts}.xlsx")\n        f_df.to_excel(out_path, index=False)\n        _beautify_timing_excel(out_path)\n        log(f"趋势与择时评分完成！共 {len(f_df)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"趋势与择时评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑩ 基金经理能力（Manager）\n# ========================================================\nMGR_WEIGHTS_NORMAL = {"tenure_years": 0.4, "tenure_ann_return": 0.5, "bonus": 10}\nMGR_WEIGHTS_YOUNG = {"tenure_years": 0.6, "tenure_ann_return": 0.3, "bonus": 10}\nMGR_MIN_AGE_MONTHS = 3\nMGR_BONUS_RULES = [(5.0, 10), (3.0, 7), (1.0, 4), (0.0, 0)]\nMGR_METRIC_EXPLAIN = {\n    "经理评分": "含义：综合经理任职年限、任期年化回报及稳定性加分的加权得分。\\n判断：分数越高代表经理管理经验越丰富且历史业绩越稳健。",\n    "任职年限估算": "含义：基于基金成立日推算的任职时长。\\n判断：⭐ 代表任职 ≥ 5年的资深经理，具备穿越牛熊的经验。",\n    "任期年化回报": "含义：将任期总回报折算为年化收益率。\\n判断：排除时间干扰，更客观地对比不同任期经理的赚钱能力。",\n    "基金经理": "提示：若经理评分高但任职年限短，需警惕其业绩的偶然性。"\n}\n\n\ndef _mgr_parse_pct(val):\n    if pd.isna(val) or val in ("--", ""):\n        return np.nan\n    if isinstance(val, (int, float)):\n        return float(val) / 100.0\n    try:\n        return float(str(val).replace("%", "").strip()) / 100.0\n    except Exception:\n        return np.nan\n\n\ndef _mgr_ann_return(total_return, years):\n    if pd.isna(total_return) or years <= 0:\n        return np.nan\n    try:\n        return (1 + total_return) ** (1 / years) - 1\n    except Exception:\n        return np.nan\n\n\ndef _mgr_bonus(years):\n    for threshold, score in MGR_BONUS_RULES:\n        if years >= threshold:\n            return score\n    return 0.0\n\n\ndef _beautify_manager_excel(path):\n    wb = load_workbook(path)\n    ws = wb.active\n    ws.freeze_panes = "A2"\n    headers = [str(cell.value) for cell in ws[1]]\n    header_fill = PatternFill("solid", fgColor="203764")\n    header_font = Font(bold=True, color="FFFFFF")\n\n    for cell in ws[1]:\n        cell.fill, cell.font = header_fill, header_font\n        cell.alignment = Alignment(horizontal="center", vertical="center")\n        if cell.value in MGR_METRIC_EXPLAIN:\n            text = MGR_METRIC_EXPLAIN[cell.value]\n            cmt = Comment(text, "FundSystem")\n            cmt.width, cmt.height = 320, 50 + (text.count("\\n") * 25)\n            cell.comment = cmt\n\n    for row in ws.iter_rows(min_row=2):\n        rmap = {headers[i]: cell for i, cell in enumerate(row)}\n        score_cell = rmap.get("经理评分")\n        if score_cell and score_cell.value not in (None, "", "--"):\n            try:\n                v = float(score_cell.value)\n                color = _score_fill_color(v)\n                if color:\n                    score_cell.fill = PatternFill("solid", fgColor=color)\n                score_cell.font = Font(bold=True)\n            except Exception:\n                pass\n        tenure_cell = rmap.get("任职年限估算")\n        if tenure_cell and "⭐" in str(tenure_cell.value):\n            tenure_cell.font = Font(color="C65911", bold=True)\n        ret_cell = rmap.get("任期年化回报")\n        if ret_cell and ret_cell.value not in (None, "", "--"):\n            try:\n                val = float(str(ret_cell.value).replace("%", ""))\n                ret_cell.font = Font(color="FF0000" if val > 0 else "00B050")\n            except Exception:\n                pass\n\n    for col in ws.columns:\n        ws.column_dimensions[get_column_letter(col[0].column)].width = 16\n    wb.save(path)\n\n\ndef run_manager_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算基金经理能力评分...")\n\n        rows = []\n        for item in results:\n            p = item.get("performance", {}) or {}\n            b = item.get("base_info", {}) or {}\n            try:\n                d0 = pd.to_datetime(b.get("setup_date"))\n                d1 = pd.to_datetime(p.get("nav_date"))\n                tenure_y = (d1 - d0).days / 365.25\n            except Exception:\n                tenure_y = 0\n            r_total = _mgr_parse_pct(p.get("since"))\n            ann_ret = _mgr_ann_return(r_total, tenure_y)\n            rows.append({\n                "code": item.get("fund_code"), "name": item.get("fund_name"),\n                "type": b.get("fund_type", "未知"), "manager": b.get("manager"),\n                "tenure_y": tenure_y, "ann_ret": ann_ret, "r_total": r_total,\n                "size": b.get("assets_size"), "setup_date": b.get("setup_date"),\n            })\n        df = pd.DataFrame(rows)\n        if df.empty:\n            log("无有效数据。")\n            return None\n\n        df["y_rank"] = df.groupby("type")["tenure_y"].rank(pct=True)\n        df["ret_rank"] = df.groupby("type")["ann_ret"].rank(pct=True)\n\n        def calc_score(r):\n            age_m = (r["tenure_y"] or 0) * 12\n            if age_m < MGR_MIN_AGE_MONTHS:\n                return np.nan\n            w = MGR_WEIGHTS_YOUNG if age_m < 12 else MGR_WEIGHTS_NORMAL\n            y_rank = r["y_rank"] if not pd.isna(r["y_rank"]) else 0\n            ret_rank = r["ret_rank"] if not pd.isna(r["ret_rank"]) else 0\n            base = (y_rank * w["tenure_years"] + ret_rank * w["tenure_ann_return"]) * 100\n            total = base + _mgr_bonus(r["tenure_y"])\n            if age_m < 6:\n                total *= 0.85\n            elif age_m < 12:\n                total *= 0.92\n            return round(max(0, min(100, total)), 2)\n\n        df["经理评分"] = df.apply(calc_score, axis=1)\n\n        def tenure_label(y):\n            if pd.isna(y):\n                return "--"\n            label = f"{y:.1f}年"\n            if y >= 5:\n                return f"⭐ {label}(资深)"\n            if y >= 3:\n                return f"{label}(熟练)"\n            return f"{label}(新锐)"\n\n        out_cols = {\n            "name": "基金名称", "经理评分": "经理评分", "code": "基金代码",\n            "manager": "基金经理", "type": "基金类型", "size": "规模",\n            "tenure_y": "任职年限估算", "ann_ret": "任期年化回报", "r_total": "任期总回报",\n            "setup_date": "成立日期",\n        }\n        f_df = df[list(out_cols.keys())].rename(columns=out_cols)\n        f_df["任职年限估算"] = f_df["任职年限估算"].apply(tenure_label)\n        f_df["任期年化回报"] = f_df["任期年化回报"].apply(lambda x: f"{x * 100:.2f}%" if not pd.isna(x) else "--")\n        f_df["任期总回报"] = f_df["任期总回报"].apply(lambda x: f"{x * 100:.2f}%" if not pd.isna(x) else "--")\n        f_df = f_df.sort_values("经理评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"经理评分_{ts}.xlsx")\n        f_df.to_excel(out_path, index=False)\n        _beautify_manager_excel(out_path)\n        log(f"基金经理能力评分完成！共 {len(f_df)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"基金经理能力评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑪ 成本、规模与结构（Cost / Size / Structure）\n# ========================================================\nCOST_WEIGHTS = {"fee": 0.4, "size": 0.4, "access": 0.2}\nCOST_METRIC_EXPLAIN = {\n    "成本评分": "含义：综合持有成本与操作灵活性的得分。\\n判断：≥80 为极致性价比，<40 需警惕高费率或清盘风险。",\n    "申购费率": "含义：买入时支付的一次性费率。\\n判断：0.00% 为满分；通常 0.1%~0.15% 为互联网代销标配。",\n    "规模(亿)": "含义：基金最新资产净值（亿元）。\\n判断：5~50亿为最优区间，兼顾流动性与业绩弹性。",\n    "规模评级": "⚠️ 警告：极小(0.5亿以下)存在清盘风险；超大(500亿以上)可能存在仓位调动困难。",\n    "买入状态": "判断：\'开放申购\'以外的状态通常代表由于额度或基金经理策略限制，无法正常买入。"\n}\n\n\ndef _cost_parse_fee(val):\n    if pd.isna(val) or val in ("--", ""):\n        return np.nan\n    s = str(val).replace("%", "").strip()\n    try:\n        return float(s)\n    except Exception:\n        return np.nan\n\n\ndef _cost_parse_size(val):\n    if pd.isna(val) or val in ("--", ""):\n        return 0.0\n    s = str(val).strip()\n    m = re.search(r"([\\d,.]+)", s)\n    if not m:\n        return 0.0\n    try:\n        num = float(m.group(1).replace(",", ""))\n        if "万" in s and "亿" not in s:\n            num /= 10000.0\n        return num\n    except Exception:\n        return 0.0\n\n\ndef _cost_score_fee(fee):\n    if pd.isna(fee):\n        return 50.0\n    if fee <= 0:\n        return 100.0\n    if fee <= 0.10:\n        return 80.0\n    if fee <= 0.15:\n        return 70.0\n    if fee <= 0.50:\n        return 40.0\n    return 20.0\n\n\ndef _cost_score_size(size):\n    if size >= 500:\n        return 50.0, "超大(大象转身难)"\n    if size >= 100:\n        return 75.0, "较大"\n    if size >= 50:\n        return 90.0, "略大(良好)"\n    if size >= 5:\n        return 100.0, "最优区间 ✓"\n    if size >= 2:\n        return 75.0, "偏小"\n    if size >= 0.5:\n        return 40.0, "过小(清盘风险)"\n    return 20.0, "极小(高清盘风险⚠️)"\n\n\ndef _cost_score_access(status):\n    if status is None or pd.isna(status):\n        return 50.0\n    s = str(status)\n    if "开放申购" in s:\n        return 100.0\n    if "暂停" in s or "--" in s:\n        return 30.0\n    return 60.0\n\n\ndef _beautify_cost_excel(path):\n    wb = load_workbook(path)\n    ws = wb.active\n    ws.freeze_panes = "A2"\n    headers = [str(cell.value) for cell in ws[1]]\n    header_fill = PatternFill("solid", fgColor="333333")\n    header_font = Font(bold=True, color="FFFFFF")\n\n    for cell in ws[1]:\n        cell.fill, cell.font = header_fill, header_font\n        cell.alignment = Alignment(horizontal="center", vertical="center")\n        if cell.value in COST_METRIC_EXPLAIN:\n            text = COST_METRIC_EXPLAIN[cell.value]\n            cmt = Comment(text, "FundSystem")\n            cmt.width, cmt.height = 320, 60 + (text.count("\\n") * 25)\n            cell.comment = cmt\n\n    for row in ws.iter_rows(min_row=2):\n        rmap = {headers[i]: cell for i, cell in enumerate(row)}\n        score_cell = rmap.get("成本评分")\n        if score_cell and score_cell.value not in (None, "", "--"):\n            try:\n                v = float(score_cell.value)\n                color = _score_fill_color(v)\n                if color:\n                    score_cell.fill = PatternFill("solid", fgColor=color)\n                score_cell.font = Font(bold=True)\n            except Exception:\n                pass\n        rating_cell = rmap.get("规模评级")\n        if rating_cell and "风险" in str(rating_cell.value):\n            rating_cell.font = Font(color="FFFFFF", bold=True)\n            rating_cell.fill = PatternFill("solid", fgColor="C00000")\n        fee_cell = rmap.get("申购费率")\n        if fee_cell and fee_cell.value not in (None, "", "--"):\n            try:\n                v = float(str(fee_cell.value).replace("%", ""))\n                if v <= 0:\n                    fee_cell.font = Font(color="006100", bold=True)\n                    fee_cell.fill = PatternFill("solid", fgColor="C6EFCE")\n            except Exception:\n                pass\n\n    for col in ws.columns:\n        ws.column_dimensions[get_column_letter(col[0].column)].width = 16\n    wb.save(path)\n\n\ndef run_cost_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算成本/规模/结构评分...")\n\n        rows = []\n        for item in results:\n            b = item.get("base_info", {}) or {}\n            s = item.get("status", {}) or {}\n            fee_raw = _cost_parse_fee(s.get("buy_fee"))\n            size_raw = _cost_parse_size(b.get("assets_size"))\n            status_raw = s.get("buy_status")\n\n            f_score = _cost_score_fee(fee_raw)\n            s_score, s_label = _cost_score_size(size_raw)\n            a_score = _cost_score_access(status_raw)\n\n            total_score = (f_score * COST_WEIGHTS["fee"]\n                           + s_score * COST_WEIGHTS["size"]\n                           + a_score * COST_WEIGHTS["access"])\n\n            rows.append({\n                "code": item.get("fund_code"),\n                "name": item.get("fund_name"),\n                "成本评分": round(total_score, 2),\n                "申购费率": f"{fee_raw:.3f}%" if not pd.isna(fee_raw) else "--",\n                "规模(亿)": round(size_raw, 2),\n                "规模评级": s_label,\n                "买入状态": status_raw or "--",\n                "type": b.get("fund_type", "未知"),\n                "manager": b.get("manager"),\n                "company": b.get("company", "--"),\n            })\n\n        df = pd.DataFrame(rows)\n        if df.empty:\n            log("无有效数据。")\n            return None\n        df = df.sort_values("成本评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"成本评分_{ts}.xlsx")\n        df.to_excel(out_path, index=False)\n        _beautify_cost_excel(out_path)\n        log(f"成本/规模/结构评分完成！共 {len(df)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"成本/规模/结构评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑫ 归因分析（Attribution）\n# ========================================================\nATT_PERIODS = ["1m", "3m", "6m", "1y"]\nATT_PW = {"1m": 1, "3m": 2, "6m": 3, "1y": 4}\nATT_WEIGHTS = {"alpha": 50, "beta_score": 30, "consistency": 20}\nATT_METRIC_EXPLAIN = {\n    "归因评分":   "含义：综合 Alpha / Beta / 一致性的加权得分。\\n判断：越高代表主动管理能力越强。",\n    "Alpha估算":  "含义：扣除同类中位数贡献后的超额收益。\\n判断：>0 代表跑赢同类平均。",\n    "Beta估算":   "含义：对同类中位数的敏感度。\\n判断：1.0 同步；>1.2 高进攻；<0.8 防御。",\n    "收益一致性": "含义：超额收益为正的频率。\\n判断：越高代表业绩越稳定。",\n}\n\n\ndef _to_monthly(r, period):\n    if pd.isna(r) or r <= -1:\n        return np.nan\n    months = {"1m": 1, "3m": 3, "6m": 6, "1y": 12}\n    return (1 + r) ** (1 / months.get(period, 1)) - 1\n\n\ndef _beta_score_fn(beta):\n    if pd.isna(beta):\n        return 0.0\n    import math as _m\n    return _m.exp(-2 * (beta - 0.9) ** 2)\n\n\ndef _compute_attribution(results):\n    """归因核心计算，返回 (scores_series, detail_df[code -> Alpha/Beta/Consistency])"""\n    rows = []\n    for it in results:\n        perf = it.get("performance", {}) or {}\n        base = it.get("base_info", {}) or {}\n        try:\n            age = (pd.to_datetime(perf.get("nav_date")) - pd.to_datetime(base.get("setup_date"))).days / 30.44\n        except Exception:\n            age = 0\n        row = {"code": it.get("fund_code"),\n               "type": base.get("fund_type", "未知"),\n               "age": age}\n        for p in ATT_PERIODS:\n            row[p] = _to_monthly(_timing_parse_pct(perf.get(p)), p)\n        rows.append(row)\n\n    df = pd.DataFrame(rows)\n    if df.empty:\n        return pd.Series(dtype=float), pd.DataFrame(columns=["Alpha", "Beta", "Consistency"])\n\n    market_med = df.groupby("type")[ATT_PERIODS].median()\n\n    def _metric(r):\n        try:\n            med = market_med.loc[r["type"]]\n        except KeyError:\n            return pd.Series({"Alpha": np.nan, "Consistency": np.nan,\n                              "Beta": np.nan, "beta_score": 0.0})\n        a_sum, w_sum, pos, total, betas = 0, 0, 0, 0, []\n        limits = {"1m": 0, "3m": 3, "6m": 6, "1y": 12}\n        for p in ATT_PERIODS:\n            if r["age"] < limits[p]:\n                continue\n            rf, rm = r[p], med[p]\n            if pd.isna(rf) or pd.isna(rm):\n                continue\n            w = ATT_PW[p]\n            a_sum += (rf - rm) * w\n            w_sum += w\n            total += 1\n            if (rf - rm) > 0:\n                pos += 1\n            if abs(rm) > 1e-6:\n                betas.append(rf / rm)\n        beta = float(np.median(betas)) if betas else np.nan\n        return pd.Series({\n            "Alpha": a_sum / w_sum if w_sum > 0 else np.nan,\n            "Consistency": pos / total if total > 0 else np.nan,\n            "Beta": beta,\n            "beta_score": _beta_score_fn(beta),\n        })\n\n    metrics = df.apply(_metric, axis=1)\n    df = pd.concat([df, metrics], axis=1)\n\n    for col in ["Alpha", "beta_score", "Consistency"]:\n        df[f"{col}_rank"] = df.groupby("type")[col].rank(pct=True)\n\n    scores = {}\n    for _, r in df.iterrows():\n        code = r["code"]\n        if r["age"] < 3:\n            scores[code] = None\n            continue\n        s = (float(r.get("Alpha_rank") or 0) * ATT_WEIGHTS["alpha"]\n             + float(r.get("beta_score_rank") or 0) * ATT_WEIGHTS["beta_score"]\n             + float(r.get("Consistency_rank") or 0) * ATT_WEIGHTS["consistency"])\n        if r["age"] < 6:\n            s *= 0.8\n        elif r["age"] < 12:\n            s *= 0.9\n        scores[code] = round(s, 2)\n\n    detail = df.set_index("code")[["Alpha", "Beta", "Consistency"]]\n    return pd.Series(scores), detail\n\n\ndef run_attribution_score(log):\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，开始计算归因分析评分...")\n\n        scores, detail = _compute_attribution(results)\n        rows = []\n        for item in results:\n            code = item.get("fund_code")\n            base = item.get("base_info", {}) or {}\n            alpha = detail["Alpha"].get(code) if not detail.empty else np.nan\n            beta = detail["Beta"].get(code) if not detail.empty else np.nan\n            cons = detail["Consistency"].get(code) if not detail.empty else np.nan\n            rows.append({\n                "基金名称": item.get("fund_name"),\n                "归因评分": scores.get(code),\n                "基金代码": code,\n                "基金类型": base.get("fund_type"),\n                "Alpha估算": round(float(alpha) * 100, 4) if pd.notna(alpha) else "--",\n                "Beta估算": round(float(beta), 3) if pd.notna(beta) else "--",\n                "收益一致性": f"{float(cons) * 100:.1f}%" if pd.notna(cons) else "--",\n                "基金经理": base.get("manager"),\n                "规模": base.get("assets_size"),\n            })\n        df = pd.DataFrame(rows).sort_values("归因评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"归因评分_{ts}.xlsx")\n        df.to_excel(out_path, index=False)\n        _beautify_with_explain(out_path, ATT_METRIC_EXPLAIN, score_col="归因评分",\n                               colored_ratio_cols=["Alpha估算"], header_fill_hex="44546A")\n        log(f"归因分析评分完成！共 {len(df)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        log(f"归因分析评分出错: {e}")\n        return None\n\n\n# ========================================================\n# ⑬ 综合评分（Composite - P1/P2/P3 + 可信度层）\n# ========================================================\n#\n# 设计原则：final = raw × confidence_factor × history_completeness\n# 关键目标：避免新基金/短样本霸榜，强化长期可信度。\n#\n# P1 核心质量（50%）:  收益 / 风险控制 / 回撤控制 / 长期表现 / 经理\n# P2 可信增强（35%）:  归因 / 效率 / 估值位置 / 成本\n# P3 战术择时（15%）:  趋势（受限，最高 80 且年轻基金再打折）\n# ========================================================\nCOMP_P1_MAP = {\n    "return":     0.20,   # 收益稳定性\n    "risk":       0.20,   # 风险控制（全样本夏普分位）\n    "drawdown":   0.20,   # 回撤控制（最大回撤分位）\n    "long_term":  0.20,   # 长期表现（3y/5y 分位）\n    "manager":    0.20,   # 经理\n}\nCOMP_P2_MAP = {\n    "attribution": 0.30,\n    "efficiency":  0.25,\n    "position":    0.25,\n    "cost":        0.20,\n}\nCOMP_P3_MAP = {"timing": 1.0}\nCOMP_LAYER = {"P1": 0.50, "P2": 0.35, "P3": 0.15}\n\n# 历史完整度：根据可用阶段数据长度衰减\nCOMP_HISTORY_TABLE = [\n    ("5y", 1.00),\n    ("3y", 0.85),\n    ("1y", 0.65),\n    ("6m", 0.45),\n    ("none", 0.25),\n]\n\n# 趋势模块在综合评分里的限幅\nCOMP_TIMING_CAP = 80.0          # 上限\nCOMP_TIMING_YOUNG_PENALTY = 0.7 # age_months < 12 时再打折\n\n# 风险样本长度基准：3 年交易日 ≈ 756\nCOMP_SAMPLE_BASELINE = 756\n\nCOMP_METRIC_EXPLAIN = {\n    "综合评分": ("最终综合分：\\n"\n                "  final = raw × reliability\\n"\n                "  reliability = max(0.45, 0.5×confidence + 0.5×history_completeness)\\n\\n"\n                "raw 公式（缺失层自动归一化）：\\n"\n                "  raw = P1 × 50% + P2 × 35% + P3 × 15%\\n\\n"\n                "  P1 = 收益×20% + 风险×20% + 回撤×20% + 长期×20% + 经理×20%\\n"\n                "  P2 = 归因×30% + 效率×25% + 位置×25% + 成本×20%\\n"\n                "  P3 = 趋势×100%（上限 80；年龄<12月再×0.7）\\n\\n"\n                "≥80 五星；≥70 四星；≥60 三星；≥50 两星；<50 一星。"),\n    "综合评级": "⭐⭐⭐⭐⭐ 五星：顶级\\n⭐⭐⭐⭐ 四星：优质\\n⭐⭐⭐ 三星：中等\\n⭐⭐ 两星：一般\\n⭐ 一星：暂不推荐",\n    "P1核心分":  ("第一优先级加权分 = 收益×20% + 风险×20% + 回撤×20% + 长期×20% + 经理×20%。\\n"\n                "缺失项不计入并重新归一化。"),\n    "可信度":    ("基金年龄可信度：\\n"\n                "  confidence = min(1.0, (age_months / 36) ** 0.5)\\n"\n                "3月≈0.29；6月≈0.41；12月≈0.58；24月≈0.82；36月=1.00。"),\n    "历史完整度":("根据长期区间数据是否可用给出的衰减系数：\\n"\n                "有5y=1.00；3y=0.85；1y=0.65；6m=0.45；<6m=0.25。"),\n    "可信度系数":("综合可信度（加权而非相乘，更温和）：\\n"\n                "  reliability = max(0.45, 0.5×confidence + 0.5×history_completeness)\\n"\n                "地板 45%，避免过度压制新基金。"),\n    "生命周期":  "<6月 新生；6月~2年 成长；2~5年 成熟；>5年 长周期。",\n    "原始分":    "未乘可信度系数前的 raw = P1×50% + P2×35% + P3×15%。",\n    "收益评分":  "模块二：70%×阶段收益分位加权 + 30%×正收益周期比率分位。",\n    "风险评分":  "模块三：全样本夏普百分位 × sample_factor (= sqrt(nav_days/756))。",\n    "回撤评分":  "模块三：最大回撤(绝对值)百分位(越小越好) × sample_factor。",\n    "长期评分":  "模块二延伸：3y×45% + 5y×55% 同类百分位，成立<36月不评分。",\n    "效率评分":  "模块四：sharpe(40)+calmar(35)+sortino(25) 百分位加权。",\n    "归因评分":  "模块五：Alpha(50)+Beta合理性(30)+一致性(20)。",\n    "位置评分":  "模块六：净值位置(50)+回撤位置(30)+成立以来位置(20)。",\n    "趋势评分":  "模块七：m1_rank(25)+m3_rank(20)+m6_rank(15)+方向(20)；综合评分中限幅≤80；age<12月再×0.7。",\n    "经理评分":  "模块八：任职年限×40% + 任期年化×50% + 稳定性加分×10%。",\n    "成本评分":  "模块九：费率×40% + 规模×40% + 申购状态×20%。",\n    "置顶推荐":  "按综合评分降序，Top 5 自动标记为 💎 核心关注。",\n}\n\n\n# ---------- 可信度相关工具 ----------\ndef _confidence_factor(age_months):\n    """基金可信度：S 型生长，36 个月封顶为 1.0"""\n    try:\n        am = float(age_months)\n    except Exception:\n        return 0.25\n    if pd.isna(am) or am <= 0:\n        return 0.25\n    return float(min(1.0, (am / 36.0) ** 0.5))\n\n\ndef _history_completeness(perf, age_months):\n    """历史完整度：依赖长期区间是否有值"""\n    if perf is None:\n        perf = {}\n\n    def _has(key, min_months):\n        if (age_months or 0) < min_months:\n            return False\n        v = parse_pct_to_float(perf.get(key))\n        return v is not None\n\n    if _has("5y", 60):\n        return 1.00\n    if _has("3y", 36):\n        return 0.85\n    if _has("1y", 12):\n        return 0.65\n    if _has("6m", 6):\n        return 0.45\n    return 0.25\n\n\ndef _lifecycle_label(age_months):\n    if age_months is None or pd.isna(age_months):\n        return "未知"\n    if age_months < 6:\n        return "新生 < 6m"\n    if age_months < 24:\n        return "成长 6m~2y"\n    if age_months < 60:\n        return "成熟 2y~5y"\n    return "长周期 > 5y"\n\n\ndef _sample_factor_from_nav(nav_history):\n    """基于 nav_history 长度的样本可信度：sqrt(nav_days / 756)"""\n    n = len(nav_history) if nav_history else 0\n    if n <= 0:\n        return 0.25\n    return float(min(1.0, (n / COMP_SAMPLE_BASELINE) ** 0.5))\n\n\ndef _age_months_of(item):\n    perf = item.get("performance", {}) or {}\n    base = item.get("base_info", {}) or {}\n    try:\n        am = (pd.to_datetime(perf.get("nav_date")) - pd.to_datetime(base.get("setup_date"))).days / 30.44\n        return float(am)\n    except Exception:\n        return np.nan\n\n\n# ---------- 聚合工具 ----------\ndef _agg_layer(scores, wmap):\n    vs, ws = 0.0, 0.0\n    for k, w in wmap.items():\n        v = scores.get(k)\n        if v is None:\n            continue\n        try:\n            fv = float(v)\n        except Exception:\n            continue\n        if pd.isna(fv):\n            continue\n        vs += fv * w\n        ws += w\n    return vs / ws if ws > 0 else None\n\n\ndef _composite_star(score):\n    if score is None or pd.isna(score) or score < 0:\n        return "⭐"\n    if score >= 80:\n        return "⭐⭐⭐⭐⭐"\n    if score >= 70:\n        return "⭐⭐⭐⭐"\n    if score >= 60:\n        return "⭐⭐⭐"\n    if score >= 50:\n        return "⭐⭐"\n    return "⭐"\n\n\n# 可信度地板（避免过度惩罚）：新基金最少保留 45% 的原始分\nCOMP_RELIABILITY_FLOOR = 0.45\n\n\ndef _reliability_factor(confidence, completeness):\n    """综合可信度：加权平均而非相乘，最低保留 45%。\n    confidence 反映年龄；completeness 反映历史完整度。"""\n    raw = 0.5 * float(confidence) + 0.5 * float(completeness)\n    return max(COMP_RELIABILITY_FLOOR, min(1.0, raw))\n\n\ndef _compose_one(sub, confidence, completeness):\n    """raw = P1*0.5 + P2*0.35 + P3*0.15（只加非缺失层）\n    final = raw × reliability_factor(confidence, completeness)\n    """\n    p1 = _agg_layer(sub, COMP_P1_MAP)\n    p2 = _agg_layer(sub, COMP_P2_MAP)\n    p3 = _agg_layer(sub, COMP_P3_MAP)\n    cs, cw = 0.0, 0.0\n    for val, w in [(p1, COMP_LAYER["P1"]), (p2, COMP_LAYER["P2"]), (p3, COMP_LAYER["P3"])]:\n        if val is not None:\n            cs += val * w\n            cw += w\n    raw = cs / cw if cw > 0 else None\n    if raw is None:\n        return None, None, None, None, "⭐"\n    reliability = _reliability_factor(confidence, completeness)\n    final = raw * reliability\n    final = round(max(0.0, min(100.0, final)), 2)\n    return (final,\n            round(raw, 2),\n            round(reliability, 3),\n            (round(p1, 2) if p1 is not None else None),\n            _composite_star(final))\n\n\n# --------- 各模块纯分值计算（基于已在本文件中的算法） ---------\ndef _positive_period_ratio(perf, age_months):\n    """一致性：正收益周期数 / 有效周期数。用于收益模块的稳定性加成。"""\n    if perf is None:\n        perf = {}\n    periods = [("1m", 1), ("3m", 3), ("6m", 6), ("1y", 12), ("3y", 36), ("5y", 60)]\n    total, pos = 0, 0\n    for key, min_age in periods:\n        if (age_months or 0) < min_age:\n            continue\n        v = parse_pct_to_float(perf.get(key))\n        if v is None:\n            continue\n        total += 1\n        if v > 0:\n            pos += 1\n    if total == 0:\n        return None\n    return pos / total\n\n\ndef _compute_return_scores(results):\n    """收益评分 = 0.7 × 阶段收益分位加权分 + 0.3 × 一致性分位 × 100\n    返回 Series[code -> score]"""\n    metric_rows = []\n    ages = {}\n    cons_map = {}\n    for item in results:\n        perf = item.get("performance", {}) or {}\n        base = item.get("base_info", {}) or {}\n        code = item.get("fund_code")\n        age = calc_age(base.get("setup_date"), perf.get("nav_date"))\n        ages[code] = age\n        cons_map[code] = _positive_period_ratio(perf, (age or 0) * 12)\n        metric_rows.append({\n            "fund_code": code,\n            "return_1m": parse_pct_to_float(perf.get("1m")),\n            "return_3m": parse_pct_to_float(perf.get("3m")),\n            "return_6m": parse_pct_to_float(perf.get("6m")),\n            "return_1y": parse_pct_to_float(perf.get("1y")),\n            "return_3y": parse_pct_to_float(perf.get("3y")),\n            "return_5y": parse_pct_to_float(perf.get("5y")),\n        })\n    metric_df = pd.DataFrame(metric_rows).set_index("fund_code")\n    age_series = pd.Series(ages)\n    raw_ret = calc_score(metric_df, age_series)  # 0~100\n\n    cons_series = pd.Series(cons_map, dtype=float)\n    if cons_series.notna().sum() > 1:\n        cons_rank = cons_series.rank(pct=True) * 100.0\n    else:\n        cons_rank = cons_series * 100.0  # 退化\n\n    # 混合：任一缺失则仅用另一个\n    codes = set(raw_ret.index) | set(cons_rank.index)\n    out = {}\n    for c in codes:\n        r = raw_ret.get(c)\n        k = cons_rank.get(c)\n        if r is None or (isinstance(r, float) and pd.isna(r)):\n            out[c] = None if (k is None or pd.isna(k)) else round(float(k), 2)\n        elif k is None or pd.isna(k):\n            out[c] = round(float(r), 2)\n        else:\n            out[c] = round(0.7 * float(r) + 0.3 * float(k), 2)\n    return pd.Series(out, dtype=float)\n\n\ndef _compute_long_term_scores(results):\n    """长期表现：3y / 5y 分位加权分，≥ 3 年才有分。\n    权重：3y=45, 5y=55，按照可用项重新归一化。"""\n    rows = []\n    ages = {}\n    for item in results:\n        perf = item.get("performance", {}) or {}\n        base = item.get("base_info", {}) or {}\n        code = item.get("fund_code")\n        age = calc_age(base.get("setup_date"), perf.get("nav_date"))\n        ages[code] = (age or 0) * 12\n        rows.append({\n            "fund_code": code,\n            "r_3y": parse_pct_to_float(perf.get("3y")),\n            "r_5y": parse_pct_to_float(perf.get("5y")),\n        })\n    df = pd.DataFrame(rows).set_index("fund_code")\n    if df.empty:\n        return pd.Series(dtype=float)\n\n    pct_df = pd.DataFrame(index=df.index)\n    for col in ["r_3y", "r_5y"]:\n        s = df[col].dropna()\n        if len(s) > 1:\n            pct_df[col] = _percentile_rank(_winsorize_series(df[col]))\n        else:\n            pct_df[col] = df[col]\n\n    weights = {"r_3y": 45, "r_5y": 55}\n    scores = {}\n    for code in df.index:\n        am = ages.get(code, 0) or 0\n        if am < 36:\n            scores[code] = None\n            continue\n        total_w, s = 0.0, 0.0\n        for m, w in weights.items():\n            v = pct_df.at[code, m] if code in pct_df.index else None\n            if v is None or (isinstance(v, float) and pd.isna(v)):\n                continue\n            total_w += w\n            s += v * w\n        if total_w == 0:\n            scores[code] = None\n        else:\n            scores[code] = round(s / total_w * 100, 2)\n    return pd.Series(scores, dtype=float)\n\n\ndef _compute_drawdown_scores(results, precomputed_mdd=None):\n    """回撤控制评分：优先用 precomputed_mdd（{code: |mdd|}），没有就现场算一次。"""\n    if precomputed_mdd is None:\n        precomputed_mdd = {}\n        for item in results:\n            code = item.get("fund_code")\n            hist = item.get("nav_history") or []\n            metrics, _ = _calc_risk_metrics_from_history(hist)\n            if not metrics:\n                continue\n            try:\n                mdd_pct = float(str(metrics.get("最大回撤", "")).replace("%", "")) / 100.0\n            except Exception:\n                continue\n            precomputed_mdd[code] = abs(mdd_pct)\n\n    if not precomputed_mdd:\n        return pd.Series(dtype=float)\n\n    # sample_factor\n    sf_map = {}\n    for item in results:\n        code = item.get("fund_code")\n        if code in precomputed_mdd:\n            sf_map[code] = _sample_factor_from_nav(item.get("nav_history") or [])\n\n    s = pd.Series(precomputed_mdd, dtype=float)\n    s_clip = s.clip(s.quantile(0.01), s.quantile(0.99))\n    rank_score = s_clip.rank(pct=True, ascending=False) * 100.0\n    sf = pd.Series(sf_map, dtype=float)\n    final = (rank_score * sf).round(2)\n    return final\n\n\ndef _compute_efficiency_scores(results):\n    metric_rows = []\n    for item in results:\n        perf = item.get("performance", {}) or {}\n        base = item.get("base_info", {}) or {}\n        age_m, age_y = _eff_age(base.get("setup_date"), perf.get("nav_date"))\n        metrics = _eff_metrics(perf, age_y)\n        metric_rows.append({"fund_code": item.get("fund_code"),\n                            **metrics, "age_months": age_m})\n    metric_df = pd.DataFrame(metric_rows).set_index("fund_code")\n    return _eff_score(metric_df[["sharpe", "calmar", "sortino"]],\n                      metric_df["age_months"])\n\n\ndef _compute_position_scores(results):\n    rows = []\n    for item in results:\n        p = item.get("performance", {}) or {}\n        b = item.get("base_info", {}) or {}\n        try:\n            age = (pd.to_datetime(b.get("setup_date")) - pd.to_datetime(p.get("nav_date"))).days\n            age = abs(age) / 30.44\n        except Exception:\n            age = 0\n        row = {\n            "code": item.get("fund_code"),\n            "type": b.get("fund_type", "未知"),\n            "age_months": age,\n            "nav_raw": _pos_safe_float(p.get("nav")),\n            "r_since": _pos_parse_pct(p.get("since")),\n        }\n        for f in POS_PERIOD_FIELDS:\n            row[f"r_{f}"] = _pos_parse_pct(p.get(f))\n        rows.append(row)\n    df = pd.DataFrame(rows)\n    if df.empty:\n        return pd.Series(dtype=float)\n    df[["nav_pos", "dd_proxy"]] = df.apply(_pos_calc_nav_stats, axis=1)\n    df["nav_score_rank"] = df.groupby("type")["nav_pos"].rank(pct=True, ascending=False)\n    df["since_score_rank"] = df.groupby("type")["r_since"].rank(pct=True, ascending=False)\n    df["dd_score_rank"] = df.groupby("type")["dd_proxy"].rank(pct=True, ascending=True)\n\n    def _final(r):\n        if (r["age_months"] or 0) < POS_MIN_AGE_MONTHS or pd.isna(r["nav_score_rank"]):\n            return np.nan\n        raw = (r["nav_score_rank"] * POS_WEIGHTS["nav_pos_score"]\n               + r["dd_score_rank"] * POS_WEIGHTS["dd_score"]\n               + r["since_score_rank"] * POS_WEIGHTS["since_pos_score"]) * 100\n        if r["age_months"] < 3:\n            raw *= 0.8\n        elif r["age_months"] < 6:\n            raw *= 0.9\n        return round(raw, 2)\n\n    df["score"] = df.apply(_final, axis=1)\n    return df.set_index("code")["score"]\n\n\ndef _compute_timing_scores_detail(results):\n    rows = []\n    for item in results:\n        p = item.get("performance", {}) or {}\n        b = item.get("base_info", {}) or {}\n        try:\n            age = (pd.to_datetime(p.get("nav_date")) - pd.to_datetime(b.get("setup_date"))).days / 30.44\n        except Exception:\n            age = 0\n        rows.append({\n            "code": item.get("fund_code"),\n            "type": b.get("fund_type", "未知"),\n            "age": age,\n            "m1": _timing_parse_pct(p.get("1m")),\n            "m3": _timing_parse_pct(p.get("3m")),\n            "m6": _timing_parse_pct(p.get("6m")),\n        })\n    df = pd.DataFrame(rows)\n    if df.empty:\n        return pd.Series(dtype=float), pd.DataFrame()\n    df[["dir_score", "dir_label"]] = df.apply(_timing_trend_details, axis=1)\n    df[["is_oh", "is_rec"]] = df.apply(_timing_signals, axis=1)\n    for col in ["m1", "m3", "m6", "dir_score"]:\n        df[f"{col}_rank"] = df.groupby("type")[col].rank(pct=True)\n\n    def _final(r):\n        if (r["age"] or 0) < TIMING_MIN_AGE_MONTHS:\n            return np.nan\n        w = dict(TIMING_WEIGHTS)\n        if r["age"] < 3:\n            for k in ["m3_rank", "m6_rank", "dir_score_rank"]:\n                w.pop(k, None)\n        elif r["age"] < 6:\n            w.pop("m6_rank", None)\n        try:\n            score_sum = sum((r[k] if not pd.isna(r[k]) else 0) * ww for k, ww in w.items())\n            weight_sum = sum(w.values())\n            base = (score_sum / weight_sum) * 100 if weight_sum else 0\n        except KeyError:\n            return 0.0\n        if r["is_oh"]:\n            base -= TIMING_OVERHEAT_PENALTY\n        if r["is_rec"]:\n            base += TIMING_RECOVERY_BONUS\n        if r["age"] < 6:\n            base *= 0.9\n        return round(max(0, min(100, base)), 2)\n\n    df["score"] = df.apply(_final, axis=1)\n    detail = df.set_index("code")[["dir_label", "is_oh", "is_rec"]]\n    return df.set_index("code")["score"], detail\n\n\ndef _compute_manager_scores_detail(results):\n    rows = []\n    for item in results:\n        p = item.get("performance", {}) or {}\n        b = item.get("base_info", {}) or {}\n        try:\n            tenure_y = (pd.to_datetime(p.get("nav_date")) - pd.to_datetime(b.get("setup_date"))).days / 365.25\n        except Exception:\n            tenure_y = 0\n        r_total = _mgr_parse_pct(p.get("since"))\n        ann_ret = _mgr_ann_return(r_total, tenure_y)\n        rows.append({\n            "code": item.get("fund_code"),\n            "type": b.get("fund_type", "未知"),\n            "tenure_y": tenure_y,\n            "ann_ret": ann_ret,\n        })\n    df = pd.DataFrame(rows)\n    if df.empty:\n        return pd.Series(dtype=float), pd.DataFrame()\n    df["y_rank"] = df.groupby("type")["tenure_y"].rank(pct=True)\n    df["ret_rank"] = df.groupby("type")["ann_ret"].rank(pct=True)\n\n    def _final(r):\n        age_m = (r["tenure_y"] or 0) * 12\n        if age_m < MGR_MIN_AGE_MONTHS:\n            return np.nan\n        w = MGR_WEIGHTS_YOUNG if age_m < 12 else MGR_WEIGHTS_NORMAL\n        y_rank = r["y_rank"] if not pd.isna(r["y_rank"]) else 0\n        ret_rank = r["ret_rank"] if not pd.isna(r["ret_rank"]) else 0\n        base = (y_rank * w["tenure_years"] + ret_rank * w["tenure_ann_return"]) * 100\n        total = base + _mgr_bonus(r["tenure_y"])\n        if age_m < 6:\n            total *= 0.85\n        elif age_m < 12:\n            total *= 0.92\n        return round(max(0, min(100, total)), 2)\n\n    df["score"] = df.apply(_final, axis=1)\n\n    def _label(y):\n        if pd.isna(y):\n            return "--"\n        label = f"{y:.1f}年"\n        if y >= 5:\n            return f"⭐ {label}(资深)"\n        if y >= 3:\n            return f"{label}(熟练)"\n        return f"{label}(新锐)"\n\n    df["任职年限估算"] = df["tenure_y"].apply(_label)\n    df["任期年化回报"] = df["ann_ret"].apply(lambda x: f"{x * 100:.2f}%" if not pd.isna(x) else "--")\n    detail = df.set_index("code")[["任职年限估算", "任期年化回报"]]\n    return df.set_index("code")["score"], detail\n\n\ndef _compute_cost_scores_detail(results):\n    rows = []\n    for item in results:\n        b = item.get("base_info", {}) or {}\n        s = item.get("status", {}) or {}\n        fee_raw = _cost_parse_fee(s.get("buy_fee"))\n        size_raw = _cost_parse_size(b.get("assets_size"))\n        status_raw = s.get("buy_status")\n        fs = _cost_score_fee(fee_raw)\n        ss, s_label = _cost_score_size(size_raw)\n        as_ = _cost_score_access(status_raw)\n        total_score = (fs * COST_WEIGHTS["fee"]\n                       + ss * COST_WEIGHTS["size"]\n                       + as_ * COST_WEIGHTS["access"])\n        rows.append({\n            "code": item.get("fund_code"),\n            "score": round(total_score, 2),\n            "申购费率": f"{fee_raw:.3f}%" if not pd.isna(fee_raw) else "--",\n            "规模(亿)": round(size_raw, 2),\n            "规模评级": s_label,\n            "买入状态": status_raw or "--",\n        })\n    df = pd.DataFrame(rows)\n    if df.empty:\n        return pd.Series(dtype=float), pd.DataFrame()\n    df = df.set_index("code")\n    return df["score"], df[["申购费率", "规模(亿)", "规模评级", "买入状态"]]\n\n\ndef _compute_risk_scores(results):\n    """基于 nav_history 计算每只基金的夏普比率与最大回撤，夏普做全样本百分位排名 × sample_factor 得 0~100 风险评分。\n    返回 (score_series, metrics_df, mdd_map)"""\n    sharpe_map, sf_map, mdd_map = {}, {}, {}\n    metrics_rows = []\n    for item in results:\n        code = item.get("fund_code")\n        hist = item.get("nav_history") or []\n        metrics, _ = _calc_risk_metrics_from_history(hist)\n        if metrics:\n            sharpe_map[code] = metrics.get("夏普比率")\n            sf_map[code] = _sample_factor_from_nav(hist)\n            try:\n                mdd_pct = float(str(metrics.get("最大回撤", "")).replace("%", "")) / 100.0\n                mdd_map[code] = abs(mdd_pct)\n            except Exception:\n                pass\n            row = {"code": code, **{k: v for k, v in metrics.items() if not k.startswith("_")}}\n            metrics_rows.append(row)\n    if not sharpe_map:\n        return pd.Series(dtype=float), pd.DataFrame(), {}\n\n    s = pd.Series(sharpe_map, dtype=float)\n    s_clip = s.clip(s.quantile(0.01), s.quantile(0.99))\n    rank_score = s_clip.rank(pct=True) * 100\n    sf = pd.Series(sf_map, dtype=float)\n    final = (rank_score * sf).round(2)\n    metrics_df = pd.DataFrame(metrics_rows).set_index("code")\n    return final, metrics_df, mdd_map\n\n\ndef _beautify_composite_excel(path, sheet_scores):\n    """为综合评分多 Sheet 文件做美化：冻结/标题/评分格/星级/列宽/批注"""\n    wb = load_workbook(path)\n\n    # 标题样式\n    header_fill = PatternFill("solid", fgColor="203764")\n    header_font = Font(bold=True, color="FFFFFF")\n\n    for sheet_name, score_cols in sheet_scores.items():\n        if sheet_name not in wb.sheetnames:\n            continue\n        ws = wb[sheet_name]\n        ws.freeze_panes = "A2"\n        headers = [str(c.value) for c in ws[1]]\n\n        for cell in ws[1]:\n            cell.fill = header_fill\n            cell.font = header_font\n            cell.alignment = Alignment(horizontal="center", vertical="center")\n            col_txt = str(cell.value)\n            if col_txt in COMP_METRIC_EXPLAIN:\n                text = COMP_METRIC_EXPLAIN[col_txt]\n                cmt = Comment(text, "FundSystem")\n                cmt.width = 320\n                cmt.height = 50 + text.count("\\n") * 22\n                cell.comment = cmt\n\n        for row in ws.iter_rows(min_row=2):\n            for cell in row:\n                col_name = headers[cell.column - 1]\n                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")\n                # 评分列条件格式\n                if col_name in score_cols and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(cell.value)\n                    except Exception:\n                        continue\n                    if col_name == "综合评分":\n                        if v >= 80:\n                            color = "00B050"\n                        elif v >= 70:\n                            color = "92D050"\n                        elif v >= 60:\n                            color = "00B0F0"\n                        elif v >= 50:\n                            color = "FFFF00"\n                        else:\n                            color = "FF4B4B"\n                        cell.fill = PatternFill("solid", fgColor=color)\n                        cell.font = Font(bold=True, size=12)\n                    else:\n                        color = _score_fill_color(v)\n                        if color:\n                            cell.fill = PatternFill("solid", fgColor=color)\n                        cell.font = Font(bold=True)\n                # 星级列\n                if col_name == "综合评级":\n                    cell.alignment = Alignment(horizontal="center", vertical="center")\n                    cell.font = Font(color="FFC000", bold=True)\n                # 置顶推荐列\n                if col_name == "置顶推荐" and "核心关注" in str(cell.value or ""):\n                    cell.fill = PatternFill("solid", fgColor="FFE699")\n                    cell.font = Font(bold=True, color="C00000")\n                    cell.alignment = Alignment(horizontal="center", vertical="center")\n                # 百分比列红绿\n                if col_name in ("近1月", "近3月", "近6月", "近1年", "近3年", "近5年",\n                                "年化收益", "最大回撤", "当前回撤", "Alpha估算",\n                                "任期年化回报", "成立以来收益") and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(str(cell.value).replace("%", ""))\n                        cell.font = Font(color="FF0000" if v > 0 else "00B050")\n                    except Exception:\n                        pass\n\n        # 列宽自适应\n        for col in ws.columns:\n            max_len = 0\n            col_letter = get_column_letter(col[0].column)\n            for cell in col:\n                if cell.value:\n                    try:\n                        curr = len(str(cell.value).encode("gbk"))\n                    except Exception:\n                        curr = len(str(cell.value))\n                    if curr > max_len:\n                        max_len = curr\n            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)\n\n    wb.save(path)\n\n\ndef run_composite_score(log):\n    """综合评分：一次性计算 8 维评分 + P1/P2/P3 聚合 + 多-Sheet Excel"""\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，启动全维度评分引擎...")\n        log("正在计算各维度评分（收益/效率/位置/趋势/经理/成本/归因/风险）...")\n\n        return_s = _compute_return_scores(results)\n        log(f"  ✅ 收益评分完成 ({return_s.notna().sum()}/{len(results)})")\n        long_s = _compute_long_term_scores(results)\n        log(f"  ✅ 长期表现评分完成 ({long_s.notna().sum()}/{len(results)})")\n        eff_s = _compute_efficiency_scores(results)\n        log(f"  ✅ 效率评分完成 ({eff_s.notna().sum()}/{len(results)})")\n        pos_s = _compute_position_scores(results)\n        log(f"  ✅ 位置评分完成 ({pos_s.notna().sum()}/{len(results)})")\n        tim_s, tim_detail = _compute_timing_scores_detail(results)\n        log(f"  ✅ 趋势评分完成 ({tim_s.notna().sum()}/{len(results)})")\n        mgr_s, mgr_detail = _compute_manager_scores_detail(results)\n        log(f"  ✅ 经理评分完成 ({mgr_s.notna().sum()}/{len(results)})")\n        cost_s, cost_detail = _compute_cost_scores_detail(results)\n        log(f"  ✅ 成本评分完成 ({cost_s.notna().sum()}/{len(results)})")\n        att_s, att_detail = _compute_attribution(results)\n        log(f"  ✅ 归因评分完成 ({att_s.notna().sum()}/{len(results)})")\n        risk_s, risk_metrics_df, mdd_map = _compute_risk_scores(results)\n        dd_s = _compute_drawdown_scores(results, precomputed_mdd=mdd_map)\n        if not risk_s.empty:\n            log(f"  ✅ 风险评分完成 (×sample_factor: {len(risk_s)}/{len(results)})")\n            log(f"  ✅ 回撤评分完成 ({dd_s.notna().sum()}/{len(results)})")\n        else:\n            log("  ⚠️ 风险/回撤评分：未检测到 nav_history，将按缺失处理。")\n\n        # 聚合\n        summary_rows = []\n        for item in results:\n            code = item.get("fund_code")\n            perf = item.get("performance", {}) or {}\n            base = item.get("base_info", {}) or {}\n            try:\n                age_y = (pd.to_datetime(perf.get("nav_date")) - pd.to_datetime(base.get("setup_date"))).days / 365.25\n            except Exception:\n                age_y = None\n            age_m = (age_y * 12) if age_y is not None else None\n\n            # 可信度 / 历史完整度\n            cf = _confidence_factor(age_m) if age_m is not None else 0.25\n            hc = _history_completeness(perf, age_m)\n            lifecycle = _lifecycle_label(age_m)\n\n            # 趋势限幅 & 年轻打折\n            tim_raw = tim_s.get(code)\n            tim_use = None\n            if tim_raw is not None and not pd.isna(tim_raw):\n                tim_use = min(float(tim_raw), COMP_TIMING_CAP)\n                if (age_m or 0) < 12:\n                    tim_use *= COMP_TIMING_YOUNG_PENALTY\n                tim_use = round(tim_use, 2)\n\n            sub = {\n                "return": return_s.get(code),\n                "risk": risk_s.get(code) if not risk_s.empty else None,\n                "drawdown": dd_s.get(code) if not dd_s.empty else None,\n                "long_term": long_s.get(code) if not long_s.empty else None,\n                "manager": mgr_s.get(code),\n                "efficiency": eff_s.get(code),\n                "attribution": att_s.get(code),\n                "position": pos_s.get(code),\n                "cost": cost_s.get(code),\n                "timing": tim_use,\n            }\n            composite, raw, reliability, p1, star = _compose_one(sub, cf, hc)\n\n            summary_rows.append({\n                "置顶推荐": "",\n                "基金代码": code,\n                "基金名称": item.get("fund_name"),\n                "综合评分": composite,\n                "综合评级": star,\n                "原始分": raw,\n                "可信度": round(cf, 2),\n                "历史完整度": round(hc, 2),\n                "可信度系数": reliability,\n                "生命周期": lifecycle,\n                "P1核心分": p1,\n                "收益评分": sub["return"],\n                "风险评分": sub["risk"],\n                "回撤评分": sub["drawdown"],\n                "长期评分": sub["long_term"],\n                "效率评分": sub["efficiency"],\n                "归因评分": sub["attribution"],\n                "位置评分": sub["position"],\n                "趋势评分": sub["timing"],\n                "经理评分": sub["manager"],\n                "成本评分": sub["cost"],\n                "近1月": perf.get("1m", "--"),\n                "近3月": perf.get("3m", "--"),\n                "近6月": perf.get("6m", "--"),\n                "近1年": perf.get("1y", "--"),\n                "近3年": perf.get("3y", "--"),\n                "近5年": perf.get("5y", "--"),\n                "最新净值": perf.get("nav", "--"),\n                "净值日期": perf.get("nav_date", "--"),\n                "基金类型": base.get("fund_type", "--"),\n                "规模": base.get("assets_size", "--"),\n                "基金经理": base.get("manager", "--"),\n                "成立日期": base.get("setup_date", "--"),\n                "成立年限": round(age_y, 2) if age_y is not None else "--",\n            })\n\n        df_summary = pd.DataFrame(summary_rows).sort_values(\n            "综合评分", ascending=False, na_position="last").reset_index(drop=True)\n        # Top 5 置顶标记\n        top_n = min(5, len(df_summary))\n        if top_n > 0:\n            df_summary.loc[:top_n - 1, "置顶推荐"] = "💎 核心关注"\n\n        # 各维度详情 Sheet\n        detail_frames = {}\n\n        # 收益详情\n        detail_frames["💹 收益详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "收益评分": return_s.get(it.get("fund_code")),\n            "近1月": (it.get("performance") or {}).get("1m", "--"),\n            "近3月": (it.get("performance") or {}).get("3m", "--"),\n            "近6月": (it.get("performance") or {}).get("6m", "--"),\n            "近1年": (it.get("performance") or {}).get("1y", "--"),\n            "近3年": (it.get("performance") or {}).get("3y", "--"),\n            "近5年": (it.get("performance") or {}).get("5y", "--"),\n            "成立以来": (it.get("performance") or {}).get("since", "--"),\n        } for it in results]).sort_values("收益评分", ascending=False, na_position="last")\n\n        # 效率详情\n        detail_frames["🔬 效率详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "效率评分": eff_s.get(it.get("fund_code")),\n            "Sharpe估算": risk_metrics_df.loc[it.get("fund_code"), "夏普比率"] if (\n                not risk_metrics_df.empty and it.get("fund_code") in risk_metrics_df.index) else "--",\n            "年化收益": risk_metrics_df.loc[it.get("fund_code"), "年化收益"] if (\n                not risk_metrics_df.empty and it.get("fund_code") in risk_metrics_df.index) else "--",\n            "波动率": risk_metrics_df.loc[it.get("fund_code"), "波动率"] if (\n                not risk_metrics_df.empty and it.get("fund_code") in risk_metrics_df.index) else "--",\n        } for it in results]).sort_values("效率评分", ascending=False, na_position="last")\n\n        # 归因详情\n        detail_frames["🎯 归因详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "归因评分": att_s.get(it.get("fund_code")),\n            "Alpha估算": round(float(att_detail["Alpha"].get(it.get("fund_code"))) * 100, 4)\n                if (not att_detail.empty and pd.notna(att_detail["Alpha"].get(it.get("fund_code")))) else "--",\n            "Beta估算": round(float(att_detail["Beta"].get(it.get("fund_code"))), 3)\n                if (not att_detail.empty and pd.notna(att_detail["Beta"].get(it.get("fund_code")))) else "--",\n            "收益一致性": f"{float(att_detail[\'Consistency\'].get(it.get(\'fund_code\'))) * 100:.1f}%"\n                if (not att_detail.empty and pd.notna(att_detail["Consistency"].get(it.get("fund_code")))) else "--",\n        } for it in results]).sort_values("归因评分", ascending=False, na_position="last")\n\n        # 位置详情\n        detail_frames["📍 位置详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "位置评分": pos_s.get(it.get("fund_code")),\n            "成立以来收益": (it.get("performance") or {}).get("since", "--"),\n            "最新净值": (it.get("performance") or {}).get("nav", "--"),\n        } for it in results]).sort_values("位置评分", ascending=False, na_position="last")\n\n        # 趋势详情\n        detail_frames["📈 趋势详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "趋势评分": tim_s.get(it.get("fund_code")),\n            "近1月": (it.get("performance") or {}).get("1m", "--"),\n            "近3月": (it.get("performance") or {}).get("3m", "--"),\n            "近6月": (it.get("performance") or {}).get("6m", "--"),\n            "趋势方向": tim_detail["dir_label"].get(it.get("fund_code"))\n                if not tim_detail.empty else "--",\n            "过热信号": ("⚠️ 过热" if tim_detail["is_oh"].get(it.get("fund_code")) else "正常")\n                if not tim_detail.empty else "--",\n            "修复信号": ("✅ 修复" if tim_detail["is_rec"].get(it.get("fund_code")) else "--")\n                if not tim_detail.empty else "--",\n        } for it in results]).sort_values("趋势评分", ascending=False, na_position="last")\n\n        # 经理详情\n        detail_frames["👤 经理详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "经理评分": mgr_s.get(it.get("fund_code")),\n            "基金经理": (it.get("base_info") or {}).get("manager", "--"),\n            "任职年限估算": mgr_detail["任职年限估算"].get(it.get("fund_code"))\n                if not mgr_detail.empty else "--",\n            "任期年化回报": mgr_detail["任期年化回报"].get(it.get("fund_code"))\n                if not mgr_detail.empty else "--",\n            "成立日期": (it.get("base_info") or {}).get("setup_date", "--"),\n        } for it in results]).sort_values("经理评分", ascending=False, na_position="last")\n\n        # 成本详情\n        detail_frames["💰 成本详情"] = pd.DataFrame([{\n            "基金代码": it.get("fund_code"),\n            "基金名称": it.get("fund_name"),\n            "成本评分": cost_s.get(it.get("fund_code")),\n            "申购费率": cost_detail["申购费率"].get(it.get("fund_code"))\n                if not cost_detail.empty else "--",\n            "规模(亿)": cost_detail["规模(亿)"].get(it.get("fund_code"))\n                if not cost_detail.empty else "--",\n            "规模评级": cost_detail["规模评级"].get(it.get("fund_code"))\n                if not cost_detail.empty else "--",\n            "买入状态": cost_detail["买入状态"].get(it.get("fund_code"))\n                if not cost_detail.empty else "--",\n        } for it in results]).sort_values("成本评分", ascending=False, na_position="last")\n\n        # 风险看板（只展示有历史数据的）\n        if not risk_metrics_df.empty:\n            rm = risk_metrics_df.reset_index().rename(columns={"code": "基金代码"})\n            name_map = {it.get("fund_code"): it.get("fund_name") for it in results}\n            rm["基金名称"] = rm["基金代码"].map(name_map)\n            cols = ["基金代码", "基金名称", "年化收益", "最大回撤", "当前回撤",\n                    "夏普比率", "卡玛比率", "索提诺比率", "波动率", "下行波动",\n                    "溃疡指数", "回撤状态", "回撤进度", "决策建议"]\n            detail_frames["⚡ 风险看板"] = rm[[c for c in cols if c in rm.columns]]\n\n        # 写 Excel\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"长期综合评分_{ts}.xlsx")\n\n        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:\n            df_summary.to_excel(writer, sheet_name="📊 长期综合总表", index=False)\n            for sheet_name, df_d in detail_frames.items():\n                df_d.to_excel(writer, sheet_name=sheet_name, index=False)\n\n        # 美化\n        sheet_scores = {\n            "📊 长期综合总表": {"综合评分", "原始分", "P1核心分",\n                            "收益评分", "风险评分", "回撤评分", "长期评分",\n                            "效率评分", "归因评分", "位置评分", "趋势评分",\n                            "经理评分", "成本评分"},\n            "💹 收益详情": {"收益评分"},\n            "🔬 效率详情": {"效率评分"},\n            "🎯 归因详情": {"归因评分"},\n            "📍 位置详情": {"位置评分"},\n            "📈 趋势详情": {"趋势评分"},\n            "👤 经理详情": {"经理评分"},\n            "💰 成本详情": {"成本评分"},\n            "⚡ 风险看板": set(),\n        }\n        _beautify_composite_excel(out_path, sheet_scores)\n\n        top = df_summary.head(3)\n        for _, row in top.iterrows():\n            log(f"  🏆 {row.get(\'基金名称\', \'\')} 综合评分 {row.get(\'综合评分\', \'--\')} {row.get(\'综合评级\', \'\')}")\n        log(f"综合评分完成！共 {len(df_summary)} 只基金")\n        log(f"已导出美化Excel: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"综合评分出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ========================================================\n# ⑭ 回撤震荡筛选（筛选法二）\n# ========================================================\n#\n# 目标：找"已经经历明显回撤、但近期下跌趋缓、波动收敛、进入震荡盘整"的基金。\n# 不是找"跌得最多"的，而是找"跌过之后开始稳住"的。\n# 完全独立于长期综合评分，只生成"观察池"。\n# ========================================================\n\n# 前置过滤\nDRAWDOWN_MIN_MDD = 0.15\nDRAWDOWN_MIN_CURRENT_DD = 0.08\nDRAWDOWN_MIN_PROGRESS = 0.40\nDRAWDOWN_MAX_R1_DROP = -0.08\nDRAWDOWN_MAX_R1_RISE = 0.10\nDRAWDOWN_MAX_R3_DROP = -0.15\nDRAWDOWN_MAX_VOL_CONTRACTION = 1.10\nDRAWDOWN_BOTTOM_VOL_CONTRACTION = 0.80\nDRAWDOWN_SIDEWAYS_R1_ABS = 0.05\nDRAWDOWN_MIN_AGE_MONTHS = 6\nDRAWDOWN_MIN_NAV_POINTS = 60\n\n# 排雷阈值\nDRAWDOWN_MIN_R1Y = -0.30          # 近1年不能过度下跌\nDRAWDOWN_MIN_SIZE_YI = 0.5        # 规模不能低于 0.5 亿\nDRAWDOWN_MAX_CONSEC_DROP_DAYS = 30  # 最近 60 日内连续下跌天数阈值\n\n# 剔除的基金类型关键字（货币 / 短债 / 理财）\n_DD_EXCLUDE_TYPES = ("货币", "短债", "理财")\n\nDRAWDOWN_METRIC_EXPLAIN = {\n    "回撤震荡总分": ("总分公式：\\n"\n                    "  回撤深度分 × 30% + 波动收敛分 × 25% + 企稳分 × 25% + 位置分 × 20%\\n\\n"\n                    "硬筛（全部需满足）：\\n"\n                    "  MDD ≥ 15% 且 当前回撤 ≥ 8%\\n"\n                    "  当前回撤 ≥ MDD × 40%（仍在回撤区间）\\n"\n                    "  近1月 ∈ [-8%, 10%]，近3月 ≥ -15%\\n"\n                    "  近1年 ≥ -30%（排雷：避免基本面恶化型）\\n"\n                    "  vol20/vol60 ≤ 1.10（波动不放大）\\n"\n                    "  规模 ≥ 0.5 亿（排雷：剔除清盘风险）\\n"\n                    "  申购未暂停（排雷：确保可交易）\\n"\n                    "  近60日净值未连续创新低（排雷：未形成持续下跌）\\n"\n                    "  成立 ≥ 6月，nav_history ≥ 60 个点"),\n    "观察等级": ("A-重点观察：总分≥75 且 \'底部震荡\'\\n"\n                "B-观察：总分≥60\\n"\n                "C-谨慎观察：总分≥45\\n"\n                "D-剔除：总分<45"),\n    "状态标签": ("底部震荡：MDD≥15% 且 当前回撤≥8% 且 40%≤进度≤100% 且 -5%≤r1≤5% 且 vol20/vol60≤0.80\\n"\n                "深度回撤-仍在下跌：当前回撤≥15% 且 r1<-8%\\n"\n                "回撤修复中：当前回撤≥8% 且 r1>5% 且 r3>0\\n"\n                "强反弹-等待回踩：当前回撤≥8% 且 r1>10%\\n"\n                "非回撤震荡：其它"),\n    "历史最大回撤": "mdd = min(NAV / cummax(NAV) - 1)。",\n    "当前回撤": "current_drawdown = NAV_last / cummax(NAV) - 1。0% = 历史新高。",\n    "回撤进度": ("当前回撤 / 历史最大回撤 × 100%。\\n"\n                "<30% 已明显修复；30%~70% 中度；70%~100% 深度；>100% 创新低。"),\n    "波动收敛率": ("vol20 / vol60。\\n"\n                  "<0.6 明显收敛；0.6~0.8 温和收敛；0.8~1.1 普通震荡；>1.1 波动放大。"),\n    "回撤深度分": "min(|当前回撤| / 30%, 1) × 100。当前回撤≥30% 满分。",\n    "波动收敛分": "≤0.6→100；≤0.8→80；≤1.0→60；≤1.2→40；其它→20。",\n    "企稳分": ("满分 100：\\n"\n              "  近1月 > -3% → +40\\n"\n              "  近3月 > -8% → +30\\n"\n              "  |近1月| ≤ 5% → +30"),\n    "位置分": "优先用模块六位置评分；若缺则用近1年净值区间反向位置 (1-pos)×100。",\n    "排雷标签": ("基于硬筛通过但仍需警惕的软提示：\\n"\n                "  规模偏小 (< 1 亿)\\n"\n                "  长期弱势 (近1年 < -20%)\\n"\n                "  反弹无量 (近20日累计 > 5% 但波动率异常收敛)\\n"\n                "  近期连跌多日 (连续下跌 ≥ 15 日)\\n"\n                "  限额申购"),\n    "申购状态": "开放申购 / 限大额 / 暂停申购 等。硬筛已剔除\'暂停\'的基金。",\n    "vol20": "近20日年化波动率 = std(ret[-20:]) × √252。",\n    "vol60": "近60日年化波动率 = std(ret[-60:]) × √252。",\n}\n\n\ndef _dd_parse_pct_dec(val):\n    """\'12.5%\' → 0.125；\'-\'或异常 → np.nan"""\n    if val is None:\n        return np.nan\n    if isinstance(val, (int, float)):\n        return float(val) / 100.0\n    s = str(val).strip().replace("%", "")\n    if s in ("", "--"):\n        return np.nan\n    try:\n        return float(s) / 100.0\n    except Exception:\n        return np.nan\n\n\ndef _dd_age_months(setup, nav_date):\n    try:\n        return (pd.to_datetime(nav_date) - pd.to_datetime(setup)).days / 30.44\n    except Exception:\n        return 0\n\n\ndef _dd_status_label(mdd_abs, curr_dd_abs, progress, r1, r3, vol_contraction):\n    """五种状态标签"""\n    if (mdd_abs >= DRAWDOWN_MIN_MDD\n            and curr_dd_abs >= DRAWDOWN_MIN_CURRENT_DD\n            and DRAWDOWN_MIN_PROGRESS <= progress <= 1.00\n            and -DRAWDOWN_SIDEWAYS_R1_ABS <= r1 <= DRAWDOWN_SIDEWAYS_R1_ABS\n            and vol_contraction <= DRAWDOWN_BOTTOM_VOL_CONTRACTION):\n        return "底部震荡"\n    if curr_dd_abs >= 0.15 and r1 < DRAWDOWN_MAX_R1_DROP:\n        return "深度回撤-仍在下跌"\n    if curr_dd_abs >= 0.08 and r1 > DRAWDOWN_SIDEWAYS_R1_ABS and r3 > 0:\n        return "回撤修复中"\n    if curr_dd_abs >= 0.08 and r1 > DRAWDOWN_MAX_R1_RISE:\n        return "强反弹-等待回踩"\n    return "非回撤震荡"\n\n\ndef _dd_nav_position_score(nav_series):\n    """近一年净值区间内的反向位置分 (1 - pos) × 100"""\n    tail = nav_series.tail(252)\n    lo, hi = tail.min(), tail.max()\n    if hi - lo < 1e-9:\n        return 50.0\n    pos = (nav_series.iloc[-1] - lo) / (hi - lo)\n    return float(max(0.0, min(100.0, (1 - pos) * 100)))\n\n\ndef _dd_calc_one(item, position_s):\n    """单只基金的回撤震荡计算。不满足硬筛条件则返回 None。"""\n    perf = item.get("performance", {}) or {}\n    base = item.get("base_info", {}) or {}\n    code = item.get("fund_code")\n    ftype = base.get("fund_type", "") or ""\n\n    # 剔除货币/短债/理财\n    for kw in _DD_EXCLUDE_TYPES:\n        if kw in ftype:\n            return None\n\n    # 年龄门槛\n    age_m = _dd_age_months(base.get("setup_date"), perf.get("nav_date"))\n    if age_m < DRAWDOWN_MIN_AGE_MONTHS:\n        return None\n\n    hist = item.get("nav_history") or []\n    if len(hist) < DRAWDOWN_MIN_NAV_POINTS:\n        return None\n\n    # 构造净值序列\n    try:\n        df = pd.DataFrame(hist)\n        df["date"] = pd.to_datetime(df["date"], errors="coerce")\n        df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)\n        nav = df["val"].astype(float).ffill().dropna()\n    except Exception:\n        return None\n    if len(nav) < DRAWDOWN_MIN_NAV_POINTS:\n        return None\n\n    returns = nav.pct_change().dropna()\n    if len(returns) < 20:\n        return None\n\n    # 回撤\n    cummax = nav.cummax()\n    drawdown = nav / cummax - 1.0\n    mdd = float(drawdown.min())\n    curr_dd = float(drawdown.iloc[-1])\n    mdd_abs = abs(mdd)\n    curr_dd_abs = abs(curr_dd)\n    progress = (curr_dd_abs / mdd_abs) if mdd_abs > 1e-9 else 0.0\n\n    # 近1月 / 近3月（直接用 profile，避免历史长度不够）\n    r1 = _dd_parse_pct_dec(perf.get("1m"))\n    r3 = _dd_parse_pct_dec(perf.get("3m"))\n    if np.isnan(r1):\n        r1 = 0.0\n    if np.isnan(r3):\n        r3 = 0.0\n\n    # 波动收敛\n    import math as _m\n    vol20 = float(returns.tail(20).std() * _m.sqrt(252)) if len(returns) >= 20 else np.nan\n    vol60 = float(returns.tail(60).std() * _m.sqrt(252)) if len(returns) >= 60 else np.nan\n    if not vol60 or np.isnan(vol60) or vol60 <= 0:\n        vol_contraction = 999.0\n    else:\n        vol_contraction = vol20 / vol60 if not np.isnan(vol20) else 999.0\n\n    # 硬筛\n    hard_ok = (\n        mdd_abs >= DRAWDOWN_MIN_MDD\n        and curr_dd_abs >= DRAWDOWN_MIN_CURRENT_DD\n        and curr_dd_abs >= mdd_abs * DRAWDOWN_MIN_PROGRESS\n        and r1 >= DRAWDOWN_MAX_R1_DROP\n        and r1 <= DRAWDOWN_MAX_R1_RISE\n        and r3 >= DRAWDOWN_MAX_R3_DROP\n        and vol_contraction <= DRAWDOWN_MAX_VOL_CONTRACTION\n    )\n    if not hard_ok:\n        return None\n\n    # ----- 排雷硬筛 -----\n    # 1) 近1年不能过度下跌\n    r1y = _dd_parse_pct_dec(perf.get("1y"))\n    if not np.isnan(r1y) and r1y < DRAWDOWN_MIN_R1Y:\n        return None\n\n    # 2) 规模 ≥ 0.5 亿（债券类专题除外；回撤震荡默认剔除货币/短债）\n    size_yi = _cost_parse_size(base.get("assets_size"))\n    if size_yi > 0 and size_yi < DRAWDOWN_MIN_SIZE_YI:\n        return None\n\n    # 3) 申购状态不能暂停\n    buy_status = ((item.get("status") or {}).get("buy_status") or "")\n    if "暂停" in buy_status:\n        return None\n\n    # 4) 近60日是否连续创新低：统计最近 60 日里"连续下跌"的最长天数\n    recent_returns = returns.tail(60)\n    max_consec_drop = 0\n    cur = 0\n    for rv in recent_returns:\n        if rv < 0:\n            cur += 1\n            max_consec_drop = max(max_consec_drop, cur)\n        else:\n            cur = 0\n    if max_consec_drop >= DRAWDOWN_MAX_CONSEC_DROP_DAYS:\n        return None\n\n    # ----- 排雷标签（给观察池参考，不一票否决） -----\n    alerts = []\n    if size_yi > 0 and size_yi < 1.0:\n        alerts.append("规模偏小")\n    if not np.isnan(r1y) and r1y < -0.20:\n        alerts.append("长期弱势")\n    # 反弹无量：近20日累计涨幅 > 5% 但波动率反而下降过快\n    recent20_cum = (1 + returns.tail(20)).prod() - 1 if len(returns) >= 20 else 0\n    if recent20_cum > 0.05 and vol_contraction < 0.5:\n        alerts.append("反弹无量")\n    if max_consec_drop >= 15:\n        alerts.append("近期连跌多日")\n    if "限" in buy_status:\n        alerts.append("限额申购")\n\n    # 子分\n    drawdown_depth_score = float(min(curr_dd_abs / 0.30, 1.0) * 100.0)\n\n    if vol_contraction <= 0.6:\n        contraction_score = 100.0\n    elif vol_contraction <= 0.8:\n        contraction_score = 80.0\n    elif vol_contraction <= 1.0:\n        contraction_score = 60.0\n    elif vol_contraction <= 1.2:\n        contraction_score = 40.0\n    else:\n        contraction_score = 20.0\n\n    stabilization_score = 0.0\n    if r1 > -0.03:\n        stabilization_score += 40\n    if r3 > -0.08:\n        stabilization_score += 30\n    if abs(r1) <= 0.05:\n        stabilization_score += 30\n\n    # 位置分：优先用模块六已算过的评分；没有则用近一年净值区间反向位置\n    pos_s = None\n    if position_s is not None and code in position_s.index:\n        v = position_s.get(code)\n        if v is not None and not (isinstance(v, float) and pd.isna(v)):\n            pos_s = float(v)\n    if pos_s is None:\n        pos_s = _dd_nav_position_score(nav)\n\n    rebound_shock_score = round(\n        drawdown_depth_score * 0.30\n        + contraction_score * 0.25\n        + stabilization_score * 0.25\n        + pos_s * 0.20, 2)\n\n    label = _dd_status_label(mdd_abs, curr_dd_abs, progress, r1, r3, vol_contraction)\n\n    if rebound_shock_score >= 75 and label == "底部震荡":\n        level = "A-重点观察"\n    elif rebound_shock_score >= 60:\n        level = "B-观察"\n    elif rebound_shock_score >= 45:\n        level = "C-谨慎观察"\n    else:\n        level = "D-剔除"\n\n    fmt_pct = lambda x: f"{x * 100:.2f}%"\n    return {\n        "基金代码": code,\n        "基金名称": item.get("fund_name"),\n        "基金类型": ftype,\n        "成立年限": round(age_m / 12.0, 2),\n        "历史最大回撤": fmt_pct(mdd),\n        "当前回撤": fmt_pct(curr_dd),\n        "回撤进度": f"{progress * 100:.1f}%",\n        "近1月": fmt_pct(r1),\n        "近3月": fmt_pct(r3),\n        "近1年": fmt_pct(r1y) if not np.isnan(r1y) else "--",\n        "vol20": round(vol20, 4) if not np.isnan(vol20) else "--",\n        "vol60": round(vol60, 4) if not np.isnan(vol60) else "--",\n        "波动收敛率": round(vol_contraction, 3) if vol_contraction < 999 else "--",\n        "回撤深度分": round(drawdown_depth_score, 2),\n        "波动收敛分": round(contraction_score, 2),\n        "企稳分": round(stabilization_score, 2),\n        "位置分": round(pos_s, 2),\n        "回撤震荡总分": rebound_shock_score,\n        "状态标签": label,\n        "观察等级": level,\n        "排雷标签": " / ".join(alerts) if alerts else "—",\n        "基金经理": base.get("manager"),\n        "规模": base.get("assets_size"),\n        "申购状态": buy_status or "--",\n        "最新净值": perf.get("nav", "--"),\n        "净值日期": perf.get("nav_date", "--"),\n    }\n\n\ndef _beautify_drawdown_excel(path):\n    wb = load_workbook(path)\n    for ws in wb.worksheets:\n        ws.freeze_panes = "A2"\n        headers = [str(c.value) for c in ws[1]]\n\n        header_fill = PatternFill("solid", fgColor="C00000")\n        header_font = Font(bold=True, color="FFFFFF")\n        for cell in ws[1]:\n            cell.fill = header_fill\n            cell.font = header_font\n            cell.alignment = Alignment(horizontal="center", vertical="center")\n            col_txt = str(cell.value)\n            if col_txt in DRAWDOWN_METRIC_EXPLAIN:\n                text = DRAWDOWN_METRIC_EXPLAIN[col_txt]\n                cmt = Comment(text, "FundSystem")\n                cmt.width = 320\n                cmt.height = 50 + text.count("\\n") * 22\n                cell.comment = cmt\n\n        score_cols = {"回撤震荡总分", "回撤深度分", "波动收敛分", "企稳分", "位置分"}\n        pct_cols = {"历史最大回撤", "当前回撤", "近1月", "近3月"}\n        for row in ws.iter_rows(min_row=2):\n            for cell in row:\n                col_name = headers[cell.column - 1]\n                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")\n\n                if col_name in score_cols and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(cell.value)\n                        color = _score_fill_color(v)\n                        if color:\n                            cell.fill = PatternFill("solid", fgColor=color)\n                        if col_name == "回撤震荡总分":\n                            cell.font = Font(bold=True, size=12)\n                        else:\n                            cell.font = Font(bold=True)\n                    except Exception:\n                        pass\n\n                if col_name == "观察等级" and cell.value:\n                    val = str(cell.value)\n                    if val.startswith("A"):\n                        cell.fill = PatternFill("solid", fgColor="00B050")\n                        cell.font = Font(bold=True, color="FFFFFF")\n                    elif val.startswith("B"):\n                        cell.fill = PatternFill("solid", fgColor="92D050")\n                        cell.font = Font(bold=True)\n                    elif val.startswith("C"):\n                        cell.fill = PatternFill("solid", fgColor="FFFF00")\n                        cell.font = Font(bold=True)\n                    cell.alignment = Alignment(horizontal="center", vertical="center")\n\n                if col_name == "状态标签" and cell.value:\n                    val = str(cell.value)\n                    if val == "底部震荡":\n                        cell.fill = PatternFill("solid", fgColor="C6EFCE")\n                        cell.font = Font(bold=True, color="006100")\n                    elif val == "深度回撤-仍在下跌":\n                        cell.fill = PatternFill("solid", fgColor="FFC7CE")\n                        cell.font = Font(bold=True, color="9C0006")\n                    elif val == "强反弹-等待回踩":\n                        cell.fill = PatternFill("solid", fgColor="FFEB9C")\n                        cell.font = Font(bold=True, color="9C6500")\n                    elif val == "回撤修复中":\n                        cell.fill = PatternFill("solid", fgColor="DDEBF7")\n                        cell.font = Font(bold=True, color="1F4E78")\n\n                if col_name == "排雷标签" and cell.value and str(cell.value) not in ("", "—", "--"):\n                    cell.fill = PatternFill("solid", fgColor="FFF2CC")\n                    cell.font = Font(bold=True, color="9C6500")\n\n                if col_name in pct_cols and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(str(cell.value).replace("%", ""))\n                        cell.font = Font(color="FF0000" if v > 0 else "00B050")\n                    except Exception:\n                        pass\n\n        # 列宽\n        for col in ws.columns:\n            max_len = 0\n            col_letter = get_column_letter(col[0].column)\n            for cell in col:\n                if cell.value:\n                    try:\n                        curr = len(str(cell.value).encode("gbk"))\n                    except Exception:\n                        curr = len(str(cell.value))\n                    if curr > max_len:\n                        max_len = curr\n            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)\n    wb.save(path)\n\n\ndef run_drawdown_shock_screen(log):\n    """筛选法二：回撤震荡观察池"""\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，启动回撤震荡筛选...")\n        log("硬筛条件：MDD≥15% 且 当前回撤≥8% 且 进度≥40% 且 "\n            "近1月∈[-8%, 10%] 且 近3月≥-15% 且 波动收敛率≤1.1")\n\n        # 先把模块六位置评分算出来，有则复用\n        try:\n            position_s = _compute_position_scores(results)\n        except Exception:\n            position_s = None\n\n        hits = []\n        rejected = 0\n        for item in results:\n            res = _dd_calc_one(item, position_s)\n            if res is None:\n                rejected += 1\n                continue\n            hits.append(res)\n\n        if not hits:\n            log("⚠️ 未命中任何基金（可能当前市场没有合格回撤震荡标的）。")\n            return None\n\n        log(f"✅ 筛出 {len(hits)} 只候选基金（{rejected} 只被过滤）")\n\n        df = pd.DataFrame(hits).sort_values("回撤震荡总分", ascending=False, na_position="last")\n\n        # 分层\n        a_pool = df[df["观察等级"].str.startswith("A")]\n        b_pool = df[df["观察等级"].str.startswith("B")]\n        c_pool = df[df["观察等级"].str.startswith("C")]\n        d_pool = df[df["观察等级"].str.startswith("D")]\n        log(f"   A-重点观察: {len(a_pool)} | B-观察: {len(b_pool)} | "\n            f"C-谨慎: {len(c_pool)} | D-剔除: {len(d_pool)}")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"回撤震荡筛选_{ts}.xlsx")\n\n        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:\n            df.to_excel(writer, sheet_name="📋 全量候选池", index=False)\n            if not a_pool.empty:\n                a_pool.to_excel(writer, sheet_name="⭐ A-重点观察", index=False)\n            if not b_pool.empty:\n                b_pool.to_excel(writer, sheet_name="🔍 B-观察", index=False)\n            if not c_pool.empty:\n                c_pool.to_excel(writer, sheet_name="⚠️ C-谨慎观察", index=False)\n            # 按状态标签单独拆一个"底部震荡"聚焦 sheet\n            bottom = df[df["状态标签"] == "底部震荡"]\n            if not bottom.empty:\n                bottom.to_excel(writer, sheet_name="🎯 底部震荡", index=False)\n\n        _beautify_drawdown_excel(out_path)\n\n        # 打印 Top 3\n        top = df.head(3)\n        for _, row in top.iterrows():\n            log(f"  🎯 {row[\'基金名称\']}  回撤震荡={row[\'回撤震荡总分\']} "\n                f"[{row[\'状态标签\']}] [{row[\'观察等级\']}]")\n\n        log(f"回撤震荡筛选完成！")\n        log(f"已导出: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"回撤震荡筛选出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ========================================================\n# ⑮ 其他综合策略（趋势突破 / 低波稳健 / 超跌反弹）\n# ========================================================\n#\n# 复用已有的 _compute_xxx 系列工具，按不同权重与硬筛形成独立策略榜。\n# 每个策略回答一个具体问题，而不是再做一个全量总分。\n# ========================================================\n\n\ndef _strategy_write_excel(df, out_path, score_col, extra_sheets=None, comments=None,\n                           header_fill_hex="1F4E78"):\n    """多策略通用的 Excel 输出：主表 + extra_sheets（dict of name -> df），\n    并对评分列做条件格式。comments 是 {列名: 批注文字} 用于标题悬停。"""\n    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:\n        df.to_excel(writer, sheet_name="📊 主榜", index=False)\n        if extra_sheets:\n            for name, df_extra in extra_sheets.items():\n                if df_extra is None or df_extra.empty:\n                    continue\n                df_extra.to_excel(writer, sheet_name=name, index=False)\n\n    # 通用美化\n    wb = load_workbook(out_path)\n    header_fill = PatternFill("solid", fgColor=header_fill_hex)\n    header_font = Font(bold=True, color="FFFFFF")\n    for ws in wb.worksheets:\n        ws.freeze_panes = "A2"\n        headers = [str(c.value) for c in ws[1]]\n        for cell in ws[1]:\n            cell.fill = header_fill\n            cell.font = header_font\n            cell.alignment = Alignment(horizontal="center", vertical="center")\n            # 标题批注（若传入）\n            col_txt = str(cell.value)\n            if comments and col_txt in comments:\n                text = comments[col_txt]\n                cmt = Comment(text, "FundSystem")\n                cmt.width = 360\n                cmt.height = 60 + text.count("\\n") * 22\n                cell.comment = cmt\n        for row in ws.iter_rows(min_row=2):\n            for cell in row:\n                col_name = headers[cell.column - 1]\n                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")\n                if col_name == score_col and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(cell.value)\n                        color = _score_fill_color(v)\n                        if color:\n                            cell.fill = PatternFill("solid", fgColor=color)\n                        cell.font = Font(bold=True, size=12)\n                    except Exception:\n                        pass\n                if col_name in ("近1月", "近3月", "近6月", "近1年", "近3年",\n                                "年化收益", "最大回撤", "当前回撤") and cell.value not in (None, "", "--"):\n                    try:\n                        v = float(str(cell.value).replace("%", ""))\n                        cell.font = Font(color="FF0000" if v > 0 else "00B050")\n                    except Exception:\n                        pass\n        # 列宽\n        for col in ws.columns:\n            max_len = 0\n            col_letter = get_column_letter(col[0].column)\n            for cell in col:\n                if cell.value:\n                    try:\n                        curr = len(str(cell.value).encode("gbk"))\n                    except Exception:\n                        curr = len(str(cell.value))\n                    if curr > max_len:\n                        max_len = curr\n            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)\n    wb.save(out_path)\n\n\ndef _build_base_row(item):\n    """生成策略输出的通用列"""\n    perf = item.get("performance", {}) or {}\n    base = item.get("base_info", {}) or {}\n    tags = build_fund_tags(item)\n    return {\n        "基金代码": item.get("fund_code"),\n        "基金名称": item.get("fund_name"),\n        "基金类型": base.get("fund_type", "--"),\n        "标签": " / ".join(tags) if tags else "--",\n        "规模": base.get("assets_size", "--"),\n        "基金经理": base.get("manager", "--"),\n        "成立日期": base.get("setup_date", "--"),\n        "最新净值": perf.get("nav", "--"),\n        "净值日期": perf.get("nav_date", "--"),\n        "近1月": perf.get("1m", "--"),\n        "近3月": perf.get("3m", "--"),\n        "近6月": perf.get("6m", "--"),\n        "近1年": perf.get("1y", "--"),\n        "近3年": perf.get("3y", "--"),\n    }\n\n\n# ---- 通用列批注：策略和专题 Excel 都复用 ----\n_COMMON_EXPLAIN = {\n    "标签":   "基金标签，由 build_fund_tags() 派生。\\n包含：资产大类 / 细分(ETF/QDII/主动) / 主题 / 申购状态。",\n    "规模":   "资产净值（爬取时刻）。仅作展示。",\n    "近1月": "1个月区间涨幅，红色代表上涨、绿色代表下跌。",\n    "近3月": "3个月区间涨幅。",\n    "近6月": "6个月区间涨幅。",\n    "近1年": "1年区间涨幅。",\n    "近3年": "3年区间涨幅。",\n    "年化收益": "基于历史净值计算：(1+总收益)^(365.25/持有天数) - 1。",\n    "最大回撤": "历史净值相对前期最高点的最大跌幅。",\n    "当前回撤": "最新净值相对历史最高点的跌幅。",\n    "夏普比率": "(年化收益 - 2%) / 年化波动率，单位风险超额收益。",\n}\n\n\n# ---------- 策略三：趋势突破 ----------\ndef run_trend_breakout_screen(log):\n    """趋势突破：强势上行 + 动量 + 非过热，适合做短中期突破机会"""\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，启动趋势突破策略...")\n\n        tim_s, tim_detail = _compute_timing_scores_detail(results)\n        pos_s = _compute_position_scores(results)\n        eff_s = _compute_efficiency_scores(results)\n\n        rows = []\n        for item in results:\n            code = item.get("fund_code")\n            perf = item.get("performance", {}) or {}\n            base = item.get("base_info", {}) or {}\n            ftype = base.get("fund_type", "") or ""\n            if any(kw in ftype for kw in ("货币", "短债", "理财")):\n                continue\n\n            r1 = _timing_parse_pct(perf.get("1m"))\n            r3 = _timing_parse_pct(perf.get("3m"))\n            r6 = _timing_parse_pct(perf.get("6m"))\n\n            # 硬筛：r1>0 且 r3>0；非过热\n            if pd.isna(r1) or pd.isna(r3):\n                continue\n            if r1 <= 0 or r3 <= 0:\n                continue\n            if r1 > 0.25:  # 过热，暴涨 >25% 剔除\n                continue\n\n            tim = tim_s.get(code)\n            pos = pos_s.get(code)\n            eff = eff_s.get(code)\n            if tim is None or pd.isna(tim):\n                continue\n\n            # 趋势60% + 效率25% + 位置15%（位置分本身=越低位越高分，直接加权）\n            pos_contrib = pos if (pos is not None and not pd.isna(pos)) else 50.0\n            eff_contrib = eff if (eff is not None and not pd.isna(eff)) else 50.0\n            score = round(float(tim) * 0.60 + float(eff_contrib) * 0.25 + float(pos_contrib) * 0.15, 2)\n\n            label = None\n            if not tim_detail.empty and code in tim_detail.index:\n                label = tim_detail.loc[code].get("dir_label")\n\n            row = _build_base_row(item)\n            row.update({\n                "突破评分": score,\n                "趋势分": tim,\n                "趋势方向": label or "--",\n                "效率分": eff if eff is not None else "--",\n                "位置分": pos if pos is not None else "--",\n            })\n            rows.append(row)\n\n        if not rows:\n            log("⚠️ 未筛出任何趋势突破标的。")\n            return None\n\n        df = pd.DataFrame(rows).sort_values("突破评分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"趋势突破_{ts}.xlsx")\n\n        comments = dict(_COMMON_EXPLAIN)\n        comments.update({\n            "突破评分": ("总分公式：\\n"\n                       "  趋势分 × 60% + 效率分 × 25% + 位置分 × 15%\\n\\n"\n                       "硬筛：r1>0 且 r3>0 且 r1≤25%；剔除货币/短债/理财。\\n"\n                       "位置分本身=越低位越高分，直接加权（鼓励低位上攻）。"),\n            "趋势分":   ("来自模块七 · 趋势择时。\\n"\n                       "基于同类内 m1_rank(25) + m3_rank(20) + m6_rank(15) + 方向分(20)，\\n"\n                       "过热 -15、修复 +20，年龄<6月再×0.9。"),\n            "效率分":   ("来自模块四 · 风险效率。\\n"\n                       "sharpe(40) + calmar(35) + sortino(25) 同类百分位加权。"),\n            "位置分":   ("来自模块六 · 位置估值。\\n"\n                       "净值位置(50) + 回撤位置(30) + 成立以来位置(20)。\\n"\n                       "数值越高代表越偏中低位，突破评分中直接加权 × 15%。"),\n            "趋势方向": "1m/3m/6m 综合方向判断：强势上行/温和上行/短期回调/下行/震荡 等。",\n        })\n        _strategy_write_excel(df, out_path, score_col="突破评分", comments=comments)\n        log(f"✅ 趋势突破筛出 {len(df)} 只。")\n        log(f"已导出: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"趋势突破策略出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ---------- 策略四：低波稳健 ----------\ndef run_low_vol_stable_screen(log):\n    """低波稳健：低波动 + 低回撤 + 正向夏普 + 不过度追求爆发，适合稳健配置"""\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，启动低波稳健策略...")\n\n        risk_s, risk_metrics_df, mdd_map = _compute_risk_scores(results)\n        dd_s = _compute_drawdown_scores(results, precomputed_mdd=mdd_map)\n        ret_s = _compute_return_scores(results)\n        mgr_s, _ = _compute_manager_scores_detail(results)\n        cost_s, _ = _compute_cost_scores_detail(results)\n\n        # 风险 40% + 回撤 30% + 收益 15% + 经理 10% + 成本 5%\n        rows = []\n        for item in results:\n            code = item.get("fund_code")\n            perf = item.get("performance", {}) or {}\n            base = item.get("base_info", {}) or {}\n            ftype = base.get("fund_type", "") or ""\n            if any(kw in ftype for kw in ("货币", "短债", "理财")):\n                continue\n\n            r = risk_s.get(code) if not risk_s.empty else None\n            d = dd_s.get(code) if not dd_s.empty else None\n            rt = ret_s.get(code)\n            mg = mgr_s.get(code)\n            co = cost_s.get(code)\n\n            parts = []\n            w_sum = 0.0\n            score = 0.0\n            for v, w in ((r, 0.40), (d, 0.30), (rt, 0.15), (mg, 0.10), (co, 0.05)):\n                if v is None or pd.isna(v):\n                    continue\n                score += float(v) * w\n                w_sum += w\n                parts.append(float(v))\n            if w_sum < 0.6 or len(parts) < 3:\n                continue  # 缺数据太多则跳过\n            final = round(score / w_sum, 2)\n\n            # 额外惩罚：年化波动率过大的基金降权\n            rm = risk_metrics_df.loc[code] if (not risk_metrics_df.empty and code in risk_metrics_df.index) else None\n            try:\n                vol = float(str(rm["波动率"]).replace("%", "")) / 100.0 if rm is not None else None\n            except Exception:\n                vol = None\n            if vol is not None and vol > 0.35:\n                final = round(final * 0.85, 2)\n            # 不加硬门槛，按分数降序展示（Top 即答案）\n\n            row = _build_base_row(item)\n            row.update({\n                "低波稳健分": final,\n                "风险分": r if r is not None else "--",\n                "回撤分": d if d is not None else "--",\n                "收益分": rt if rt is not None else "--",\n                "经理分": mg if mg is not None else "--",\n                "年化波动率": (rm["波动率"] if rm is not None else "--"),\n                "最大回撤": (rm["最大回撤"] if rm is not None else "--"),\n                "夏普比率": (rm["夏普比率"] if rm is not None else "--"),\n            })\n            rows.append(row)\n\n        if not rows:\n            log("⚠️ 未筛出任何低波稳健标的。")\n            return None\n\n        df = pd.DataFrame(rows).sort_values("低波稳健分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"低波稳健_{ts}.xlsx")\n\n        comments = dict(_COMMON_EXPLAIN)\n        comments.update({\n            "低波稳健分": ("总分公式（缺失项自动归一化）：\\n"\n                         "  风险分 × 40% + 回撤分 × 30% + 收益分 × 15% + 经理分 × 10% + 成本分 × 5%\\n\\n"\n                         "惩罚：年化波动率 > 35% → ×0.85。\\n"\n                         "硬筛：权重覆盖率 < 60% 跳过；剔除货币/短债/理财。"),\n            "风险分":   ("全样本夏普百分位 × sample_factor。\\n"\n                       "sample_factor = min(1, sqrt(nav_days/756))，\\n"\n                       "样本越长越可信。"),\n            "回撤分":   ("全样本最大回撤(绝对值)百分位，越小得分越高，再 × sample_factor。"),\n            "收益分":   ("= 70% × 阶段收益分位加权 + 30% × 正收益周期比率分位。\\n"\n                       "见模块二 收益表现。"),\n            "经理分":   ("模块八：任职年限(40) + 任期年化回报(50) + 稳定性加分(10)。\\n"\n                       "短任期惩罚：< 6月×0.85，<12月×0.92。"),\n            "年化波动率":"基于历史净值：std(daily_return) × √252。",\n        })\n        _strategy_write_excel(df, out_path, score_col="低波稳健分", comments=comments)\n        log(f"✅ 低波稳健筛出 {len(df)} 只。")\n        log(f"已导出: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"低波稳健策略出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ---------- 策略五：超跌反弹 ----------\ndef run_oversold_rebound_screen(log):\n    """超跌反弹：深度回撤 + 近期已经反转向上 + 位置偏低，博短期修复。\n    与"回撤震荡"的区别：这里要的是\'已经启动反弹\'，而回撤震荡要\'还在震荡\'。"""\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n        log(f"共加载 {len(results)} 只基金，启动超跌反弹策略...")\n\n        pos_s = _compute_position_scores(results)\n\n        rows = []\n        for item in results:\n            perf = item.get("performance", {}) or {}\n            base = item.get("base_info", {}) or {}\n            ftype = base.get("fund_type", "") or ""\n            if any(kw in ftype for kw in ("货币", "短债", "理财")):\n                continue\n\n            r1 = _timing_parse_pct(perf.get("1m"))\n            r3 = _timing_parse_pct(perf.get("3m"))\n            r6 = _timing_parse_pct(perf.get("6m"))\n            r1y = _timing_parse_pct(perf.get("1y"))\n            if pd.isna(r1) or pd.isna(r3) or pd.isna(r6):\n                continue\n\n            # 硬筛：6m<-8% 或 1y<-15%（有明显回撤），r1>0 且 r3>-5%（反弹启动）\n            deep_drop = (r6 < -0.08) or (not pd.isna(r1y) and r1y < -0.15)\n            started = (r1 > 0) and (r3 > -0.05)\n            if not (deep_drop and started):\n                continue\n            # 防范：r1 > 20% 已经反弹过头\n            if r1 > 0.20:\n                continue\n\n            # 反弹强度：r1 相对 |r6|\n            rebound_strength = min(1.0, r1 / max(abs(r6), 0.05))\n            # 超跌幅度：r6 越小得分越高（取 -r6 / 0.30）\n            oversold = min(1.0, abs(r6) / 0.30)\n\n            pos = pos_s.get(item.get("fund_code"))\n            pos_bonus = (pos / 100.0) if (pos is not None and not pd.isna(pos)) else 0.5\n\n            # 评分：超跌 40% + 反弹 40% + 位置 20%\n            score = round((oversold * 0.40 + rebound_strength * 0.40 + pos_bonus * 0.20) * 100, 2)\n\n            label = "启动反弹" if r1 > 0.05 else "初步修复"\n\n            row = _build_base_row(item)\n            row.update({\n                "超跌反弹分": score,\n                "反弹状态": label,\n                "超跌幅度(r6)": f"{r6*100:.2f}%",\n                "反弹强度(r1)": f"{r1*100:.2f}%",\n                "位置分": pos if pos is not None else "--",\n            })\n            rows.append(row)\n\n        if not rows:\n            log("⚠️ 未筛出任何超跌反弹标的。")\n            return None\n\n        df = pd.DataFrame(rows).sort_values("超跌反弹分", ascending=False, na_position="last")\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"超跌反弹_{ts}.xlsx")\n\n        comments = dict(_COMMON_EXPLAIN)\n        comments.update({\n            "超跌反弹分": ("总分公式：\\n"\n                         "  超跌幅度 × 40% + 反弹强度 × 40% + 位置因子 × 20%\\n\\n"\n                         "子项：\\n"\n                         "  超跌幅度 = min(1, |r6| / 30%)\\n"\n                         "  反弹强度 = min(1, r1 / max(|r6|, 5%))\\n"\n                         "  位置因子 = 位置分 / 100\\n\\n"\n                         "硬筛：r6<-8% 或 r1y<-15%；r1>0 且 r3>-5%；r1≤20%。"),\n            "反弹状态":   "r1>5% 视为\'启动反弹\'，否则\'初步修复\'。",\n            "超跌幅度(r6)":"近6月收益率，越负代表跌得越深、反弹空间越大。",\n            "反弹强度(r1)":"近1月收益率，需为正且不超过20%（避免追高）。",\n            "位置分":     "来自模块六，数值越低越接近历史低位，超跌反弹给予加分。",\n        })\n        _strategy_write_excel(df, out_path, score_col="超跌反弹分", comments=comments)\n        log(f"✅ 超跌反弹筛出 {len(df)} 只。")\n        log(f"已导出: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"超跌反弹策略出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ========================================================\n# ⑯ 专题基金池筛选（按关键词 + 行业专项权重）\n# ========================================================\n#\n# 每个专题都有自己的关键词池（基金名 / 类型 匹配）与权重模型。\n# 非同类基金不做同一套评分：专题内部再做同类排序。\n# ========================================================\n\n# ====== 专题模板：不同资产类型使用不同评分偏好 ======\n#   index     : 指数/ETF/ETF联接 —— 重成本、规模、跟踪质量\n#   active    : 主动权益 —— 重经理、长期收益、回撤\n#   bond      : 债券 —— 重回撤、稳定收益、效率\n#   commodity : 商品 —— 重长期、回撤、择时\n#   foreign   : 海外指数 —— 重长期、成本、申购状态\n\n\ndef _infer_topic_template(spec):\n    """推断专题使用的模板类型：index / active / bond / commodity / foreign。\n    用于批注展示；不影响权重计算（权重以 spec.weights 为准）。\n    """\n    if spec.get("template"):\n        return spec["template"]\n    fn = spec.get("filename", "")\n    bond_set = {"债券专题", "短债专题", "纯债专题", "可转债专题",\n                "一级债基专题", "二级债基专题", "美债专题"}\n    if fn in bond_set:\n        return "bond"\n    if fn in ("黄金专题", "原油专题"):\n        return "commodity"\n    if fn in ("主动偏股专题", "灵活配置专题"):\n        return "active"\n    # 含 QDII 海外类：纳斯达克 / 标普500 / 道琼斯 / 日经 / 欧洲 / 美股科技 / 美股医药\n    if fn in ("纳斯达克专题", "标普500专题", "道琼斯专题", "日经225专题", "欧洲专题",\n              "美股科技专题", "美股医药专题"):\n        return "foreign"\n    if fn in ("FOF专题",):\n        return "fof"\n    if fn in ("REITs专题",):\n        return "reits"\n    # 其它（A股宽基 / 行业 / 港股 / 红利）默认指数模板\n    return "index"\n\n\nTOPIC_SPECS = {\n    "纳斯达克": {\n        "keywords": ["纳斯达克", "纳指", "NASDAQ", "Nasdaq", "纳100"],\n        "type_hint": ["QDII", "指数", "ETF", "联接"],\n        "weights": {"cost": 0.20, "risk": 0.15, "timing": 0.15,\n                    "position": 0.15, "long_term": 0.25, "access": 0.10},\n        "filename": "纳斯达克专题",\n    },\n    "标普500": {\n        "keywords": ["标普500", "标普", "S&P500", "SP500", "标普五百"],\n        "type_hint": ["QDII", "指数", "ETF", "联接"],\n        "weights": {"long_term": 0.25, "risk": 0.20, "cost": 0.20,\n                    "drawdown": 0.15, "timing": 0.10, "access": 0.10},\n        "filename": "标普500专题",\n    },\n    "债券": {\n        "keywords": [],  # 用基金类型过滤而非名称\n        "type_hint": ["偏债", "固收", "纯债", "混合债", "可转债"],\n        "weights": {"drawdown": 0.30, "return": 0.25, "efficiency": 0.20,\n                    "cost": 0.15, "manager": 0.10},\n        "filename": "债券专题",\n    },\n    "红利": {\n        "keywords": ["红利", "股息", "高股息", "红利低波"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.20,\n                    "cost": 0.15, "position": 0.10, "manager": 0.10},\n        "filename": "红利专题",\n    },\n    "黄金": {\n        "keywords": ["黄金", "金ETF", "有色黄金"],\n        "type_hint": ["QDII", "ETF", "联接", "商品"],\n        "weights": {"long_term": 0.25, "drawdown": 0.25, "cost": 0.20,\n                    "timing": 0.15, "position": 0.15},\n        "filename": "黄金专题",\n    },\n    "港股科技": {\n        "keywords": ["港股", "恒生科技", "恒生互联网", "中国互联网", "港股通科技", "H股科技"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "drawdown": 0.20, "timing": 0.20,\n                    "position": 0.15, "risk": 0.15, "cost": 0.10},\n        "filename": "港股科技专题",\n    },\n    "医药": {\n        "keywords": ["医药", "医疗", "生物科技", "医疗器械", "创新药"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.20, "risk": 0.15,\n                    "position": 0.15, "timing": 0.15, "cost": 0.10},\n        "filename": "医药专题",\n    },\n    "半导体": {\n        "keywords": ["半导体", "芯片", "集成电路", "存储"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "risk": 0.15, "drawdown": 0.15,\n                    "timing": 0.25, "position": 0.15, "cost": 0.10},\n        "filename": "半导体专题",\n    },\n    "新能源": {\n        "keywords": ["新能源", "光伏", "锂电", "电动车", "碳中和", "储能"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "drawdown": 0.20, "risk": 0.15,\n                    "timing": 0.20, "position": 0.15, "cost": 0.10},\n        "filename": "新能源专题",\n    },\n\n    # ====== A股宽基 ======\n    "沪深300": {\n        "keywords": ["沪深300", "沪深 300"],\n        "type_hint": [],\n        "weights": {"long_term": 0.30, "risk": 0.20, "cost": 0.20, "drawdown": 0.15,\n                    "timing": 0.10, "access": 0.05},\n        "filename": "沪深300专题",\n    },\n    "中证500": {\n        "keywords": ["中证500", "中证 500"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.20, "cost": 0.15,\n                    "position": 0.10, "timing": 0.10},\n        "filename": "中证500专题",\n    },\n    "中证1000": {\n        "keywords": ["中证1000", "中证 1000"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.20, "risk": 0.20, "cost": 0.15,\n                    "timing": 0.10, "position": 0.10},\n        "filename": "中证1000专题",\n    },\n    "上证50": {\n        "keywords": ["上证50", "上证 50"],\n        "type_hint": [],\n        "weights": {"long_term": 0.30, "risk": 0.20, "drawdown": 0.20, "cost": 0.15,\n                    "position": 0.10, "access": 0.05},\n        "filename": "上证50专题",\n    },\n    "创业板": {\n        "keywords": ["创业板"],\n        "exclude_keywords": [],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "timing": 0.22, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.13, "cost": 0.10},\n        "filename": "创业板专题",\n    },\n    "科创板": {\n        "keywords": ["科创50", "科创100", "科创板", "科创"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "timing": 0.22, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.13, "cost": 0.10},\n        "filename": "科创板专题",\n    },\n    "深证100": {\n        "keywords": ["深证100", "深证 100"],\n        "type_hint": [],\n        "weights": {"long_term": 0.28, "risk": 0.20, "drawdown": 0.17, "cost": 0.15,\n                    "timing": 0.10, "position": 0.10},\n        "filename": "深证100专题",\n    },\n\n    # ====== 港股 ======\n    "恒生指数": {\n        "keywords": ["恒生指数", "恒指"],\n        "exclude_keywords": ["恒生科技", "恒生互联网"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.20, "cost": 0.15,\n                    "timing": 0.10, "access": 0.10},\n        "filename": "恒生指数专题",\n    },\n    "恒生科技": {\n        "keywords": ["恒生科技"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.22, "drawdown": 0.20, "risk": 0.15,\n                    "position": 0.13, "cost": 0.10},\n        "filename": "恒生科技专题",\n    },\n    "港股红利": {\n        "keywords": ["港股红利", "港股通红利", "恒生红利", "港股股息"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.22, "risk": 0.18, "cost": 0.15,\n                    "position": 0.10, "timing": 0.10},\n        "filename": "港股红利专题",\n    },\n    "港股互联网": {\n        "keywords": ["港股互联网", "恒生互联网", "中国互联网"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.22, "drawdown": 0.20, "risk": 0.15,\n                    "position": 0.13, "cost": 0.10},\n        "filename": "港股互联网专题",\n    },\n\n    # ====== 美股细分 ======\n    "道琼斯": {\n        "keywords": ["道琼斯", "道指"],\n        "type_hint": [],\n        "weights": {"long_term": 0.30, "risk": 0.20, "drawdown": 0.15, "cost": 0.15,\n                    "timing": 0.10, "access": 0.10},\n        "filename": "道琼斯专题",\n    },\n    "美股科技": {\n        "keywords": ["美股科技", "标普科技", "美国科技", "FAANG", "美股信息"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "timing": 0.22, "risk": 0.18, "drawdown": 0.15,\n                    "cost": 0.12, "access": 0.11},\n        "filename": "美股科技专题",\n    },\n    "美股医药": {\n        "keywords": ["美股医药", "标普生物", "纳斯达克生物", "美国医药"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.20, "risk": 0.18, "timing": 0.12,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "美股医药专题",\n    },\n\n    # ====== 行业：科技 ======\n    "人工智能": {\n        "keywords": ["人工智能", "AI", "智能机器", "机器学习"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.22, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "人工智能专题",\n    },\n    "云计算": {\n        "keywords": ["云计算", "云服务", "软件ETF", "软件指数"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "timing": 0.20, "risk": 0.18, "drawdown": 0.15,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "云计算专题",\n    },\n    "机器人": {\n        "keywords": ["机器人", "人形机器人", "工业机器人", "智能制造"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.22, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "机器人专题",\n    },\n    "通信": {\n        "keywords": ["通信", "5G", "6G"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "risk": 0.18, "drawdown": 0.18, "timing": 0.17,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "通信专题",\n    },\n    "信创": {\n        "keywords": ["信创", "软件发展", "国产软件", "国产替代"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.22, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "信创专题",\n    },\n\n    # ====== 行业：消费/金融/周期 ======\n    "白酒": {\n        "keywords": ["白酒", "酒", "食品饮料"],\n        "exclude_keywords": ["饮料消费"],\n        "type_hint": [],\n        "weights": {"long_term": 0.28, "risk": 0.18, "drawdown": 0.18, "position": 0.15,\n                    "cost": 0.11, "timing": 0.10},\n        "filename": "白酒专题",\n    },\n    "消费": {\n        "keywords": ["消费", "大消费", "消费ETF"],\n        "exclude_keywords": ["新消费"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.18, "position": 0.15,\n                    "timing": 0.12, "cost": 0.10},\n        "filename": "消费专题",\n    },\n    "银行": {\n        "keywords": ["银行"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.22, "risk": 0.18, "cost": 0.15,\n                    "position": 0.10, "timing": 0.10},\n        "filename": "银行专题",\n    },\n    "证券": {\n        "keywords": ["证券", "券商"],\n        "type_hint": [],\n        "weights": {"long_term": 0.20, "timing": 0.25, "drawdown": 0.20, "risk": 0.15,\n                    "position": 0.10, "cost": 0.10},\n        "filename": "证券专题",\n    },\n    "地产": {\n        "keywords": ["地产", "房地产", "房产"],\n        "type_hint": [],\n        "weights": {"drawdown": 0.25, "risk": 0.20, "long_term": 0.18, "position": 0.17,\n                    "timing": 0.10, "cost": 0.10},\n        "filename": "地产专题",\n    },\n    "有色": {\n        "keywords": ["有色", "金属", "稀土"],\n        "exclude_keywords": ["黄金"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "timing": 0.20, "drawdown": 0.18, "risk": 0.15,\n                    "position": 0.15, "cost": 0.10},\n        "filename": "有色专题",\n    },\n    "煤炭": {\n        "keywords": ["煤炭", "煤"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "drawdown": 0.20, "risk": 0.18, "cost": 0.12,\n                    "timing": 0.15, "position": 0.10},\n        "filename": "煤炭专题",\n    },\n\n    # ====== 主动权益 ======\n    "主动偏股": {\n        "keywords": [],\n        "type_hint": ["混合型-偏股", "股票型"],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.18, "manager": 0.15,\n                    "return": 0.12, "cost": 0.10},\n        "filename": "主动偏股专题",\n    },\n    "灵活配置": {\n        "keywords": [],\n        "type_hint": ["混合型-灵活", "混合型-平衡"],\n        "weights": {"long_term": 0.22, "risk": 0.20, "drawdown": 0.20, "manager": 0.15,\n                    "return": 0.13, "cost": 0.10},\n        "filename": "灵活配置专题",\n    },\n\n    # ====== 债券细分 ======\n    "短债": {\n        "keywords": ["短债", "超短债", "现金"],\n        "exclude_keywords": ["可转债"],\n        "type_hint": [],\n        "weights": {"drawdown": 0.30, "return": 0.25, "risk": 0.20, "cost": 0.15,\n                    "manager": 0.10},\n        "filename": "短债专题",\n    },\n    "纯债": {\n        "keywords": ["纯债", "利率债", "国开", "政策性金融债"],\n        "exclude_keywords": ["可转债", "转债"],\n        "type_hint": [],\n        "weights": {"drawdown": 0.28, "return": 0.22, "efficiency": 0.20, "risk": 0.15,\n                    "cost": 0.10, "manager": 0.05},\n        "filename": "纯债专题",\n    },\n    "可转债": {\n        "keywords": ["可转债", "转债"],\n        "type_hint": [],\n        "weights": {"long_term": 0.22, "drawdown": 0.22, "timing": 0.18, "risk": 0.15,\n                    "return": 0.13, "cost": 0.10},\n        "filename": "可转债专题",\n    },\n    "一级债基": {\n        "keywords": [],\n        "type_hint": ["一级债"],\n        "weights": {"drawdown": 0.25, "return": 0.25, "risk": 0.20, "efficiency": 0.15,\n                    "cost": 0.10, "manager": 0.05},\n        "filename": "一级债基专题",\n    },\n    "二级债基": {\n        "keywords": [],\n        "type_hint": ["二级债"],\n        "weights": {"drawdown": 0.25, "return": 0.20, "risk": 0.20, "efficiency": 0.15,\n                    "manager": 0.10, "cost": 0.10},\n        "filename": "二级债基专题",\n    },\n\n    # ====== 海外其它 ======\n    "日经225": {\n        "keywords": ["日经", "日经225", "日本"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.18, "cost": 0.15,\n                    "timing": 0.12, "access": 0.10},\n        "filename": "日经225专题",\n    },\n    "欧洲": {\n        "keywords": ["欧洲", "德国", "法国", "欧元"],\n        "type_hint": [],\n        "weights": {"long_term": 0.25, "risk": 0.20, "drawdown": 0.18, "cost": 0.15,\n                    "timing": 0.12, "access": 0.10},\n        "filename": "欧洲专题",\n    },\n    "美债": {\n        "keywords": ["美债", "美元债", "美国国债"],\n        "type_hint": [],\n        "weights": {"drawdown": 0.30, "return": 0.25, "risk": 0.20, "cost": 0.15,\n                    "manager": 0.10},\n        "filename": "美债专题",\n    },\n\n    # ====== 商品其它 ======\n    "原油": {\n        "keywords": ["原油", "石油", "能源化工"],\n        "exclude_keywords": ["新能源"],\n        "type_hint": [],\n        "weights": {"long_term": 0.18, "drawdown": 0.25, "risk": 0.18, "timing": 0.18,\n                    "position": 0.11, "cost": 0.10},\n        "filename": "原油专题",\n    },\n\n    # ====== REITs ======\n    "REITs": {\n        "keywords": ["REITs", "REITS", "不动产"],\n        "type_hint": ["REITs"],\n        "weights": {"drawdown": 0.28, "risk": 0.22, "return": 0.20, "position": 0.15,\n                    "cost": 0.10, "access": 0.05},\n        "filename": "REITs专题",\n    },\n\n    # ====== FOF ======\n    "FOF": {\n        "keywords": ["FOF", "养老目标", "目标风险", "目标日期"],\n        "type_hint": ["FOF"],\n        "weights": {"long_term": 0.25, "risk": 0.22, "drawdown": 0.20, "manager": 0.15,\n                    "cost": 0.10, "return": 0.08},\n        "filename": "FOF专题",\n    },\n}\n\n\n# ========================================================\n# 四级分类体系：资产大类 → 子类 → 主题 → 策略\n# 驱动 GUI 菜单 Tab 结构与标签系统（build_fund_tags）\n# ========================================================\nTAXONOMY = {\n    "权益类": {\n        "宽基指数": {\n            "A股宽基": ["沪深300", "中证500", "中证1000", "上证50", "创业板", "科创板", "深证100"],\n            "港股宽基": ["恒生指数", "恒生科技", "港股红利", "港股互联网"],\n            "美股宽基": ["纳斯达克", "标普500", "道琼斯", "美股科技", "美股医药"],\n        },\n        "行业指数": {\n            "科技": ["半导体", "人工智能", "云计算", "机器人", "通信", "信创"],\n            "医药": ["医药"],\n            "消费": ["白酒", "消费"],\n            "金融地产": ["银行", "证券", "地产"],\n            "周期": ["有色", "煤炭"],\n            "新能源": ["新能源"],\n        },\n        "红利价值": {\n            "A股红利": ["红利"],\n            "港股红利": ["港股红利"],\n        },\n        "主动权益": {\n            "偏股/灵活": ["主动偏股", "灵活配置"],\n        },\n    },\n    "债券类": {\n        "纯债": {\n            "短中长": ["短债", "纯债"],\n        },\n        "混合债": {\n            "一级/二级": ["一级债基", "二级债基"],\n        },\n        "可转债": {\n            "可转债": ["可转债"],\n        },\n        "海外债": {\n            "美债": ["美债"],\n        },\n    },\n    "海外类": {\n        "美股": {\n            "宽基": ["纳斯达克", "标普500", "道琼斯"],\n            "行业": ["美股科技", "美股医药"],\n        },\n        "港股": {\n            "港股": ["恒生指数", "恒生科技", "港股红利", "港股科技", "港股互联网"],\n        },\n        "其它地区": {\n            "日欧": ["日经225", "欧洲"],\n        },\n    },\n    "商品类": {\n        "贵金属": {\n            "黄金": ["黄金"],\n        },\n        "能源": {\n            "原油": ["原油"],\n        },\n    },\n    "REITs / FOF": {\n        "REITs": {\n            "REITs": ["REITs"],\n        },\n        "FOF": {\n            "FOF": ["FOF"],\n        },\n    },\n}\n\n\n# ---- 标签系统：为每只基金派生多维度标签 ----\ndef build_fund_tags(item, as_dict=False):\n    """根据基金类型 / 名称 / 申购状态 派生结构化标签。\n    返回：\n      - as_dict=False（默认）：已按一二三级合并的字符串列表（兼容旧用法）\n      - as_dict=True：结构化字典，内部程序使用\n        {\n          "asset_class": str,        # 资产大类：权益/债券/海外/商品/REITs/FOF\n          "region": str | None,      # 地区：A股/港股/美股/日本/欧洲/全球 等\n          "product_type": [str],     # ETF / ETF联接 / LOF / 场内 / 主动 等\n          "strategy_type": str | None, # 指数 / 主动 / 混合 / 债券 / FOF ...\n          "theme": [str],            # 沪深300/纳斯达克/半导体/红利 等\n          "risk_style": [str],       # 高波动/成长/低波 等\n          "access": str,             # 可申购 / 暂停申购 / 限额\n        }\n    """\n    name = (item.get("fund_name") or "")\n    ftype = ((item.get("base_info") or {}).get("fund_type") or "")\n    status = ((item.get("status") or {}).get("buy_status") or "")\n\n    # ---------- 一级：资产大类 ----------\n    if any(x in ftype for x in ("偏债", "纯债", "固收", "混合债", "可转债")) or "债" in ftype:\n        asset_class = "债券类"\n    elif "QDII" in ftype:\n        asset_class = "海外类"\n    elif "REITs" in ftype:\n        asset_class = "REITs"\n    elif "FOF" in ftype:\n        asset_class = "FOF"\n    elif "商品" in ftype or any(kw in name for kw in ("黄金", "原油", "能源化工", "白银")):\n        asset_class = "商品类"\n    else:\n        asset_class = "权益类"\n\n    # ---------- 地区 ----------\n    region = None\n    region_map = [\n        ("港股", ("港股", "恒生", "H股", "恒指")),\n        ("美股", ("纳斯达克", "纳指", "标普", "道琼斯", "S&P500", "SP500", "美股", "FAANG", "罗素")),\n        ("日本", ("日经", "日本", "TOPIX")),\n        ("欧洲", ("欧洲", "德国", "法国", "欧元")),\n        ("全球", ("全球", "MSCI", "新兴市场")),\n    ]\n    for reg, kws in region_map:\n        if any(kw in name for kw in kws):\n            region = reg\n            break\n    if region is None and asset_class not in ("海外类",):\n        region = "A股" if asset_class == "权益类" else None\n\n    # ---------- 产品形态 ----------\n    product_type = []\n    if "ETF" in name or "ETF" in ftype:\n        product_type.append("ETF")\n    if "联接" in name:\n        product_type.append("ETF联接")\n    if "LOF" in name:\n        product_type.append("LOF")\n    if "QDII" in name or "QDII" in ftype:\n        product_type.append("QDII")\n    if not product_type and asset_class == "权益类" and "混合型-偏股" in ftype:\n        product_type.append("主动偏股")\n    if not product_type and ("灵活" in ftype or "平衡" in ftype):\n        product_type.append("灵活配置")\n\n    # ---------- 策略类型 ----------\n    if any(kw in ftype for kw in ("指数型", "ETF")):\n        strategy_type = "指数"\n    elif asset_class == "债券类":\n        strategy_type = "债券"\n    elif asset_class == "FOF":\n        strategy_type = "FOF"\n    elif asset_class == "REITs":\n        strategy_type = "REITs"\n    elif asset_class == "商品类":\n        strategy_type = "商品"\n    elif "灵活" in ftype or "平衡" in ftype:\n        strategy_type = "混合"\n    elif "混合型-偏股" in ftype or "股票型" in ftype:\n        strategy_type = "主动"\n    else:\n        strategy_type = None\n\n    # ---------- 三级：主题 ----------\n    theme = []\n    theme_keywords = [\n        ("纳斯达克100", ["纳斯达克", "纳指", "纳100"]),\n        ("标普500",    ["标普500", "S&P500", "SP500"]),\n        ("道琼斯",     ["道琼斯", "道指"]),\n        ("恒生科技",   ["恒生科技"]),\n        ("恒生指数",   ["恒生指数", "恒指"]),\n        ("港股互联网", ["港股互联网", "恒生互联网", "中国互联网"]),\n        ("沪深300",    ["沪深300"]),\n        ("中证500",    ["中证500"]),\n        ("中证1000",   ["中证1000"]),\n        ("上证50",     ["上证50"]),\n        ("创业板",     ["创业板"]),\n        ("科创板",     ["科创50", "科创100", "科创板"]),\n        ("半导体",     ["半导体", "芯片", "集成电路"]),\n        ("人工智能",   ["人工智能", "AI"]),\n        ("新能源",     ["新能源", "光伏", "锂电", "电动车", "储能"]),\n        ("医药",       ["医药", "医疗", "生物科技", "创新药"]),\n        ("消费",       ["消费", "食品饮料"]),\n        ("白酒",       ["白酒"]),\n        ("银行",       ["银行"]),\n        ("证券",       ["证券", "券商"]),\n        ("红利",       ["红利", "股息", "高股息"]),\n        ("黄金",       ["黄金", "金ETF"]),\n        ("原油",       ["原油", "石油"]),\n        ("日经",       ["日经", "日本"]),\n        ("欧洲",       ["欧洲", "德国", "法国"]),\n        ("可转债",     ["可转债", "转债"]),\n        ("短债",       ["短债", "超短债"]),\n        ("纯债",       ["纯债", "利率债", "国开"]),\n    ]\n    for t, kws in theme_keywords:\n        if any(kw in name for kw in kws):\n            theme.append(t)\n\n    # ---------- 风格 ----------\n    risk_style = []\n    if any(kw in name for kw in ("低波", "红利低波", "稳健")):\n        risk_style.append("低波")\n    if any(kw in name for kw in ("成长", "进取", "科技", "半导体", "人工智能")):\n        risk_style.append("成长")\n    if any(kw in name for kw in ("价值", "蓝筹", "基本面", "红利")):\n        risk_style.append("价值")\n    if any(kw in name for kw in ("小盘", "中小", "中证1000", "国证2000")):\n        risk_style.append("小盘")\n    if any(kw in name for kw in ("大盘", "上证50", "沪深300", "MSCI A50")):\n        risk_style.append("大盘")\n\n    # ---------- 交易状态 ----------\n    if "开放申购" in status:\n        access = "可申购"\n    elif "暂停" in status:\n        access = "暂停申购"\n    elif "限" in status:\n        access = "限额"\n    else:\n        access = "未知"\n\n    structured = {\n        "asset_class":   asset_class,\n        "region":        region,\n        "product_type":  product_type,\n        "strategy_type": strategy_type,\n        "theme":         theme,\n        "risk_style":    risk_style,\n        "access":        access,\n    }\n\n    if as_dict:\n        return structured\n\n    # 兼容旧用法：合并成去重字符串列表（用于 Excel 单元格展示）\n    flat = set()\n    flat.add(asset_class)\n    if region:\n        flat.add(region)\n    for p in product_type:\n        flat.add(p)\n    if strategy_type:\n        flat.add(strategy_type)\n    for t in theme:\n        flat.add(t)\n    for r in risk_style:\n        flat.add(r)\n    flat.add(access)\n    return sorted(flat)\n\n\ndef _match_topic(item, spec):\n    """判断基金是否属于某专题。\n    规则：\n      - 有 keywords 时：name 必须命中至少一个关键词，且不能命中 exclude_keywords\n      - 没有 keywords 但有 type_hint 时：fund_type 必须命中\n    """\n    name = (item.get("fund_name") or "")\n    ftype = ((item.get("base_info") or {}).get("fund_type") or "")\n    filename = spec.get("filename", "")\n\n    # 货币/理财/短债 默认排除（债券类专题例外）\n    bond_topics = ("债券专题", "短债专题", "纯债专题", "可转债专题",\n                   "一级债基专题", "二级债基专题", "美债专题")\n    if filename not in bond_topics:\n        for bad in ("货币", "理财"):\n            if bad in ftype:\n                return False\n        if "短债" in ftype:\n            return False\n\n    kws = spec.get("keywords") or []\n    exclude_kws = spec.get("exclude_keywords") or []\n    type_hints = spec.get("type_hint") or []\n\n    # 排除关键词：名称命中任一即剔除\n    for ex in exclude_kws:\n        if ex in name:\n            return False\n\n    if kws:\n        return any(kw in name for kw in kws)\n    if type_hints:\n        return any(h in ftype for h in type_hints)\n    return False\n\n\ndef _compute_long_term_via(results):\n    """包装长期表现评分，外层使用。"""\n    return _compute_long_term_scores(results)\n\n\ndef run_topic_screen(topic_name, log):\n    """统一的专题筛选入口。"""\n    spec = TOPIC_SPECS.get(topic_name)\n    if spec is None:\n        log(f"未知专题：{topic_name}")\n        return None\n\n    try:\n        _, results = _load_latest_json(log)\n        if not results:\n            return None\n\n        # 1. 专题池过滤\n        pool = [it for it in results if _match_topic(it, spec)]\n        log(f"专题 [{topic_name}] 匹配到 {len(pool)} 只基金（全市场 {len(results)}）")\n        if not pool:\n            log("⚠️ 未匹配到任何基金。可能关键词需要调整。")\n            return None\n\n        # 2. 先计算专题内需要的各项分数（在池子里做同类分位更有意义）\n        risk_s, risk_metrics_df, mdd_map = _compute_risk_scores(pool)\n        dd_s = _compute_drawdown_scores(pool, precomputed_mdd=mdd_map)\n        ret_s = _compute_return_scores(pool)\n        long_s = _compute_long_term_via(pool)\n        eff_s = _compute_efficiency_scores(pool)\n        pos_s = _compute_position_scores(pool)\n        tim_s, _ = _compute_timing_scores_detail(pool)\n        mgr_s, _ = _compute_manager_scores_detail(pool)\n        cost_s, cost_detail = _compute_cost_scores_detail(pool)\n\n        # 专题的 access 维度（申购状态子分），从 cost_detail 的状态派生\n        def access_score_for(item):\n            st = (item.get("status") or {}).get("buy_status") or ""\n            return _cost_score_access(st)\n\n        weights = spec["weights"]\n\n        # 评分映射\n        score_map = {\n            "return": ret_s, "risk": risk_s, "drawdown": dd_s, "long_term": long_s,\n            "efficiency": eff_s, "position": pos_s, "timing": tim_s,\n            "manager": mgr_s, "cost": cost_s,\n        }\n\n        rows = []\n        for item in pool:\n            code = item.get("fund_code")\n            total_w = 0.0\n            score = 0.0\n            for key, w in weights.items():\n                if key == "access":\n                    val = access_score_for(item)\n                else:\n                    s = score_map.get(key)\n                    val = s.get(code) if s is not None and not s.empty else None\n                if val is None or pd.isna(val):\n                    continue\n                score += float(val) * w\n                total_w += w\n            if total_w < 0.5:\n                continue\n            final = round(score / total_w, 2)\n\n            row = _build_base_row(item)\n            row.update({\n                "专题评分": final,\n                "专题": topic_name,\n            })\n            # 附带维度分展示\n            for key in weights:\n                if key == "access":\n                    row["申购状态分"] = round(access_score_for(item), 2)\n                else:\n                    s = score_map.get(key)\n                    v = s.get(code) if s is not None and not s.empty else None\n                    row[f"{key}分"] = round(float(v), 2) if (v is not None and not pd.isna(v)) else "--"\n\n            # 成本/规模展示\n            if not cost_detail.empty and code in cost_detail.index:\n                row["申购费率"] = cost_detail.loc[code, "申购费率"]\n                row["规模(亿)"] = cost_detail.loc[code, "规模(亿)"]\n                row["买入状态"] = cost_detail.loc[code, "买入状态"]\n\n            # 风险指标展示（如有）\n            if not risk_metrics_df.empty and code in risk_metrics_df.index:\n                row["年化收益"] = risk_metrics_df.loc[code, "年化收益"]\n                row["最大回撤"] = risk_metrics_df.loc[code, "最大回撤"]\n                row["夏普比率"] = risk_metrics_df.loc[code, "夏普比率"]\n\n            rows.append(row)\n\n        if not rows:\n            log("⚠️ 专题池内无法计算评分（数据不足）。")\n            return None\n\n        df = pd.DataFrame(rows).sort_values("专题评分", ascending=False, na_position="last")\n        top10 = df.head(10).copy()\n\n        OUTPUT_DIR = "fund_excel"\n        os.makedirs(OUTPUT_DIR, exist_ok=True)\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        out_path = os.path.join(OUTPUT_DIR, f"{spec[\'filename\']}_{ts}.xlsx")\n\n        extra = {"🏆 Top 10": top10} if not top10.empty else None\n\n        # 动态生成权重公式 + 各维度说明\n        _DIM_MEANING = {\n            "return":      ("收益分",   "= 70%×阶段收益分位加权 + 30%×正周期比率分位（模块二）"),\n            "risk":        ("风险分",   "= 全样本夏普百分位 × sample_factor"),\n            "drawdown":    ("回撤分",   "= 最大回撤(绝对值)百分位(越小越好) × sample_factor"),\n            "long_term":   ("长期分",   "= 3y(45) + 5y(55) 同类百分位加权，年龄<36月不评分"),\n            "efficiency":  ("效率分",   "= sharpe(40)+calmar(35)+sortino(25) 百分位加权（模块四）"),\n            "position":    ("位置分",   "= 净值位置(50)+回撤位置(30)+成立以来位置(20)（模块六）"),\n            "timing":      ("趋势分",   "= m1_rank(25)+m3_rank(20)+m6_rank(15)+方向(20)（模块七）"),\n            "manager":     ("经理分",   "= 任职年限(40)+任期年化(50)+稳定性加分(10)（模块八）"),\n            "cost":        ("成本分",   "= 费率(40)+规模(40)+申购状态(20)（模块九）"),\n            "access":      ("申购状态分","= 开放申购 100 / 暂停 30 / 其它 60"),\n        }\n        weight_lines = []\n        for key, w in weights.items():\n            label, meaning = _DIM_MEANING.get(key, (key, "--"))\n            weight_lines.append(f"  {label:<8} × {w*100:.0f}%   {meaning}")\n        template = _infer_topic_template(spec)\n        template_desc = {\n            "index":     "指数模板（重成本/规模/申购状态，跟踪质量）",\n            "active":    "主动模板（重经理/长期收益/回撤控制）",\n            "bond":      "债券模板（重回撤/稳定收益/效率）",\n            "commodity": "商品模板（重长期/回撤/择时）",\n            "foreign":   "海外模板（重长期/成本/申购状态）",\n            "reits":     "REITs 模板（重回撤/风险/收益）",\n            "fof":       "FOF 模板（重长期/风险/经理）",\n        }.get(template, "通用模板")\n\n        formula_txt = (f"专题模板：{template_desc}\\n\\n"\n                       "总分公式（缺失项自动归一化；权重覆盖率<50% 不进榜）：\\n"\n                       + "\\n".join(weight_lines))\n\n        comments = dict(_COMMON_EXPLAIN)\n        comments.update({\n            "专题评分": formula_txt,\n            "专题":     f"当前专题 = {topic_name}。\\n匹配规则：见 TOPIC_SPECS[{topic_name}].keywords / exclude_keywords。",\n            "申购费率": "购买时一次性费率；0% 满分，>0.5% 得 20 分。",\n            "规模(亿)": "基金资产净值，5~50亿为最优区间，<0.5亿有清盘风险，>500亿调仓困难。",\n            "规模评级": "倒 U 型评价：最优区间 / 偏小 / 过小⚠️ / 超大(大象转身难)",\n            "买入状态": "开放申购 / 限大额 / 暂停申购 等。",\n        })\n        # 把每个维度分列的批注也加上\n        for key in weights:\n            label, meaning = _DIM_MEANING.get(key, (key, "--"))\n            col_name = "申购状态分" if key == "access" else f"{key}分"\n            comments[col_name] = f"{label}（本专题权重 {weights[key]*100:.0f}%）\\n{meaning}"\n\n        _strategy_write_excel(df, out_path, score_col="专题评分",\n                              extra_sheets=extra, comments=comments)\n\n        for _, row in top10.head(3).iterrows():\n            log(f"  🏆 {row.get(\'基金名称\', \'\')}  {topic_name}专题分 = {row.get(\'专题评分\', \'--\')}")\n\n        log(f"✅ {topic_name}专题筛选完成，共 {len(df)} 只基金。")\n        log(f"已导出: {out_path}")\n        return os.path.abspath(out_path)\n    except Exception as e:\n        import traceback\n        log(f"{topic_name}专题筛选出错: {e}")\n        log(traceback.format_exc())\n        return None\n\n\n# ========================================================\n# ⑰ 市场分析\n# ========================================================\ndef run_analyze(log):\n    try:\n        import akshare as ak\n        log("=" * 70)\n        log("基金市场全景分析")\n        log("=" * 70)\n        all_funds = ak.fund_name_em()\n        log(f"已加载 {len(all_funds):,} 只基金")\n        log("\\n市场分析完成！（详细统计可后续扩展）")\n    except Exception as e:\n        log(f"发生错误: {e}")\n\n\n# ========================================================\n# GUI 主界面\n# ========================================================\nclass FundToolsApp:\n    def __init__(self):\n        self.root = tk.Tk()\n        self.root.title("基金数据工具集成平台 v3.2 - 四级分类 + 标签系统")\n        self.root.geometry("1820x960")\n        self.root.configure(bg="#1e1e2e")\n        self._crawling = False\n        self._paused = False\n        self._auto_shutdown = tk.BooleanVar(value=False)\n        self._scheduler_active = False\n        self._scheduler_thread = None\n        self._scheduler_time = tk.StringVar(value="02:00")\n        self._build_ui()\n\n    def _build_ui(self):\n        # 标题栏\n        header = tk.Frame(self.root, bg="#181825", height=70)\n        header.pack(fill="x")\n        tk.Label(header, text="基金数据工具集成平台 v3.2  |  多策略 + 四级分类 + 标签系统",\n                 font=("微软雅黑", 16, "bold"),\n                 fg="#cdd6f4", bg="#181825").pack(pady=18)\n\n        # 一键运行快捷区 + 搜索 + 定时/关机\n        hero = tk.Frame(self.root, bg="#1e1e2e", pady=8)\n        hero.pack(fill="x", padx=20)\n        tk.Label(hero, text="快捷入口",\n                 font=("微软雅黑", 9, "bold"),\n                 fg="#6c7086", bg="#1e1e2e", width=10, anchor="w").pack(side="left", padx=(4, 8))\n        tk.Button(hero, text="⚡ 一键运行", font=("微软雅黑", 11, "bold"),\n                  width=14, height=2, fg="#1e1e2e", bg="#fab387",\n                  relief="flat", cursor="hand2",\n                  command=self._run_one_click).pack(side="left", padx=4)\n        tk.Checkbutton(hero, text="完成后自动关机", variable=self._auto_shutdown,\n                       font=("微软雅黑", 9),\n                       fg="#cdd6f4", bg="#1e1e2e", selectcolor="#1e1e2e",\n                       activebackground="#1e1e2e", activeforeground="#f9e2af",\n                       cursor="hand2").pack(side="left", padx=(12, 18))\n        tk.Label(hero, text="|", font=("微软雅黑", 10),\n                 fg="#45475a", bg="#1e1e2e").pack(side="left")\n        tk.Label(hero, text="每日定时：", font=("微软雅黑", 9, "bold"),\n                 fg="#6c7086", bg="#1e1e2e").pack(side="left", padx=(10, 2))\n        tk.Entry(hero, textvariable=self._scheduler_time, font=("微软雅黑", 9),\n                 width=6, justify="center", bg="#313244", fg="#cdd6f4",\n                 relief="flat").pack(side="left", padx=(0, 6))\n        self.btn_scheduler_toggle = tk.Button(\n            hero, text="▶ 启动定时", font=("微软雅黑", 9, "bold"),\n            width=10, height=1, fg="#1e1e2e", bg="#7fbf9f",\n            relief="flat", cursor="hand2",\n            command=self._toggle_scheduler\n        )\n        self.btn_scheduler_toggle.pack(side="left", padx=4)\n        self.lbl_scheduler_status = tk.Label(\n            hero, text="", font=("微软雅黑", 9),\n            fg="#6c7086", bg="#1e1e2e"\n        )\n        self.lbl_scheduler_status.pack(side="left", padx=6)\n\n        # ttk 样式\n        style = ttk.Style()\n        try:\n            style.theme_use("clam")\n        except Exception:\n            pass\n        style.configure("TNotebook", background="#1e1e2e", borderwidth=0)\n        style.configure("TNotebook.Tab",\n                        padding=[12, 6], font=("微软雅黑", 10, "bold"),\n                        background="#313244", foreground="#cdd6f4")\n        style.map("TNotebook.Tab",\n                  background=[("selected", "#45475a")],\n                  foreground=[("selected", "#f9e2af")])\n\n        nb_wrap = tk.Frame(self.root, bg="#1e1e2e", pady=4)\n        nb_wrap.pack(fill="x", padx=20)\n        nb = ttk.Notebook(nb_wrap)\n        nb.pack(fill="x")\n\n        # ---------- Tab 1 · 数据中心 ----------\n        self._add_tab_grouped(nb, "📂 数据中心", [\n            ("数据采集", [\n                ("全量导出",  "#89b4fa", self._run_export),\n                ("智能清洗",  "#7fbf9f", self._run_clean),\n                ("开始爬取",  "#fab387", self._run_crawler),\n                ("转Excel",   "#f9e2af", self._run_to_excel),\n            ]),\n            ("运维", [\n                ("市场分析",  "#b4befe", self._run_analyze),\n            ]),\n        ])\n\n        # ---------- Tab 2 · 单模块评分 ----------\n        self._add_tab_grouped(nb, "🧮 单模块评分", [\n            ("收益/风险", [\n                ("收益表现",  "#cba6f7", self._run_performance),\n                ("风险与回撤","#94e2d5", self._run_risk),\n                ("风险效率",  "#74c7ec", self._run_efficiency),\n            ]),\n            ("位置/趋势/归因", [\n                ("归因分析",  "#89dceb", self._run_attribution),\n                ("位置估值",  "#b4befe", self._run_position),\n                ("趋势择时",  "#eba0ac", self._run_timing),\n            ]),\n            ("结构/经理", [\n                ("经理能力",  "#f5c2e7", self._run_manager),\n                ("成本结构",  "#fab387", self._run_cost),\n            ]),\n        ])\n\n        # ---------- Tab 3 · 综合策略 ----------\n        self._add_tab_grouped(nb, "🎯 综合策略", [\n            ("长期配置", [\n                ("长期综合",  "#f9e2af", self._run_composite),\n                ("低波稳健",  "#94e2d5", self._run_low_vol_stable),\n            ]),\n            ("战术机会", [\n                ("趋势突破",  "#fab387", self._run_trend_breakout),\n                ("回撤震荡",  "#f38ba8", self._run_drawdown_shock),\n                ("超跌反弹",  "#cba6f7", self._run_oversold_rebound),\n            ]),\n        ])\n\n        # ---------- Tab 4 · 权益类（专题） ----------\n        self._add_taxonomy_tab(nb, "📈 权益类", TAXONOMY["权益类"])\n\n        # ---------- Tab 5 · 债券类 ----------\n        self._add_taxonomy_tab(nb, "💰 债券类", TAXONOMY["债券类"])\n\n        # ---------- Tab 6 · 海外类 ----------\n        self._add_taxonomy_tab(nb, "🌏 海外类", TAXONOMY["海外类"])\n\n        # ---------- Tab 7 · 商品类 ----------\n        self._add_taxonomy_tab(nb, "🥇 商品类", TAXONOMY["商品类"])\n\n        # ---------- Tab 8 · REITs / FOF ----------\n        self._add_taxonomy_tab(nb, "🏢 REITs / FOF", TAXONOMY["REITs / FOF"])\n\n        # 爬取控制区\n        self.ctrl_frame = tk.Frame(self.root, bg="#2a2a3e", pady=12)\n        ctrl_frame = self.ctrl_frame\n\n        tk.Label(ctrl_frame, text="爬取控制：",\n                 font=("微软雅黑", 10, "bold"), fg="#a6adc8", bg="#2a2a3e").pack(side="left", padx=(12, 6))\n\n        self.btn_pause = tk.Button(\n            ctrl_frame, text="暂停爬取",\n            font=("微软雅黑", 10, "bold"), width=14, height=1,\n            fg="#1e1e2e", bg="#f5c2e7", relief="flat", cursor="hand2",\n            state="disabled", command=self._toggle_pause\n        )\n        self.btn_pause.pack(side="left", padx=6)\n\n        self.btn_stop = tk.Button(\n            ctrl_frame, text="停止爬取",\n            font=("微软雅黑", 10, "bold"), width=14, height=1,\n            fg="#1e1e2e", bg="#f38ba8", relief="flat", cursor="hand2",\n            state="disabled", command=self._stop_crawl\n        )\n        self.btn_stop.pack(side="left", padx=6)\n\n        self.btn_export_now = tk.Button(\n            ctrl_frame, text="导出当前数据",\n            font=("微软雅黑", 10, "bold"), width=16, height=1,\n            fg="#1e1e2e", bg="#7fbf9f", relief="flat", cursor="hand2",\n            state="disabled", command=self._export_current\n        )\n        self.btn_export_now.pack(side="left", padx=6)\n\n        self.lbl_progress = tk.Label(\n            ctrl_frame, text="",\n            font=("微软雅黑", 10), fg="#cba6f7", bg="#2a2a3e"\n        )\n        self.lbl_progress.pack(side="left", padx=14)\n\n        self.lbl_status_dot = tk.Label(\n            ctrl_frame, text="●", font=("微软雅黑", 14),\n            fg="#45475a", bg="#2a2a3e"\n        )\n        self.lbl_status_dot.pack(side="right", padx=12)\n        self.lbl_status_text = tk.Label(\n            ctrl_frame, text="就绪",\n            font=("微软雅黑", 10), fg="#6c7086", bg="#2a2a3e"\n        )\n        self.lbl_status_text.pack(side="right")\n\n        # 日志区\n        self.log_frame = tk.Frame(self.root, bg="#1e1e2e", padx=30, pady=6)\n        self.log_frame.pack(fill="both", expand=True)\n        tk.Label(self.log_frame, text="运行日志",\n                 font=("微软雅黑", 10), fg="#6c7086", bg="#1e1e2e").pack(anchor="w")\n        self.log_text = scrolledtext.ScrolledText(\n            self.log_frame, font=("Consolas", 10),\n            bg="#181825", fg="#cdd6f4", relief="flat"\n        )\n        self.log_text.pack(fill="both", expand=True)\n        self.log_text.configure(state="disabled")\n\n    # ---------- Notebook Tab 辅助 ----------\n    _TAB_PALETTE = [\n        "#89b4fa", "#7fbf9f", "#fab387", "#f9e2af", "#cba6f7", "#94e2d5",\n        "#74c7ec", "#89dceb", "#b4befe", "#eba0ac", "#f5c2e7", "#f38ba8",\n    ]\n\n    def _add_tab_grouped(self, notebook, tab_title, groups):\n        """groups: [(group_label, [(text, color, cmd), ...]), ...]"""\n        page = tk.Frame(notebook, bg="#1e1e2e", pady=6)\n        notebook.add(page, text=tab_title)\n        for group_label, btns in groups:\n            row = tk.Frame(page, bg="#1e1e2e", pady=4)\n            row.pack(fill="x", padx=6)\n            tk.Label(row, text=group_label,\n                     font=("微软雅黑", 9, "bold"),\n                     fg="#6c7086", bg="#1e1e2e", width=12, anchor="w").pack(side="left", padx=(4, 8))\n            for text, color, cmd in btns:\n                tk.Button(row, text=text, font=("微软雅黑", 10, "bold"),\n                          width=10, height=2,\n                          fg="#1e1e2e", bg=color, relief="flat", cursor="hand2",\n                          command=cmd).pack(side="left", padx=3)\n\n    def _add_taxonomy_tab(self, notebook, tab_title, sub_tree):\n        """按四级分类在 Tab 内显示：每个二级做一行 LabelFrame，三级专题按钮排列。\n        sub_tree: { 二级名: { 三级名: [专题名, ...] } }\n        """\n        page = tk.Frame(notebook, bg="#1e1e2e", pady=6)\n        notebook.add(page, text=tab_title)\n\n        color_idx = 0\n        for level2, level3_dict in sub_tree.items():\n            group_box = tk.LabelFrame(\n                page, text=level2,\n                font=("微软雅黑", 9, "bold"),\n                fg="#f9e2af", bg="#1e1e2e", bd=1,\n                labelanchor="nw", padx=6, pady=4)\n            group_box.pack(fill="x", padx=6, pady=3)\n\n            for level3, topics in level3_dict.items():\n                line = tk.Frame(group_box, bg="#1e1e2e", pady=2)\n                line.pack(fill="x")\n                tk.Label(line, text=level3,\n                         font=("微软雅黑", 9),\n                         fg="#a6adc8", bg="#1e1e2e", width=12, anchor="w").pack(side="left", padx=(2, 8))\n                for topic in topics:\n                    color = self._TAB_PALETTE[color_idx % len(self._TAB_PALETTE)]\n                    color_idx += 1\n                    tk.Button(line, text=topic, font=("微软雅黑", 10, "bold"),\n                              width=10, height=1,\n                              fg="#1e1e2e", bg=color, relief="flat", cursor="hand2",\n                              command=lambda t=topic: self._run_topic(t)).pack(side="left", padx=2)\n\n    def _log(self, msg):\n        self.log_text.configure(state="normal")\n        ts = dt.now().strftime("%H:%M")\n        self.log_text.insert("end", f"[{ts}] {msg}\\n")\n        self.log_text.see("end")\n        self.log_text.configure(state="disabled")\n        self.root.update_idletasks()\n\n    def _run_in_thread(self, func):\n        threading.Thread(target=func, args=(self._log,), daemon=True).start()\n\n    # 功能按钮回调\n    def _run_export(self):   self._run_in_thread(run_export_all)\n    def _run_clean(self):    self._run_in_thread(run_clean_list)\n\n    def _run_to_excel(self):\n        def _task(log):\n            path = run_to_excel(log)\n            if path:\n                self.root.after(0, lambda: self._ask_open_excel(path))\n        threading.Thread(target=_task, args=(self._log,), daemon=True).start()\n\n    def _run_performance(self):\n        """收益表现按钮：自动对最新 JSON 进行评分并输出美化 Excel"""\n        def _task(log):\n            path = run_performance_score(log)\n            if path:\n                self.root.after(0, lambda: self._ask_open_excel(path))\n        threading.Thread(target=_task, args=(self._log,), daemon=True).start()\n\n    def _run_risk(self):\n        """风险与回撤按钮：独立抓取历史净值计算风险指标，输出美化 Excel"""\n        def _task(log):\n            path = run_risk_drawdown(log)\n            if path:\n                self.root.after(0, lambda: self._ask_open_excel(path))\n        threading.Thread(target=_task, args=(self._log,), daemon=True).start()\n\n    def _run_scoring_job(self, func):\n        """通用：在后台线程中运行一个评分函数，结束后询问是否打开 Excel"""\n        def _task(log):\n            path = func(log)\n            if path:\n                self.root.after(0, lambda: self._ask_open_excel(path))\n        threading.Thread(target=_task, args=(self._log,), daemon=True).start()\n\n    def _run_efficiency(self): self._run_scoring_job(run_efficiency_score)\n    def _run_position(self):   self._run_scoring_job(run_position_score)\n    def _run_timing(self):     self._run_scoring_job(run_timing_score)\n    def _run_manager(self):    self._run_scoring_job(run_manager_score)\n    def _run_cost(self):       self._run_scoring_job(run_cost_score)\n    def _run_attribution(self):self._run_scoring_job(run_attribution_score)\n    def _run_composite(self):  self._run_scoring_job(run_composite_score)\n    def _run_drawdown_shock(self): self._run_scoring_job(run_drawdown_shock_screen)\n    def _run_trend_breakout(self):    self._run_scoring_job(run_trend_breakout_screen)\n    def _run_low_vol_stable(self):    self._run_scoring_job(run_low_vol_stable_screen)\n    def _run_oversold_rebound(self):  self._run_scoring_job(run_oversold_rebound_screen)\n\n    def _run_topic(self, topic_name):\n        """专题筛选通用回调"""\n        def _task(log):\n            path = run_topic_screen(topic_name, log)\n            if path:\n                self.root.after(0, lambda: self._ask_open_excel(path))\n        threading.Thread(target=_task, args=(self._log,), daemon=True).start()\n\n    # ---------- 一键运行流水线 ----------\n    def _run_one_click(self):\n        """一键运行：智能清洗 → 开始爬取 → 综合评分 → 打开终极榜单"""\n        if self._crawling:\n            self._log("爬取正在进行中，请先停止再使用一键运行。")\n            return\n\n        has_pool = os.path.exists("target_funds.json")\n        has_data = bool(glob.glob(os.path.join("fund_data", "*.json")))\n\n        # 让用户选择是否重新清洗/爬取\n        do_clean = False\n        do_crawl = True\n        if has_pool:\n            do_clean = messagebox.askyesno(\n                "一键运行",\n                "检测到已有 target_funds.json。\\n\\n"\n                "是否重新运行【智能清洗】以刷新基金池？\\n"\n                "（选【否】将沿用现有 target_funds.json）"\n            )\n        else:\n            # 没有池子，必须先清洗\n            do_clean = True\n\n        if has_data:\n            do_crawl = messagebox.askyesno(\n                "一键运行",\n                "检测到 fund_data 中已有爬取结果。\\n\\n"\n                "是否重新运行【开始爬取】以刷新行情数据？\\n"\n                "（选【否】将直接在现有数据上评分）"\n            )\n        else:\n            # 没有缓存，必须爬\n            do_crawl = True\n\n        confirm = messagebox.askyesno(\n            "一键运行",\n            f"将按以下步骤自动执行：\\n\\n"\n            f"1. 智能清洗 : {\'是\' if do_clean else \'跳过（沿用现有基金池）\'}\\n"\n            f"2. 开始爬取 : {\'是\' if do_crawl else \'跳过（复用现有数据）\'}\\n"\n            f"3. 综合评分 : 是（输出多-Sheet Excel 终极榜单）\\n\\n"\n            f"预计耗时：基金池较大时请耐心等待。\\n是否继续？"\n        )\n        if not confirm:\n            return\n\n        # 把 UI 切到"运行中"状态（借用爬取的控制条）\n        self.ctrl_frame.pack(fill="x", padx=30, pady=(0, 6), before=self.log_frame)\n        self._update_status("一键运行中", "#f9e2af")\n\n        def _pipeline():\n            try:\n                self._log("=" * 60)\n                self._log("⚡ 一键运行：开始")\n                self._log("=" * 60)\n                t_all = time.time()\n\n                # Step 1: 智能清洗\n                if do_clean:\n                    self._log("\\n[1/3] 智能清洗 —— 获取并筛选基金池")\n                    run_clean_list(self._log)\n                else:\n                    self._log("\\n[1/3] 智能清洗 —— 跳过，沿用现有 target_funds.json")\n\n                if not os.path.exists("target_funds.json"):\n                    self._log("❌ target_funds.json 不存在，流水线终止。")\n                    return\n\n                # Step 2: 开始爬取\n                if do_crawl:\n                    self._log("\\n[2/3] 开始爬取 —— 并发抓取 profile + 历史净值")\n                    controller.reset()\n                    self._crawling = True\n                    self._paused = False\n                    # 切爬取控制按钮可用\n                    self.root.after(0, lambda: self._set_crawl_ui(running=True))\n                    try:\n                        run_crawler(\n                            log=self._log,\n                            on_progress=self._on_progress,\n                            on_done=None,  # 我们自己控制\n                        )\n                    finally:\n                        self._crawling = False\n                        self.root.after(0, lambda: self._set_crawl_ui(running=False))\n                        self.root.after(0, lambda: self.lbl_progress.config(text=""))\n                else:\n                    self._log("\\n[2/3] 开始爬取 —— 跳过，复用现有 fund_data")\n\n                # Step 3: 综合评分（一次性跑全部维度）\n                self._log("\\n[3/3] 综合评分 —— 计算8大维度并聚合")\n                path = run_composite_score(self._log)\n\n                elapsed = time.time() - t_all\n                self._log("\\n" + "=" * 60)\n                self._log(f"⚡ 一键运行完成！总耗时 {elapsed:.1f}s")\n                self._log("=" * 60)\n\n                if path:\n                    self.root.after(0, lambda: self._ask_open_excel(path))\n                else:\n                    self._log("❌ 综合评分未生成文件，请检查上游数据。")\n\n                # 一键运行完成后自动关机（如果勾选）\n                if self._auto_shutdown.get():\n                    self._log("💤 自动关机选项已启用，60 秒后关机...")\n                    self._log("（如需取消，请在 60 秒内在本程序或终端执行：shutdown /a）")\n                    try:\n                        if sys.platform == "win32":\n                            os.system("shutdown /s /t 60 /c \\"基金数据工具-一键运行已完成，系统即将关机\\"")\n                        else:\n                            os.system("shutdown -h +1")\n                    except Exception as e2:\n                        self._log(f"关机命令执行失败: {e2}")\n            except Exception as e:\n                import traceback\n                self._log(f"一键运行出错: {e}")\n                self._log(traceback.format_exc())\n            finally:\n                self.root.after(0, lambda: self._update_status("就绪", "#6c7086"))\n\n        threading.Thread(target=_pipeline, daemon=True).start()\n\n    def _run_analyze(self):  self._run_in_thread(run_analyze)\n\n    def _ask_open_excel(self, path):\n        if messagebox.askyesno("导出成功", f"Excel 已保存至：\\n{path}\\n\\n是否立即打开？"):\n            try:\n                if sys.platform == "win32":\n                    os.startfile(path)\n                elif sys.platform == "darwin":\n                    import subprocess; subprocess.call(["open", path])\n                else:\n                    import subprocess; subprocess.call(["xdg-open", path])\n            except Exception as e:\n                self._log(f"无法自动打开文件: {e}")\n\n    # 爬取控制\n    def _run_crawler(self):\n        if self._crawling:\n            self._log("爬取正在进行中，请先停止再重新开始。")\n            return\n\n        controller.reset()\n        self._crawling = True\n        self._paused = False\n        self.ctrl_frame.pack(fill="x", padx=30, pady=(0, 6), before=self.log_frame)\n        self._set_crawl_ui(running=True)\n        self._update_status("爬取中", "#7fbf9f")\n\n        def _thread():\n            run_crawler(\n                log=self._log,\n                on_progress=self._on_progress,\n                on_done=self._on_crawl_done\n            )\n\n        threading.Thread(target=_thread, daemon=True).start()\n\n    def _toggle_pause(self):\n        if not self._crawling:\n            return\n        if not self._paused:\n            controller.pause()\n            self._paused = True\n            self.btn_pause.config(text="继续爬取", bg="#89dceb")\n            self._update_status("已暂停", "#f9e2af")\n            self._log("爬取已暂停。可点击【导出当前数据】检查格式，完成后点【继续爬取】。")\n            self.btn_export_now.config(state="normal")\n        else:\n            controller.resume()\n            self._paused = False\n            self.btn_pause.config(text="暂停爬取", bg="#f5c2e7")\n            self._update_status("爬取中", "#7fbf9f")\n            self._log("爬取继续...")\n\n    def _stop_crawl(self):\n        controller.stop()\n        self._log("正在停止，请稍候...")\n        self.btn_stop.config(state="disabled")\n        self.btn_pause.config(state="disabled")\n\n    def _export_current(self):\n        results = controller.get_results()\n        if not results:\n            self._log("当前没有已爬取的数据，无法导出。")\n            return\n\n        OUTPUT_DIR = "fund_excel"\n        ts = dt.now().strftime("%Y%m%d_%H%M%S")\n        output_path = os.path.join(OUTPUT_DIR, f"fund_partial_{ts}.xlsx")\n        self._log(f"正在导出 {len(results)} 条数据...")\n        save_results_to_excel(results, output_path, self._log)\n\n    def _on_progress(self, ok, total):\n        self.lbl_progress.config(text=f"进度：{ok} / {total} 条")\n\n    def _on_crawl_done(self):\n        self._crawling = False\n        self._paused = False\n        self._set_crawl_ui(running=False)\n        self._update_status("就绪", "#6c7086")\n        results = controller.get_results()\n        if results:\n            self.btn_export_now.config(state="normal")\n\n    def _set_crawl_ui(self, running: bool):\n        state_on  = "normal" if running else "disabled"\n        self.btn_pause.config(state=state_on, text="暂停爬取", bg="#f5c2e7")\n        self.btn_stop.config(state=state_on)\n        self.btn_export_now.config(state="disabled" if running and not self._paused else "disabled")\n        if not running:\n            self.lbl_progress.config(text="")\n\n    def _update_status(self, text, color):\n        self.lbl_status_dot.config(fg=color)\n        self.lbl_status_text.config(text=text, fg=color)\n\n    # ---------- 定时调度 ----------\n    def _toggle_scheduler(self):\n        """启动 / 停止每日定时调度"""\n        if self._scheduler_active:\n            self._stop_scheduler()\n            return\n        time_str = self._scheduler_time.get().strip()\n        if not re.match(r\'^\\d{1,2}:\\d{2}$\', time_str):\n            messagebox.showerror("格式错误", "请输入 HH:MM 格式的时间，如 02:00")\n            return\n        self._scheduler_active = True\n        self._scheduler_thread = threading.Thread(\n            target=self._scheduler_loop, args=(time_str,), daemon=True)\n        self._scheduler_thread.start()\n        self.btn_scheduler_toggle.config(text="⏹ 停止定时", bg="#f38ba8")\n        self.lbl_scheduler_status.config(\n            text=f"⏰ 已设定：每日 {time_str}", fg="#7fbf9f")\n        self._log(f"⏰ 每日定时已启动：将在每日 {time_str} 自动执行一键运行。")\n\n    def _stop_scheduler(self):\n        """停止定时调度"""\n        self._scheduler_active = False\n        self.btn_scheduler_toggle.config(text="▶ 启动定时", bg="#7fbf9f")\n        self.lbl_scheduler_status.config(text="", fg="#6c7086")\n        self._log("⏰ 每日定时已停止。")\n\n    def _scheduler_loop(self, time_str):\n        """定时调度后台循环（使用 schedule 库每 30 秒检查一次）"""\n        try:\n            h, m = map(int, time_str.split(":"))\n        except Exception:\n            return\n        schedule.every().day.at(f"{h:02d}:{m:02d}").do(self._scheduled_run)\n        self._log(f"⏰ 调度器已注册：{h:02d}:{m:02d}")\n        while self._scheduler_active:\n            schedule.run_pending()\n            time.sleep(30)\n        schedule.clear()\n\n    def _scheduled_run(self):\n        """定时触发的执行体——等同于\'一键运行 + 自动关机如果勾了\'"""\n        self._log("=" * 60)\n        self._log("⏰ 定时任务触发！开始自动执行一键运行...")\n        self._log("=" * 60)\n\n        has_pool = os.path.exists("target_funds.json")\n        t_all = time.time()\n\n        try:\n            # 清洗（有池则复用，避免每次都清）\n            if not has_pool:\n                self._log("\\n[定时·1/3] 智能清洗 —— 获取并筛选基金池")\n                run_clean_list(self._log)\n            else:\n                self._log("\\n[定时·1/3] 智能清洗 —— 跳过，沿用现有 target_funds.json")\n\n            if not os.path.exists("target_funds.json"):\n                self._log("❌ target_funds.json 不存在，定时任务终止。")\n                return\n\n            # 爬取\n            self._log("\\n[定时·2/3] 开始爬取 —— 并发抓取 profile + 历史净值")\n            controller.reset()\n            run_crawler(\n                log=self._log,\n                on_progress=None,\n                on_done=None,\n            )\n\n            # 综合评分\n            self._log("\\n[定时·3/3] 综合评分 —— 计算8大维度并聚合")\n            run_composite_score(self._log)\n\n            elapsed = time.time() - t_all\n            self._log("\\n" + "=" * 60)\n            self._log(f"⏰ 定时任务完成！总耗时 {elapsed:.1f}s")\n            self._log("=" * 60)\n        except Exception as e:\n            import traceback\n            self._log(f"⏰ 定时任务出错: {e}")\n            self._log(traceback.format_exc())\n\n        # 自动关机\n        if self._auto_shutdown.get():\n            self._log("💤 自动关机选项已启用，60 秒后关机...")\n            self._log("（如需取消，请在 60 秒内关闭本程序并在终端执行：shutdown /a）")\n            try:\n                if sys.platform == "win32":\n                    os.system("shutdown /s /t 60 /c \\"基金数据工具-定时任务已完成，系统即将关机\\"")\n                else:\n                    os.system("shutdown -h +1")\n            except Exception as e:\n                self._log(f"关机命令执行失败: {e}")\n\n    def run(self):\n        self.root.mainloop()\n\n\n# ========================================================\n# 入口\n# ========================================================\nif __name__ == "__main__":\n    try:\n        import akshare\n        app = FundToolsApp()\n        app.run()\n    except ImportError:\n        py_path = sys.executable\n        print("\\n缺少依赖！请运行以下命令安装：")\n        print(f\'& "{py_path}" -m pip install akshare pandas matplotlib requests openpyxl\')\n        input("按回车退出...")\n',
    'desktop_app.py': '"""\n同花顺风格 - 基金数据桌面可视化系统\nDesktop GUI Application (Tkinter + Matplotlib)\n"""\n\nimport json\nimport os\nimport glob\nimport math\nimport re\nimport threading\nimport tkinter as tk\nfrom tkinter import ttk, messagebox, font as tkfont\nfrom datetime import datetime\n\nimport matplotlib\nmatplotlib.use(\'TkAgg\')\nimport matplotlib.font_manager as fm\nimport matplotlib.pyplot as plt\nfrom matplotlib.backends.backend_tkagg import FigureCanvasTkAgg\nfrom matplotlib.figure import Figure\n\n# 配置 matplotlib 中文字体\n_AVAILABLE_FONTS = [f.name for f in fm.fontManager.ttflist]\nif \'Microsoft YaHei\' in _AVAILABLE_FONTS:\n    _MATPLOTLIB_FONT = \'Microsoft YaHei\'\nelif \'SimHei\' in _AVAILABLE_FONTS:\n    _MATPLOTLIB_FONT = \'SimHei\'\nelse:\n    _MATPLOTLIB_FONT = None\n\nif _MATPLOTLIB_FONT:\n    plt.rcParams[\'font.sans-serif\'] = [_MATPLOTLIB_FONT, \'DejaVu Sans\']\n    plt.rcParams[\'axes.unicode_minus\'] = False\n    _FONT_PROPS = fm.FontProperties(fname=fm.findfont(_MATPLOTLIB_FONT))\nelse:\n    _FONT_PROPS = None\n\ndef _get_font(size=10):\n    """返回 matplotlib 字体属性"""\n    if _FONT_PROPS:\n        return fm.FontProperties(fname=fm.findfont(_MATPLOTLIB_FONT), size=size)\n    return None\n\n# ═══════════════════════════════════════════════════════\n# 主题 / 配色（同花顺深色风格）\n# ═══════════════════════════════════════════════════════\nCOLORS = {\n    \'bg\': \'#1a1a2e\',\n    \'bg2\': \'#16213e\',\n    \'card\': \'#1e2a3a\',\n    \'text\': \'#e0e0e0\',\n    \'text2\': \'#a0a0b0\',\n    \'accent\': \'#e8b830\',\n    \'red\': \'#e74c3c\',\n    \'green\': \'#2ecc71\',\n    \'blue\': \'#3498db\',\n    \'header\': \'#0f3460\',\n    \'row_odd\': \'#1e2a3a\',\n    \'row_even\': \'#243447\',\n    \'border\': \'#2a3a4a\',\n}\n\nFONT_FAMILY = \'Microsoft YaHei\'\n\n# ═══════════════════════════════════════════════════════\n# 数据加载\n# ═══════════════════════════════════════════════════════\n_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))\n_DATA_DIR = os.path.join(_SCRIPT_DIR, \'..\', \'fund_data\')\nif not os.path.isdir(_DATA_DIR):\n    _DATA_DIR = os.path.join(os.getcwd(), \'4.douyin_jijin\', \'fund_data\')\nif not os.path.isdir(_DATA_DIR):\n    _DATA_DIR = os.path.join(os.getcwd(), \'fund_data\')\nDATA_DIR = _DATA_DIR  # 兼容旧引用\n\n\ndef clean_html(raw):\n    """清洗 HTML 标签，提取纯文本"""\n    if raw is None:\n        return \'--\'\n    text = re.sub(r\'<[^>]+>\', \'\', str(raw))\n    text = text.strip()\n    return text if text else \'--\'\n\n\ndef extract_subscription_info(status_data):\n    """从 buy_limit_full 提取限购信息（如"限大额 单日上限50元"、"暂停申购"等）"""\n    raw = status_data.get(\'buy_limit_full\', \'\')\n    if not raw:\n        return \'开放申购\', \'--\'\n    text = clean_html(raw)\n    # 简化常见模式\n    text_orig = text\n    # 提取上限金额\n    limit_amount = \'--\'\n    m = re.search(r\'单日累计购买上限\\s*([\\d.,]+)\\s*元\', text)\n    if m:\n        limit_amount = f\'¥{m.group(1)}/日\'\n    else:\n        m = re.search(r\'单日累计购买上限\\s*([\\d.,]+)\\s*美元\', text)\n        if m:\n            limit_amount = f\'${m.group(1)}/日\'\n    # 提取状态\n    if \'暂停申购\' in text:\n        status_text = \'暂停申购\'\n    elif \'限大额\' in text:\n        status_text = \'限大额\'\n    elif \'开放申购\' in text:\n        status_text = \'开放申购\'\n    else:\n        status_text = text[:20] if len(text) > 20 else text\n    return status_text, limit_amount\n\n\ndef safe_float(val, default=0.0):\n    if val is None or val == \'\' or val == \'--\':\n        return default\n    if isinstance(val, (int, float)):\n        return float(val)\n    try:\n        # 去除百分号、逗号、空格等\n        cleaned = str(val).replace(\'%\', \'\').replace(\',\', \'\').replace(\'，\', \'\').strip()\n        return float(cleaned)\n    except (ValueError, TypeError):\n        return default\n\n\ndef safe_str(val, default=\'--\'):\n    if val is None:\n        return default\n    return str(val)\n\n\ndef get_latest_data_file():\n    if not os.path.isdir(_DATA_DIR):\n        return None\n    pattern = os.path.join(_DATA_DIR, \'fund_profile_*.json\')\n    files = glob.glob(pattern)\n    if not files:\n        return None\n    files.sort(reverse=True)\n    return files[0]\n\n\ndef load_funds():\n    filepath = get_latest_data_file()\n    if not filepath:\n        raise FileNotFoundError(\n            f\'未找到基金数据文件。\\n\'\n            f\'搜索路径: {os.path.join(_DATA_DIR, "fund_profile_*.json")}\\n\'\n            f\'请确保 fund_data 目录包含 fund_profile_*.json 文件。\'\n        )\n    print(f\'[INFO] 正在加载数据文件: {filepath}\')\n    with open(filepath, \'r\', encoding=\'utf-8\') as f:\n        data = json.load(f)\n    print(f\'[INFO] 加载完成: {len(data)} 只基金\')\n    return data\n\n\nTHEME_MAP = {\n    \'科创板\': [\'科创\'],\n    \'创业板\': [\'创业板\'],\n    \'沪深300\': [\'沪深300\'],\n    \'中证500\': [\'中证500\'],\n    \'上证50\': [\'上证50\'],\n    \'纳斯达克\': [\'纳斯达克\', \'纳指\', \'纳斯达\'],\n    \'标普500\': [\'标普500\', \'标普 500\', \'SP500\'],\n    \'恒生科技\': [\'恒生科技\'],\n    \'恒生指数\': [\'恒生\'],\n    \'半导体\': [\'半导体\', \'芯片\'],\n    \'人工智能\': [\'人工智能\', \'AI\'],\n    \'新能源\': [\'新能源\', \'光伏\', \'锂电\', \'风电\', \'储能\'],\n    \'医药\': [\'医药\', \'医疗\', \'生物\', \'制药\', \'中药\'],\n    \'消费\': [\'消费\', \'食品\', \'饮料\', \'白酒\'],\n    \'科技\': [\'科技\', \'TMT\', \'信息\'],\n    \'军工\': [\'军工\', \'国防\'],\n    \'红利\': [\'红利\', \'股息\'],\n    \'债券\': [\'债券\', \'债\', \'纯债\', \'短债\', \'中短债\'],\n    \'可转债\': [\'可转债\', \'可转\'],\n    \'FOF\': [\'FOF\', \'fof\'],\n    \'REITs\': [\'REIT\', \'reit\'],\n    \'原油\': [\'原油\', \'石油\'],\n    \'黄金\': [\'黄金\', \'贵金属\'],\n    \'机器人\': [\'机器人\'],\n    \'低波\': [\'低波\', \'稳健\', \'固收\'],\n}\n\n\ndef classify_theme(name):\n    results = []\n    for theme, keywords in THEME_MAP.items():\n        for kw in keywords:\n            if kw.lower() in name.lower():\n                results.append(theme)\n                break\n    return results\n\n\ndef fmt_pct(val):\n    """格式化百分比"""\n    if val is None:\n        return \'--\'\n    return f\'{val:+.2f}%\'\n\n\ndef fmt_pct_color(val):\n    """返回颜色标记的百分比"""\n    if val is None:\n        return \'--\', COLORS[\'text2\']\n    s = f\'{val:+.2f}%\'\n    if val > 0:\n        return s, COLORS[\'red\']\n    elif val < 0:\n        return s, COLORS[\'green\']\n    return s, COLORS[\'text2\']\n\n\ndef extract_summary(fund):\n    perf = fund.get(\'performance\', {})\n    base = fund.get(\'base_info\', {})\n    status = fund.get(\'status\', {})\n    buy_status_text, buy_limit_amount = extract_subscription_info(status)\n    return {\n        \'fund_code\': fund.get(\'fund_code\', \'\'),\n        \'fund_name\': fund.get(\'fund_name\', \'\'),\n        \'fund_type\': safe_str(base.get(\'fund_type\')),\n        \'risk_level\': safe_str(base.get(\'risk_level\')),\n        \'assets_size\': safe_str(base.get(\'assets_size\')),\n        \'manager\': safe_str(base.get(\'manager\')),\n        \'company\': safe_str(base.get(\'company\')),\n        \'setup_date\': safe_str(base.get(\'setup_date\')),\n        \'nav\': safe_float(perf.get(\'nav\')),\n        \'nav_date\': safe_str(perf.get(\'nav_date\')),\n        \'daily_growth\': safe_float(perf.get(\'daily_growth_rate\')),\n        \'return_1m\': safe_float(perf.get(\'1m\')),\n        \'return_3m\': safe_float(perf.get(\'3m\')),\n        \'return_6m\': safe_float(perf.get(\'6m\')),\n        \'return_1y\': safe_float(perf.get(\'1y\')),\n        \'return_3y\': safe_float(perf.get(\'3y\')),\n        \'return_since\': safe_float(perf.get(\'since\')),\n        \'buy_status\': buy_status_text,\n        \'buy_limit\': buy_limit_amount,\n        \'buy_fee\': safe_str(status.get(\'buy_fee\')),\n        \'sell_status\': safe_str(status.get(\'sell_status\')),\n    }\n\n\n# ═══════════════════════════════════════════════════════\n# 主应用程序\n# ═══════════════════════════════════════════════════════\nclass FundDesktopApp:\n    def __init__(self, root):\n        self.root = root\n        self.root.title(\'基金数据可视化系统 - 同花顺风格\')\n        self.root.geometry(\'1600x900\')\n        self.root.configure(bg=COLORS[\'bg\'])\n        self.root.minsize(1300, 750)\n\n        # 设置图标字体\n        self.default_font = (FONT_FAMILY, 10)\n        self.bold_font = (FONT_FAMILY, 10, \'bold\')\n        self.title_font = (FONT_FAMILY, 13, \'bold\')\n        self.header_font = (FONT_FAMILY, 11, \'bold\')\n        self.small_font = (FONT_FAMILY, 9)\n\n        # 数据\n        self.funds_raw = []\n        self.summaries = []\n        self.filtered = []\n        self.current_sort_col = \'return_1y\'\n        self.current_sort_order = \'desc\'\n        self.current_theme = \'\'\n        self.current_type = \'\'\n        self.search_text = \'\'\n        self.current_page = 1\n        self.per_page = 50\n        self.selected_fund = None\n\n        # 获取最新数据日期\n        self.data_date = \'\'\n        self.data_daily_growth_avg = 0.0\n\n        # 加载数据\n        self.loading = True\n        self._show_loading()\n\n        self.root.after(100, self._load_data_thread)\n\n    def _show_loading(self):\n        self.loading_frame = tk.Frame(self.root, bg=COLORS[\'bg\'])\n        self.loading_frame.pack(expand=True, fill=\'both\')\n        lbl = tk.Label(self.loading_frame, text=\'正在加载基金数据...\\n请稍候\',\n                       font=(FONT_FAMILY, 16), fg=COLORS[\'text\'], bg=COLORS[\'bg\'])\n        lbl.pack(expand=True)\n        self.progress = ttk.Progressbar(self.loading_frame, mode=\'indeterminate\', length=400)\n        self.progress.pack(pady=20)\n        self.progress.start(10)\n\n    def _load_data_thread(self):\n        def load():\n            try:\n                self.funds_raw = load_funds()\n                self.summaries = [extract_summary(f) for f in self.funds_raw]\n                self.filtered = list(self.summaries)\n                # 获取数据日期和平均日涨幅\n                dates = set(s[\'nav_date\'] for s in self.summaries if s[\'nav_date\'] != \'--\')\n                self.data_date = max(dates) if dates else datetime.now().strftime(\'%Y-%m-%d\')\n                daily_vals = [s[\'daily_growth\'] for s in self.summaries if s[\'daily_growth\'] is not None and s[\'daily_growth\'] != 0]\n                self.data_daily_growth_avg = sum(daily_vals) / len(daily_vals) if daily_vals else 0.0\n                self.root.after(0, self._init_ui)\n            except Exception as e:\n                error_msg = str(e)\n                print(f\'[ERROR] 数据加载失败: {error_msg}\')\n                import traceback\n                traceback.print_exc()\n                self.root.after(0, lambda: self._show_load_error(error_msg))\n\n        t = threading.Thread(target=load, daemon=True)\n        t.start()\n\n    def _show_load_error(self, error_msg):\n        self.loading_frame.destroy()\n        error_frame = tk.Frame(self.root, bg=COLORS[\'bg\'])\n        error_frame.pack(expand=True, fill=\'both\')\n        tk.Label(error_frame, text=\'⚠️ 数据加载失败\',\n                 font=(FONT_FAMILY, 18, \'bold\'), fg=COLORS[\'red\'], bg=COLORS[\'bg\']).pack(pady=(100, 20))\n        tk.Label(error_frame, text=error_msg,\n                 font=(FONT_FAMILY, 11), fg=COLORS[\'text2\'], bg=COLORS[\'bg\'],\n                 justify=\'left\').pack(pady=10)\n        diag = f\'脚本目录: {_SCRIPT_DIR}\\n\'\n        diag += f\'数据目录: {_DATA_DIR}\\n\'\n        diag += f\'目录存在: {os.path.isdir(_DATA_DIR)}\\n\'\n        if os.path.isdir(_DATA_DIR):\n            files = os.listdir(_DATA_DIR)\n            diag += f\'文件列表: {files[:5]}...\' if len(files) > 5 else f\'文件列表: {files}\'\n        tk.Label(error_frame, text=diag,\n                 font=(FONT_FAMILY, 9), fg=COLORS[\'text2\'], bg=COLORS[\'bg\'],\n                 justify=\'left\').pack(pady=10)\n\n    def _init_ui(self):\n        self.loading_frame.destroy()\n        self._build_header()\n        self._build_notebook()\n\n    # ── 顶部标题栏 ──────────────────────────────────────\n    def _build_header(self):\n        header_frame = tk.Frame(self.root, bg=COLORS[\'header\'], height=70)\n        header_frame.pack(fill=\'x\')\n        header_frame.pack_propagate(False)\n\n        # 左侧标题\n        left_frame = tk.Frame(header_frame, bg=COLORS[\'header\'])\n        left_frame.pack(side=\'left\', padx=20)\n\n        title = tk.Label(left_frame, text=\'📊  基金数据可视化系统\',\n                         font=(FONT_FAMILY, 16, \'bold\'), fg=COLORS[\'accent\'], bg=COLORS[\'header\'])\n        title.pack(side=\'top\', anchor=\'w\')\n\n        # 数据日期\n        self.date_label = tk.Label(left_frame,\n                                   text=f\'数据日期: {self.data_date}\',\n                                   font=(FONT_FAMILY, 12, \'bold\'),\n                                   fg=COLORS[\'text\'], bg=COLORS[\'header\'])\n        self.date_label.pack(side=\'top\', anchor=\'w\', pady=(2, 0))\n\n        self.refresh_button = tk.Button(left_frame, text=\'🔄 刷新数据\', font=self.default_font,\n                                        fg=COLORS[\'bg\'], bg=COLORS[\'blue\'], relief=\'flat\',\n                                        cursor=\'hand2\', command=self._refresh_data)\n        self.refresh_button.pack(side=\'top\', anchor=\'w\', pady=(6, 0))\n\n        # 右侧统计\n        right_frame = tk.Frame(header_frame, bg=COLORS[\'header\'])\n        right_frame.pack(side=\'right\', padx=20)\n\n        total = len(self.summaries)\n        types_count = len(set(s[\'fund_type\'] for s in self.summaries if s[\'fund_type\'] != \'--\'))\n        companies_count = len(set(s[\'company\'] for s in self.summaries if s[\'company\'] != \'--\'))\n\n        self.total_label = tk.Label(right_frame, text=f\'基金总数: {total:,}\', font=self.bold_font,\n                                    fg=COLORS[\'text\'], bg=COLORS[\'header\'])\n        self.total_label.pack(side=\'top\', anchor=\'e\', pady=1)\n\n        self.types_label = tk.Label(right_frame, text=f\'类型: {types_count}\', font=self.bold_font,\n                                    fg=COLORS[\'blue\'], bg=COLORS[\'header\'])\n        self.types_label.pack(side=\'top\', anchor=\'e\', pady=1)\n\n        self.companies_label = tk.Label(right_frame, text=f\'基金公司: {companies_count}\', font=self.bold_font,\n                                       fg=COLORS[\'accent\'], bg=COLORS[\'header\'])\n        self.companies_label.pack(side=\'top\', anchor=\'e\', pady=1)\n\n    def _update_header_stats(self):\n        total = len(self.summaries)\n        types_count = len(set(s[\'fund_type\'] for s in self.summaries if s[\'fund_type\'] != \'--\'))\n        companies_count = len(set(s[\'company\'] for s in self.summaries if s[\'company\'] != \'--\'))\n        self.date_label.config(text=f\'数据日期: {self.data_date}\')\n        self.total_label.config(text=f\'基金总数: {total:,}\')\n        self.types_label.config(text=f\'类型: {types_count}\')\n        self.companies_label.config(text=f\'基金公司: {companies_count}\')\n\n    def _refresh_data(self):\n        self.refresh_button.config(state=\'disabled\', text=\'刷新中...\')\n        def task():\n            try:\n                self.funds_raw = load_funds()\n                self.summaries = [extract_summary(f) for f in self.funds_raw]\n                self.filtered = list(self.summaries)\n                self.current_page = 1\n                dates = set(s[\'nav_date\'] for s in self.summaries if s[\'nav_date\'] != \'--\')\n                self.data_date = max(dates) if dates else self.data_date\n                self.root.after(0, lambda: [self._update_header_stats(), self._apply_filter()])\n            except Exception as e:\n                self.root.after(0, lambda: messagebox.showerror(\'刷新失败\', str(e)))\n            finally:\n                self.root.after(0, lambda: self.refresh_button.config(state=\'normal\', text=\'🔄 刷新数据\'))\n        threading.Thread(target=task, daemon=True).start()\n\n    # ── 选项卡 ──────────────────────────────────────────\n    def _build_notebook(self):\n        style = ttk.Style()\n        style.theme_use(\'clam\')\n        style.configure(\'TNotebook\', background=COLORS[\'bg\'], borderwidth=0)\n        style.configure(\'TNotebook.Tab\', background=COLORS[\'bg2\'], foreground=COLORS[\'text\'],\n                        padding=[20, 8], font=self.default_font, borderwidth=0)\n        style.map(\'TNotebook.Tab\', background=[(\'selected\', COLORS[\'header\'])],\n                  foreground=[(\'selected\', COLORS[\'accent\'])])\n\n        self.notebook = ttk.Notebook(self.root)\n        self.notebook.pack(expand=True, fill=\'both\', padx=5, pady=5)\n\n        # Tab 1: 基金列表\n        self.tab_list = tk.Frame(self.notebook, bg=COLORS[\'bg\'])\n        self.notebook.add(self.tab_list, text=\' 基金列表 \')\n        self._build_list_tab()\n\n        # Tab 2: 排行榜\n        self.tab_rank = tk.Frame(self.notebook, bg=COLORS[\'bg\'])\n        self.notebook.add(self.tab_rank, text=\' 排行榜 \')\n        self._build_rank_tab()\n\n        # Tab 3: 基金详情\n        self.tab_detail = tk.Frame(self.notebook, bg=COLORS[\'bg\'])\n        self.notebook.add(self.tab_detail, text=\' 基金详情 \')\n        self._build_detail_tab()\n\n        # Tab 4: 市场统计\n        self.tab_stats = tk.Frame(self.notebook, bg=COLORS[\'bg\'])\n        self.notebook.add(self.tab_stats, text=\' 市场统计 \')\n        self._build_stats_tab()\n\n    # ═══════════════════════════════════════════════════════\n    # Tab 1: 基金列表 (全字段表格)\n    # ═══════════════════════════════════════════════════════\n    def _build_list_tab(self):\n        # 顶部工具栏\n        toolbar = tk.Frame(self.tab_list, bg=COLORS[\'bg2\'], height=50)\n        toolbar.pack(fill=\'x\', padx=5, pady=(5, 0))\n        toolbar.pack_propagate(False)\n\n        # 搜索框\n        tk.Label(toolbar, text=\'🔍 搜索:\', font=self.default_font,\n                 fg=COLORS[\'text\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=(10, 5), pady=12)\n        self.search_var = tk.StringVar()\n        self.search_var.trace(\'w\', lambda *a: self._on_search())\n        search_entry = tk.Entry(toolbar, textvariable=self.search_var, font=self.default_font,\n                                bg=COLORS[\'card\'], fg=COLORS[\'text\'], insertbackground=COLORS[\'text\'],\n                                relief=\'flat\', width=25, bd=5)\n        search_entry.pack(side=\'left\', padx=5, pady=10)\n\n        # 主题筛选\n        tk.Label(toolbar, text=\'主题:\', font=self.default_font,\n                 fg=COLORS[\'text\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=(15, 5), pady=12)\n        self.theme_var = tk.StringVar(value=\'全部\')\n        themes_list = [\'全部\'] + sorted(THEME_MAP.keys())\n        theme_cb = ttk.Combobox(toolbar, textvariable=self.theme_var, values=themes_list,\n                                font=self.default_font, width=10, state=\'readonly\')\n        theme_cb.pack(side=\'left\', padx=5, pady=10)\n        theme_cb.bind(\'<<ComboboxSelected>>\', lambda e: self._on_filter_change())\n\n        # 类型筛选\n        tk.Label(toolbar, text=\'类型:\', font=self.default_font,\n                 fg=COLORS[\'text\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=(15, 5), pady=12)\n        self.type_var = tk.StringVar(value=\'全部\')\n        fund_types = [\'全部\'] + sorted(set(s[\'fund_type\'] for s in self.summaries if s[\'fund_type\'] != \'--\'))\n        type_cb = ttk.Combobox(toolbar, textvariable=self.type_var, values=fund_types,\n                               font=self.default_font, width=14, state=\'readonly\')\n        type_cb.pack(side=\'left\', padx=5, pady=10)\n        type_cb.bind(\'<<ComboboxSelected>>\', lambda e: self._on_filter_change())\n\n        # 排序\n        tk.Label(toolbar, text=\'排序:\', font=self.default_font,\n                 fg=COLORS[\'text\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=(15, 5), pady=12)\n        self.sort_var = tk.StringVar(value=\'1年收益率 ↓\')\n        sort_options = [\'1年收益率 ↓\', \'1年收益率 ↑\', \'6月收益率 ↓\', \'6月收益率 ↑\',\n                        \'3月收益率 ↓\', \'3月收益率 ↑\', \'1月收益率 ↓\', \'1月收益率 ↑\',\n                        \'日涨幅 ↓\', \'日涨幅 ↑\']\n        sort_cb = ttk.Combobox(toolbar, textvariable=self.sort_var, values=sort_options,\n                               font=self.default_font, width=12, state=\'readonly\')\n        sort_cb.pack(side=\'left\', padx=5, pady=10)\n        sort_cb.bind(\'<<ComboboxSelected>>\', lambda e: self._on_sort_change())\n\n        # 表格容器\n        table_container = tk.Frame(self.tab_list, bg=COLORS[\'bg\'])\n        table_container.pack(expand=True, fill=\'both\', padx=5, pady=5)\n\n        # Treeview 表格 —— 展示所有字段\n        columns = (\n            \'code\', \'name\', \'type\', \'risk\', \'nav\', \'daily\', \'r1m\', \'r3m\', \'r6m\', \'r1y\', \'r3y\',\n            \'since\', \'assets\', \'manager\', \'company\', \'setup\', \'buy_status\', \'buy_limit\', \'buy_fee\',\n        )\n        self.tree = ttk.Treeview(table_container, columns=columns, show=\'headings\',\n                                 height=20, selectmode=\'browse\')\n\n        headings = {\n            \'code\': (\'代码\', 85),\n            \'name\': (\'基金名称\', 200),\n            \'type\': (\'类型\', 100),\n            \'risk\': (\'风险\', 55),\n            \'nav\': (\'净值\', 70),\n            \'daily\': (\'日涨幅\', 75),\n            \'r1m\': (\'1月收益\', 75),\n            \'r3m\': (\'3月收益\', 75),\n            \'r6m\': (\'6月收益\', 75),\n            \'r1y\': (\'1年收益\', 75),\n            \'r3y\': (\'3年收益\', 75),\n            \'since\': (\'成立以来\', 75),\n            \'assets\': (\'规模\', 90),\n            \'manager\': (\'基金经理\', 80),\n            \'company\': (\'基金公司\', 100),\n            \'setup\': (\'成立日期\', 85),\n            \'buy_status\': (\'申购状态\', 85),\n            \'buy_limit\': (\'限购金额\', 85),\n            \'buy_fee\': (\'申购费率\', 70),\n        }\n\n        for col, (text, width) in headings.items():\n            self.tree.heading(col, text=text, command=lambda c=col: self._on_header_click(c))\n            self.tree.column(col, width=width, anchor=\'center\', minwidth=50)\n\n        # Treeview 样式\n        style = ttk.Style()\n        style.configure(\'Treeview\',\n                        background=COLORS[\'card\'],\n                        foreground=COLORS[\'text\'],\n                        fieldbackground=COLORS[\'card\'],\n                        font=self.small_font,\n                        rowheight=28)\n        style.configure(\'Treeview.Heading\',\n                        background=COLORS[\'header\'],\n                        foreground=COLORS[\'accent\'],\n                        font=(FONT_FAMILY, 9, \'bold\'),\n                        relief=\'flat\')\n        style.map(\'Treeview.Heading\', background=[(\'active\', COLORS[\'bg2\'])])\n        style.map(\'Treeview\',\n                  background=[(\'selected\', COLORS[\'blue\'])],\n                  foreground=[(\'selected\', \'#fff\')])\n\n        # 滚动条\n        y_scroll = ttk.Scrollbar(table_container, orient=\'vertical\', command=self.tree.yview)\n        x_scroll = ttk.Scrollbar(table_container, orient=\'horizontal\', command=self.tree.xview)\n        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)\n\n        self.tree.grid(row=0, column=0, sticky=\'nsew\')\n        y_scroll.grid(row=0, column=1, sticky=\'ns\')\n        x_scroll.grid(row=1, column=0, sticky=\'ew\')\n        table_container.grid_rowconfigure(0, weight=1)\n        table_container.grid_columnconfigure(0, weight=1)\n\n        # 双击查看详情\n        self.tree.bind(\'<Double-Button-1>\', lambda e: self._on_double_click())\n\n        # 底部分页栏\n        self._build_pagination(self.tab_list)\n\n        # 初始加载数据\n        self._apply_filter()\n\n    def _build_pagination(self, parent):\n        self.page_frame = tk.Frame(parent, bg=COLORS[\'bg2\'], height=40)\n        self.page_frame.pack(fill=\'x\', padx=5, pady=(0, 5))\n        self.page_frame.pack_propagate(False)\n\n        self.page_info_label = tk.Label(self.page_frame, text=\'\', font=self.small_font,\n                                        fg=COLORS[\'text2\'], bg=COLORS[\'bg2\'])\n        self.page_info_label.pack(side=\'left\', padx=15)\n\n        btn_frame = tk.Frame(self.page_frame, bg=COLORS[\'bg2\'])\n        btn_frame.pack(side=\'right\', padx=15)\n\n        for text, cmd in [(\'◀◀\', lambda: self._go_page(1)),\n                          (\'◀\', lambda: self._go_page(self.current_page - 1)),\n                          (\'▶\', lambda: self._go_page(self.current_page + 1)),\n                          (\'▶▶\', lambda: self._go_page(self.total_pages))]:\n            btn = tk.Button(btn_frame, text=text, command=cmd,\n                            font=self.small_font, bg=COLORS[\'card\'], fg=COLORS[\'text\'],\n                            relief=\'flat\', width=3, cursor=\'hand2\')\n            btn.pack(side=\'left\', padx=2)\n            btn.bind(\'<Enter>\', lambda e, b=btn: b.configure(bg=COLORS[\'blue\']))\n            btn.bind(\'<Leave>\', lambda e, b=btn: b.configure(bg=COLORS[\'card\']))\n\n    def _on_search(self):\n        self.search_text = self.search_var.get().strip()\n        self.current_page = 1\n        self._apply_filter()\n\n    def _on_filter_change(self):\n        self.current_page = 1\n        self._apply_filter()\n\n    def _on_sort_change(self):\n        val = self.sort_var.get()\n        mapping = {\n            \'1年收益率 ↓\': (\'return_1y\', \'desc\'),\n            \'1年收益率 ↑\': (\'return_1y\', \'asc\'),\n            \'6月收益率 ↓\': (\'return_6m\', \'desc\'),\n            \'6月收益率 ↑\': (\'return_6m\', \'asc\'),\n            \'3月收益率 ↓\': (\'return_3m\', \'desc\'),\n            \'3月收益率 ↑\': (\'return_3m\', \'asc\'),\n            \'1月收益率 ↓\': (\'return_1m\', \'desc\'),\n            \'1月收益率 ↑\': (\'return_1m\', \'asc\'),\n            \'日涨幅 ↓\': (\'daily_growth\', \'desc\'),\n            \'日涨幅 ↑\': (\'daily_growth\', \'asc\'),\n        }\n        self.current_sort_col, self.current_sort_order = mapping.get(val, (\'return_1y\', \'desc\'))\n        self.current_page = 1\n        self._apply_filter()\n\n    def _on_header_click(self, col):\n        col_map = {\n            \'code\': \'fund_code\', \'name\': \'fund_name\', \'type\': \'fund_type\',\n            \'risk\': (\'risk_level\', str), \'nav\': \'nav\', \'daily\': \'daily_growth\',\n            \'r1m\': \'return_1m\', \'r3m\': \'return_3m\', \'r6m\': \'return_6m\',\n            \'r1y\': \'return_1y\', \'r3y\': \'return_3y\',\n            \'since\': \'return_since\', \'assets\': (\'assets_size\', str),\n            \'manager\': (\'manager\', str), \'company\': (\'company\', str),\n            \'setup\': (\'setup_date\', str), \'buy_status\': (\'buy_status\', str),\n            \'buy_limit\': (\'buy_limit\', str), \'buy_fee\': (\'buy_fee\', str),\n        }\n        mapped = col_map.get(col, \'return_1y\')\n        if isinstance(mapped, tuple):\n            sort_col = mapped[0]\n        else:\n            sort_col = mapped\n        if self.current_sort_col == sort_col:\n            self.current_sort_order = \'asc\' if self.current_sort_order == \'desc\' else \'desc\'\n        else:\n            self.current_sort_col = sort_col\n            self.current_sort_order = \'desc\'\n        self.current_page = 1\n        self._apply_filter()\n\n    def _apply_filter(self):\n        data = list(self.summaries)\n\n        # 搜索\n        if self.search_text:\n            q = self.search_text.lower()\n            data = [s for s in data if q in s[\'fund_name\'].lower() or q in s[\'fund_code\']]\n\n        # 主题\n        theme = self.theme_var.get()\n        if theme and theme != \'全部\':\n            data = [s for s in data if theme in classify_theme(s[\'fund_name\'])]\n\n        # 类型\n        ftype = self.type_var.get()\n        if ftype and ftype != \'全部\':\n            data = [s for s in data if ftype == s[\'fund_type\']]\n\n        # 排序\n        reverse = self.current_sort_order == \'desc\'\n        try:\n            data.sort(key=lambda x: safe_float(x.get(self.current_sort_col, 0)), reverse=reverse)\n        except Exception:\n            pass\n\n        self.filtered = data\n        self.total_pages = max(1, math.ceil(len(data) / self.per_page))\n        if self.current_page > self.total_pages:\n            self.current_page = self.total_pages\n        self._render_page()\n\n    def _render_page(self):\n        # 清空\n        for item in self.tree.get_children():\n            self.tree.delete(item)\n\n        start = (self.current_page - 1) * self.per_page\n        end = start + self.per_page\n        page_data = self.filtered[start:end]\n\n        for i, s in enumerate(page_data):\n            bg_tag = \'even\' if i % 2 == 0 else \'odd\'\n            values = (\n                s[\'fund_code\'],\n                s[\'fund_name\'],\n                s[\'fund_type\'],\n                s[\'risk_level\'],\n                f\'{s["nav"]:.4f}\' if s[\'nav\'] else \'--\',\n                fmt_pct(s[\'daily_growth\']),\n                fmt_pct(s[\'return_1m\']),\n                fmt_pct(s[\'return_3m\']),\n                fmt_pct(s[\'return_6m\']),\n                fmt_pct(s[\'return_1y\']),\n                fmt_pct(s[\'return_3y\']),\n                fmt_pct(s[\'return_since\']),\n                s[\'assets_size\'],\n                s[\'manager\'],\n                s[\'company\'],\n                s[\'setup_date\'],\n                s[\'buy_status\'],\n                s[\'buy_limit\'],\n                s[\'buy_fee\'],\n            )\n            self.tree.insert(\'\', \'end\', values=values, tags=(bg_tag,))\n\n        # 颜色标签\n        self.tree.tag_configure(\'odd\', background=COLORS[\'row_odd\'])\n        self.tree.tag_configure(\'even\', background=COLORS[\'row_even\'])\n\n        # 更新分页信息\n        total = len(self.filtered)\n        self.page_info_label.config(\n            text=f\'共 {total:,} 只基金 | 第 {self.current_page}/{self.total_pages} 页 | \'\n                 f\'显示 {start+1}-{min(end, total)}\')\n\n    def _go_page(self, page):\n        if 1 <= page <= self.total_pages:\n            self.current_page = page\n            self._render_page()\n\n    def _on_double_click(self):\n        sel = self.tree.selection()\n        if not sel:\n            return\n        values = self.tree.item(sel[0], \'values\')\n        code = values[0]\n        self._show_detail(code)\n        self.notebook.select(2)\n\n    # ═══════════════════════════════════════════════════════\n    # Tab 2: 排行榜 (全字段)\n    # ═══════════════════════════════════════════════════════\n    def _build_rank_tab(self):\n        r_toolbar = tk.Frame(self.tab_rank, bg=COLORS[\'bg2\'], height=45)\n        r_toolbar.pack(fill=\'x\', padx=5, pady=(5, 0))\n        r_toolbar.pack_propagate(False)\n\n        tk.Label(r_toolbar, text=\'🏆 排行榜 Top 100\', font=self.title_font,\n                 fg=COLORS[\'accent\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=15, pady=10)\n\n        self.rank_sort_var = tk.StringVar(value=\'1年收益率\')\n        periods = [\'1年收益率\', \'6月收益率\', \'3月收益率\', \'1月收益率\', \'日涨幅\']\n        for p in periods:\n            rb = tk.Radiobutton(r_toolbar, text=p, variable=self.rank_sort_var, value=p,\n                                font=self.default_font, fg=COLORS[\'text\'], bg=COLORS[\'bg2\'],\n                                selectcolor=COLORS[\'header\'],\n                                activebackground=COLORS[\'bg2\'],\n                                command=self._render_rank)\n            rb.pack(side=\'right\', padx=5, pady=10)\n\n        rank_container = tk.Frame(self.tab_rank, bg=COLORS[\'bg\'])\n        rank_container.pack(expand=True, fill=\'both\', padx=5, pady=5)\n\n        columns_r = (\n            \'rank\', \'code\', \'name\', \'type\', \'nav\', \'daily\', \'r1m\', \'r3m\', \'r6m\', \'r1y\', \'r3y\',\n            \'since\', \'assets\', \'manager\', \'company\', \'buy_status\', \'buy_limit\', \'buy_fee\',\n        )\n        self.rank_tree = ttk.Treeview(rank_container, columns=columns_r, show=\'headings\', height=22)\n\n        rank_headings = {\n            \'rank\': (\'排名\', 50), \'code\': (\'代码\', 85), \'name\': (\'基金名称\', 180),\n            \'type\': (\'类型\', 90), \'nav\': (\'净值\', 70), \'daily\': (\'日涨幅\', 70),\n            \'r1m\': (\'1月\', 65), \'r3m\': (\'3月\', 65), \'r6m\': (\'6月\', 65),\n            \'r1y\': (\'1年\', 65), \'r3y\': (\'3年\', 65), \'since\': (\'成立以来\', 70),\n            \'assets\': (\'规模\', 80), \'manager\': (\'经理\', 70), \'company\': (\'公司\', 90),\n            \'buy_status\': (\'申购状态\', 80), \'buy_limit\': (\'限购金额\', 80), \'buy_fee\': (\'费率\', 60),\n        }\n        for col, (text, width) in rank_headings.items():\n            self.rank_tree.heading(col, text=text)\n            self.rank_tree.column(col, width=width, anchor=\'center\', minwidth=45)\n\n        rank_scroll_y = ttk.Scrollbar(rank_container, orient=\'vertical\', command=self.rank_tree.yview)\n        rank_scroll_x = ttk.Scrollbar(rank_container, orient=\'horizontal\', command=self.rank_tree.xview)\n        self.rank_tree.configure(yscrollcommand=rank_scroll_y.set, xscrollcommand=rank_scroll_x.set)\n        self.rank_tree.grid(row=0, column=0, sticky=\'nsew\')\n        rank_scroll_y.grid(row=0, column=1, sticky=\'ns\')\n        rank_scroll_x.grid(row=1, column=0, sticky=\'ew\')\n        rank_container.grid_rowconfigure(0, weight=1)\n        rank_container.grid_columnconfigure(0, weight=1)\n\n        self.rank_tree.bind(\'<Double-Button-1>\', lambda e: self._on_rank_double_click())\n\n        self._render_rank()\n\n    def _render_rank(self):\n        for item in self.rank_tree.get_children():\n            self.rank_tree.delete(item)\n\n        period = self.rank_sort_var.get()\n        col_map = {\n            \'1年收益率\': \'return_1y\', \'6月收益率\': \'return_6m\',\n            \'3月收益率\': \'return_3m\', \'1月收益率\': \'return_1m\', \'日涨幅\': \'daily_growth\'\n        }\n        sort_col = col_map.get(period, \'return_1y\')\n\n        data = [s for s in self.summaries if safe_float(s.get(sort_col, 0)) != 0]\n        data.sort(key=lambda x: safe_float(x.get(sort_col, 0)), reverse=True)\n        top100 = data[:100]\n\n        for i, s in enumerate(top100):\n            rank = i + 1\n            medal = [\'🥇\', \'🥈\', \'🥉\'][i] if i < 3 else str(rank)\n            bg_tag = \'top3\' if i < 3 else (\'even\' if i % 2 == 0 else \'odd\')\n            values = (\n                medal, s[\'fund_code\'], s[\'fund_name\'], s[\'fund_type\'],\n                f\'{s["nav"]:.4f}\' if s[\'nav\'] else \'--\',\n                fmt_pct(s[\'daily_growth\']), fmt_pct(s[\'return_1m\']),\n                fmt_pct(s[\'return_3m\']), fmt_pct(s[\'return_6m\']),\n                fmt_pct(s[\'return_1y\']), fmt_pct(s[\'return_3y\']),\n                fmt_pct(s[\'return_since\']), s[\'assets_size\'],\n                s[\'manager\'], s[\'company\'],\n                s[\'buy_status\'], s[\'buy_limit\'], s[\'buy_fee\'],\n            )\n            self.rank_tree.insert(\'\', \'end\', values=values, tags=(bg_tag,))\n\n        self.rank_tree.tag_configure(\'top3\', background=\'#2d1a00\', foreground=COLORS[\'accent\'])\n        self.rank_tree.tag_configure(\'odd\', background=COLORS[\'row_odd\'])\n        self.rank_tree.tag_configure(\'even\', background=COLORS[\'row_even\'])\n\n    def _on_rank_double_click(self):\n        sel = self.rank_tree.selection()\n        if not sel:\n            return\n        values = self.rank_tree.item(sel[0], \'values\')\n        code = values[1]\n        self._show_detail(code)\n        self.notebook.select(2)\n\n    # ═══════════════════════════════════════════════════════\n    # Tab 3: 基金详情\n    # ═══════════════════════════════════════════════════════\n    def _build_detail_tab(self):\n        self.detail_container = tk.Frame(self.tab_detail, bg=COLORS[\'bg\'])\n        self.detail_container.pack(expand=True, fill=\'both\', padx=10, pady=10)\n\n        search_frame = tk.Frame(self.detail_container, bg=COLORS[\'bg2\'], height=45)\n        search_frame.pack(fill=\'x\', pady=(0, 10))\n        search_frame.pack_propagate(False)\n\n        tk.Label(search_frame, text=\'输入基金代码/名称:\', font=self.default_font,\n                 fg=COLORS[\'text\'], bg=COLORS[\'bg2\']).pack(side=\'left\', padx=10, pady=10)\n\n        self.detail_search_var = tk.StringVar()\n        self.detail_search_entry = tk.Entry(search_frame, textvariable=self.detail_search_var,\n                                            font=self.default_font, bg=COLORS[\'card\'],\n                                            fg=COLORS[\'text\'], insertbackground=COLORS[\'text\'],\n                                            relief=\'flat\', width=30, bd=5)\n        self.detail_search_entry.pack(side=\'left\', padx=5, pady=8)\n        self.detail_search_entry.bind(\'<Return>\', lambda e: self._search_detail())\n\n        btn = tk.Button(search_frame, text=\'查询\', command=self._search_detail,\n                        font=self.bold_font, bg=COLORS[\'blue\'], fg=\'#fff\',\n                        relief=\'flat\', padx=20, cursor=\'hand2\')\n        btn.pack(side=\'left\', padx=10, pady=8)\n\n        self.detail_display = tk.Frame(self.detail_container, bg=COLORS[\'bg\'])\n        self.detail_display.pack(expand=True, fill=\'both\')\n\n        self.empty_detail = tk.Label(self.detail_display,\n                                     text=\'👈 请先在基金列表双击或搜索基金代码查看详情\',\n                                     font=(FONT_FAMILY, 13), fg=COLORS[\'text2\'], bg=COLORS[\'bg\'])\n        self.empty_detail.pack(expand=True)\n\n    def _search_detail(self):\n        q = self.detail_search_var.get().strip()\n        if not q:\n            return\n        for f in self.funds_raw:\n            code = f.get(\'fund_code\', \'\')\n            name = f.get(\'fund_name\', \'\')\n            if q == code or q.lower() in name.lower():\n                self._show_detail_full(f)\n                return\n        messagebox.showinfo(\'未找到\', f\'未找到基金: {q}\')\n\n    def _show_detail(self, code):\n        for f in self.funds_raw:\n            if f.get(\'fund_code\') == code:\n                self._show_detail_full(f)\n                return\n\n    def _show_detail_full(self, fund):\n        self.selected_fund = fund\n        for w in self.detail_display.winfo_children():\n            w.destroy()\n\n        perf = fund.get(\'performance\', {})\n        base = fund.get(\'base_info\', {})\n        status = fund.get(\'status\', {})\n        nav_history = fund.get(\'nav_history\', [])\n\n        # 基本信息\n        info_frame = tk.Frame(self.detail_display, bg=COLORS[\'card\'], bd=0,\n                              highlightbackground=COLORS[\'border\'], highlightthickness=1)\n        info_frame.pack(fill=\'x\', pady=(0, 10))\n\n        code = fund.get(\'fund_code\', \'\')\n        name = fund.get(\'fund_name\', \'\')\n\n        tk.Label(info_frame, text=f\'{code}  {name}\', font=(FONT_FAMILY, 15, \'bold\'),\n                 fg=COLORS[\'accent\'], bg=COLORS[\'card\']).pack(anchor=\'w\', padx=15, pady=(10, 5))\n\n        info_grid = tk.Frame(info_frame, bg=COLORS[\'card\'])\n        info_grid.pack(fill=\'x\', padx=15, pady=(5, 10))\n\n        # 提取限购信息\n        buy_status_text, buy_limit_amount = extract_subscription_info(status)\n        buy_fee = safe_str(status.get(\'buy_fee\'))\n\n        info_items = [\n            (\'基金类型\', safe_str(base.get(\'fund_type\'))),\n            (\'风险等级\', safe_str(base.get(\'risk_level\'))),\n            (\'基金规模\', safe_str(base.get(\'assets_size\'))),\n            (\'基金经理\', safe_str(base.get(\'manager\'))),\n            (\'基金公司\', safe_str(base.get(\'company\'))),\n            (\'成立日期\', safe_str(base.get(\'setup_date\'))),\n            (\'最新净值\', f\'{safe_float(perf.get("nav")):.4f}\'),\n            (\'净值日期\', safe_str(perf.get(\'nav_date\'))),\n            (\'日涨幅\', fmt_pct(safe_float(perf.get(\'daily_growth_rate\')))),\n            (\'申购状态\', buy_status_text),\n            (\'限购金额\', buy_limit_amount),\n            (\'申购费率\', buy_fee),\n            (\'赎回状态\', safe_str(status.get(\'sell_status\'))),\n        ]\n\n        for i, (label, value) in enumerate(info_items):\n            row, col = i // 2, i % 2\n            frm = tk.Frame(info_grid, bg=COLORS[\'card\'])\n            frm.grid(row=row, column=col * 2, sticky=\'w\', padx=(0, 30), pady=2)\n            tk.Label(frm, text=f\'{label}:\', font=self.small_font,\n                     fg=COLORS[\'text2\'], bg=COLORS[\'card\']).pack(side=\'left\')\n            tk.Label(frm, text=f\' {value}\', font=self.bold_font,\n                     fg=COLORS[\'text\'], bg=COLORS[\'card\']).pack(side=\'left\')\n\n        # 收益率展示\n        returns_frame = tk.Frame(self.detail_display, bg=COLORS[\'card\'], bd=0,\n                                 highlightbackground=COLORS[\'border\'], highlightthickness=1)\n        returns_frame.pack(fill=\'x\', pady=(0, 10))\n\n        tk.Label(returns_frame, text=\'📈 收益率表现\', font=self.title_font,\n                 fg=COLORS[\'accent\'], bg=COLORS[\'card\']).pack(anchor=\'w\', padx=15, pady=(10, 5))\n\n        ret_grid = tk.Frame(returns_frame, bg=COLORS[\'card\'])\n        ret_grid.pack(fill=\'x\', padx=15, pady=(0, 10))\n\n        periods = [\n            (\'日涨幅\', \'daily_growth_rate\'), (\'1个月\', \'1m\'), (\'3个月\', \'3m\'),\n            (\'6个月\', \'6m\'), (\'1年\', \'1y\'), (\'3年\', \'3y\'), (\'成立以来\', \'since\'),\n        ]\n\n        for i, (label, key) in enumerate(periods):\n            val = safe_float(perf.get(key))\n            color = COLORS[\'red\'] if val > 0 else (COLORS[\'green\'] if val < 0 else COLORS[\'text\'])\n            frm = tk.Frame(ret_grid, bg=COLORS[\'bg2\'], bd=0,\n                           highlightbackground=COLORS[\'border\'], highlightthickness=1)\n            frm.grid(row=0, column=i, padx=5, pady=5, ipadx=10, ipady=10)\n\n            tk.Label(frm, text=label, font=self.small_font,\n                     fg=COLORS[\'text2\'], bg=COLORS[\'bg2\']).pack()\n            tk.Label(frm, text=f\'{val:+.2f}%\', font=(FONT_FAMILY, 14, \'bold\'),\n                     fg=color, bg=COLORS[\'bg2\']).pack()\n\n        # 净值走势图\n        if nav_history:\n            chart_frame = tk.Frame(self.detail_display, bg=COLORS[\'card\'], bd=0,\n                                   highlightbackground=COLORS[\'border\'], highlightthickness=1)\n            chart_frame.pack(expand=True, fill=\'both\')\n\n            tk.Label(chart_frame, text=\'📊 净值走势\', font=self.title_font,\n                     fg=COLORS[\'accent\'], bg=COLORS[\'card\']).pack(anchor=\'w\', padx=15, pady=(10, 0))\n\n            self._draw_nav_chart(chart_frame, nav_history)\n\n    def _draw_nav_chart(self, parent, nav_history):\n        dates = []\n        values = []\n        for item in nav_history:\n            try:\n                d = item.get(\'date\', \'\')\n                v = safe_float(item.get(\'val\'))\n                if d and v > 0:\n                    dates.append(d)\n                    values.append(v)\n            except Exception:\n                continue\n\n        if not values:\n            tk.Label(parent, text=\'暂无净值数据\', font=self.default_font,\n                     fg=COLORS[\'text2\'], bg=COLORS[\'card\']).pack(pady=30)\n            return\n\n        dates.reverse()\n        values.reverse()\n\n        if len(dates) > 180:\n            dates = dates[-180:]\n            values = values[-180:]\n\n        fig = Figure(figsize=(12, 4), dpi=80, facecolor=COLORS[\'card\'])\n        ax = fig.add_subplot(111)\n        ax.set_facecolor(COLORS[\'card\'])\n\n        start_val = values[0]\n        color = COLORS[\'red\'] if values[-1] >= start_val else COLORS[\'green\']\n\n        ax.plot(range(len(values)), values, color=color, linewidth=1.5)\n        ax.fill_between(range(len(values)), values, min(values) * 0.98, alpha=0.1, color=color)\n\n        tick_positions = [i for i, d in enumerate(dates) if i % max(1, len(dates) // 8) == 0]\n        tick_labels = [dates[i] if i < len(dates) else \'\' for i in tick_positions]\n        ax.set_xticks(tick_positions)\n        ax.set_xticklabels(tick_labels, rotation=30, ha=\'right\', fontsize=7, color=COLORS[\'text2\'])\n        ax.tick_params(axis=\'y\', colors=COLORS[\'text2\'], labelsize=8)\n        ax.spines[\'top\'].set_visible(False)\n        ax.spines[\'right\'].set_visible(False)\n        ax.spines[\'left\'].set_color(COLORS[\'border\'])\n        ax.spines[\'bottom\'].set_color(COLORS[\'border\'])\n        ax.grid(axis=\'y\', alpha=0.15, color=COLORS[\'text\'])\n\n        canvas = FigureCanvasTkAgg(fig, parent)\n        canvas.draw()\n        canvas.get_tk_widget().pack(expand=True, fill=\'both\', padx=10, pady=5)\n\n    # ═══════════════════════════════════════════════════════\n    # Tab 4: 市场统计\n    # ═══════════════════════════════════════════════════════\n    def _build_stats_tab(self):\n        stat_container = tk.Frame(self.tab_stats, bg=COLORS[\'bg\'])\n        stat_container.pack(expand=True, fill=\'both\', padx=10, pady=10)\n\n        chart1_frame = tk.Frame(stat_container, bg=COLORS[\'card\'], bd=0,\n                                highlightbackground=COLORS[\'border\'], highlightthickness=1)\n        chart1_frame.pack(side=\'left\', expand=True, fill=\'both\', padx=(0, 5))\n\n        tk.Label(chart1_frame, text=\'基金类型分布\', font=self.title_font,\n                 fg=COLORS[\'accent\'], bg=COLORS[\'card\']).pack(anchor=\'w\', padx=15, pady=(10, 0))\n\n        self._draw_type_pie(chart1_frame)\n\n        chart2_frame = tk.Frame(stat_container, bg=COLORS[\'card\'], bd=0,\n                                highlightbackground=COLORS[\'border\'], highlightthickness=1)\n        chart2_frame.pack(side=\'right\', expand=True, fill=\'both\', padx=(5, 0))\n\n        tk.Label(chart2_frame, text=\'主题分布\', font=self.title_font,\n                 fg=COLORS[\'accent\'], bg=COLORS[\'card\']).pack(anchor=\'w\', padx=15, pady=(10, 0))\n\n        self._draw_theme_bar(chart2_frame)\n\n    def _draw_type_pie(self, parent):\n        fund_types = {}\n        for s in self.summaries:\n            ft = s[\'fund_type\']\n            if ft != \'--\':\n                fund_types[ft] = fund_types.get(ft, 0) + 1\n\n        sorted_types = sorted(fund_types.items(), key=lambda x: -x[1])\n        top10 = sorted_types[:10]\n        other_count = sum(v for _, v in sorted_types[10:])\n        if other_count > 0:\n            top10.append((\'其他\', other_count))\n\n        labels = [t for t, _ in top10]\n        sizes = [c for _, c in top10]\n        colors = plt.cm.tab20([i / len(labels) for i in range(len(labels))])\n\n        fig = Figure(figsize=(5, 4), dpi=80, facecolor=COLORS[\'card\'])\n        ax = fig.add_subplot(111)\n        wedges, texts, autotexts = ax.pie(sizes, labels=None, autopct=\'%1.1f%%\',\n                                          colors=colors, startangle=90,\n                                          textprops={\'color\': COLORS[\'text\'], \'fontsize\': 7})\n        font_props = _get_font(8)\n        title_font_props = _get_font(9)\n        legend_kwargs = dict(\n            title=\'基金类型\', loc=\'center left\',\n            bbox_to_anchor=(1, 0.5),\n            labelcolor=COLORS[\'text\'],\n            facecolor=COLORS[\'card\'], edgecolor=COLORS[\'border\']\n        )\n        if font_props:\n            legend_kwargs[\'prop\'] = font_props\n            legend_kwargs[\'title_fontproperties\'] = title_font_props\n        else:\n            legend_kwargs[\'fontsize\'] = 7\n            legend_kwargs[\'title_fontsize\'] = 8\n        ax.legend(wedges, labels, **legend_kwargs)\n\n        canvas = FigureCanvasTkAgg(fig, parent)\n        canvas.draw()\n        canvas.get_tk_widget().pack(expand=True, fill=\'both\', padx=10, pady=5)\n\n    def _draw_theme_bar(self, parent):\n        theme_counts = {}\n        for s in self.summaries:\n            themes = classify_theme(s[\'fund_name\'])\n            for t in themes:\n                theme_counts[t] = theme_counts.get(t, 0) + 1\n\n        sorted_themes = sorted(theme_counts.items(), key=lambda x: -x[1])[:15]\n        themes, counts = zip(*sorted_themes) if sorted_themes else ([], [])\n\n        fig = Figure(figsize=(5.5, 4), dpi=80, facecolor=COLORS[\'card\'])\n        ax = fig.add_subplot(111)\n        ax.set_facecolor(COLORS[\'card\'])\n\n        y_pos = range(len(themes))\n        bars = ax.barh(y_pos, counts, height=0.6, color=COLORS[\'blue\'], alpha=0.8)\n        ax.set_yticks(y_pos)\n        font_props = _get_font(8)\n        if font_props:\n            ax.set_yticklabels(themes, fontproperties=font_props, color=COLORS[\'text\'])\n            xlabel_font = _get_font(8)\n            ax.set_xlabel(\'基金数量\', fontproperties=xlabel_font, color=COLORS[\'text2\'])\n        else:\n            ax.set_yticklabels(themes, fontsize=8, color=COLORS[\'text\'])\n            ax.set_xlabel(\'基金数量\', fontsize=8, color=COLORS[\'text2\'])\n        ax.tick_params(axis=\'x\', colors=COLORS[\'text2\'], labelsize=7)\n        ax.spines[\'top\'].set_visible(False)\n        ax.spines[\'right\'].set_visible(False)\n        ax.spines[\'left\'].set_color(COLORS[\'border\'])\n        ax.spines[\'bottom\'].set_color(COLORS[\'border\'])\n        ax.invert_yaxis()\n\n        for bar, count in zip(bars, counts):\n            ax.text(bar.get_width() + max(counts) * 0.01, bar.get_y() + bar.get_height() / 2,\n                    str(count), va=\'center\', fontsize=7, color=COLORS[\'text2\'])\n\n        canvas = FigureCanvasTkAgg(fig, parent)\n        canvas.draw()\n        canvas.get_tk_widget().pack(expand=True, fill=\'both\', padx=10, pady=5)\n\n\n# ═══════════════════════════════════════════════════════\n# 启动入口\n# ═══════════════════════════════════════════════════════\nif __name__ == \'__main__\':\n    root = tk.Tk()\n    app = FundDesktopApp(root)\n    root.mainloop()',
    'app.py': '"""\n同花顺风格 - 基金数据可视化系统\nFlask Backend API\n"""\nimport json\nimport os\nimport re\nimport glob\nimport math\nimport time\nimport threading\nfrom datetime import datetime\nfrom flask import Flask, jsonify, request, send_from_directory\nfrom jijin_system import run_clean_list, run_crawler\n\napp = Flask(__name__, static_folder=\'static\', static_url_path=\'\')\n\nDATA_DIR = os.path.join(os.path.dirname(__file__), \'..\', \'fund_data\')\nEXCEL_DIR = os.path.join(os.path.dirname(__file__), \'..\', \'fund_excel\')\n\n# ── 数据缓存 ──────────────────────────────────────────────\n_fund_cache = None\n_cache_time = None\nCACHE_TTL = 600  # 10分钟\n\nUPDATE_STATE = {\n    \'status\': \'idle\',\n    \'message\': \'等待更新\',\n    \'start_time\': None,\n    \'end_time\': None,\n}\nUPDATE_LOGS = []\nUPDATE_LOCK = threading.Lock()\n\n\ndef get_latest_data_file():\n    """获取最新的 fund_profile JSON 文件"""\n    pattern = os.path.join(DATA_DIR, \'fund_profile_*.json\')\n    files = glob.glob(pattern)\n    if not files:\n        return None\n    files.sort(reverse=True)\n    return files[0]\n\n\ndef load_funds(force=False):\n    """加载基金数据（带缓存）"""\n    global _fund_cache, _cache_time\n    now = datetime.now()\n    if (not force and _fund_cache is not None and _cache_time is not None\n            and (now - _cache_time).total_seconds() < CACHE_TTL):\n        return _fund_cache\n\n    filepath = get_latest_data_file()\n    if not filepath:\n        return []\n\n    try:\n        with open(filepath, \'r\', encoding=\'utf-8\') as f:\n            data = json.load(f)\n    except Exception:\n        return []\n\n    _fund_cache = data\n    _cache_time = now\n    return data\n\n\ndef _log_update(message):\n    timestamp = datetime.now().strftime(\'%H:%M:%S\')\n    log_entry = f\'[{timestamp}] {message}\'\n    UPDATE_LOGS.append(log_entry)\n    UPDATE_STATE[\'message\'] = message\n\n\ndef _set_update_state(status, message=\'\'):\n    with UPDATE_LOCK:\n        UPDATE_STATE[\'status\'] = status\n        UPDATE_STATE[\'message\'] = message or UPDATE_STATE.get(\'message\', \'\')\n        if status == \'running\':\n            UPDATE_STATE[\'start_time\'] = datetime.now().strftime(\'%Y-%m-%d %H:%M:%S\')\n            UPDATE_STATE[\'end_time\'] = None\n        else:\n            UPDATE_STATE[\'end_time\'] = datetime.now().strftime(\'%Y-%m-%d %H:%M:%S\')\n\n\ndef _run_update_pipeline():\n    _set_update_state(\'running\', \'正在执行数据更新...\')\n    _log_update(\'开始执行数据更新流水线。\')\n    try:\n        _log_update(\'执行智能清洗，生成或更新 target_funds.json。\')\n        run_clean_list(_log_update)\n        _log_update(\'智能清洗完成，开始爬取最新行情数据。\')\n        run_crawler(log=_log_update, on_progress=None, on_done=None)\n        _log_update(\'爬取完成，刷新缓存数据。\')\n        load_funds(force=True)\n        _set_update_state(\'completed\', \'数据更新完成，最新数据已就绪。\')\n        _log_update(\'数据更新完成。\')\n    except Exception as e:\n        import traceback\n        _set_update_state(\'error\', f\'更新失败：{e}\')\n        _log_update(f\'更新出错：{e}\')\n        _log_update(traceback.format_exc())\n\n\n@app.route(\'/api/update\', methods=[\'POST\'])\ndef api_update():\n    with UPDATE_LOCK:\n        if UPDATE_STATE[\'status\'] == \'running\':\n            return jsonify({\'status\': \'running\', \'message\': \'更新任务正在进行中，请稍后\'}), 409\n        UPDATE_LOGS.clear()\n        updater = threading.Thread(target=_run_update_pipeline, daemon=True)\n        updater.start()\n        return jsonify({\'status\': \'started\', \'message\': \'更新任务已启动，请稍后查看状态。\'}), 202\n\n\n@app.route(\'/api/update_status\')\ndef api_update_status():\n    with UPDATE_LOCK:\n        return jsonify(UPDATE_STATE)\n\n\n@app.route(\'/api/update_log\')\ndef api_update_log():\n    return jsonify({\'logs\': UPDATE_LOGS[-200:]})\n\n\ndef safe_float(val, default=0.0):\n    """安全转换为 float，自动去除 % 和逗号"""\n    if val is None or val == \'\' or val == \'--\':\n        return default\n    if isinstance(val, (int, float)):\n        return float(val)\n    try:\n        # 去除百分号、逗号、空格等\n        cleaned = str(val).replace(\'%\', \'\').replace(\',\', \'\').replace(\'，\', \'\').strip()\n        return float(cleaned)\n    except (ValueError, TypeError):\n        return default\n\n\ndef safe_str(val, default=\'--\'):\n    if val is None:\n        return default\n    return str(val)\n\n\ndef clean_html(raw):\n    """清洗 HTML 标签，提取纯文本"""\n    if raw is None:\n        return \'--\'\n    text = re.sub(r\'<[^>]+>\', \'\', str(raw))\n    text = text.strip()\n    return text if text else \'--\'\n\n\ndef extract_subscription_info(status_data):\n    """从 buy_limit_full 提取限购信息"""\n    raw = status_data.get(\'buy_limit_full\', \'\')\n    if not raw:\n        return \'开放申购\', \'--\'\n    text = clean_html(raw)\n\n    # 提取状态\n    if \'暂停申购\' in text:\n        status_text = \'暂停申购\'\n    elif \'限大额\' in text:\n        status_text = \'限大额\'\n    elif \'开放申购\' in text:\n        status_text = \'开放申购\'\n    else:\n        status_text = text[:20] if len(text) > 20 else text\n\n    # 提取上限金额\n    limit_amount = \'--\'\n    m = re.search(r\'单日累计购买上限\\s*([\\d.,]+)\\s*元\', text)\n    if m:\n        limit_amount = f\'¥{m.group(1)}/日\'\n    else:\n        m = re.search(r\'单日累计购买上限\\s*([\\d.,]+)\\s*美元\', text)\n        if m:\n            limit_amount = f\'${m.group(1)}/日\'\n\n    return status_text, limit_amount\n\n\ndef extract_fund_summary(fund):\n    """提取基金摘要信息"""\n    perf = fund.get(\'performance\', {})\n    base = fund.get(\'base_info\', {})\n    status = fund.get(\'status\', {})\n\n    # 解析限购信息\n    buy_status_text, buy_limit_amount = extract_subscription_info(status)\n\n    return {\n        \'fund_code\': fund.get(\'fund_code\', \'\'),\n        \'fund_name\': fund.get(\'fund_name\', \'\'),\n        \'fund_type\': safe_str(base.get(\'fund_type\')),\n        \'risk_level\': safe_str(base.get(\'risk_level\')),\n        \'assets_size\': safe_str(base.get(\'assets_size\')),\n        \'manager\': safe_str(base.get(\'manager\')),\n        \'company\': safe_str(base.get(\'company\')),\n        \'setup_date\': safe_str(base.get(\'setup_date\')),\n        \'nav\': safe_float(perf.get(\'nav\')),\n        \'nav_date\': safe_str(perf.get(\'nav_date\')),\n        \'daily_growth\': safe_float(perf.get(\'daily_growth_rate\')),\n        \'return_1m\': safe_float(perf.get(\'1m\')),\n        \'return_3m\': safe_float(perf.get(\'3m\')),\n        \'return_6m\': safe_float(perf.get(\'6m\')),\n        \'return_1y\': safe_float(perf.get(\'1y\')),\n        \'return_3y\': safe_float(perf.get(\'3y\')),\n        \'return_since\': safe_float(perf.get(\'since\')),\n        \'sell_status\': safe_str(status.get(\'sell_status\')),\n        \'buy_status\': buy_status_text,\n        \'buy_limit\': buy_limit_amount,\n        \'buy_fee\': safe_str(status.get(\'buy_fee\')),\n    }\n\n\ndef extract_fund_detail(fund):\n    """提取基金详细信息"""\n    perf = fund.get(\'performance\', {})\n    base = fund.get(\'base_info\', {})\n    status = fund.get(\'status\', {})\n    nav_history = fund.get(\'nav_history\', [])\n\n    # 处理净值历史\n    nav_data = []\n    for item in nav_history:\n        try:\n            nav_data.append({\n                \'date\': item.get(\'date\', \'\'),\n                \'val\': safe_float(item.get(\'val\'))\n            })\n        except Exception:\n            continue\n\n    return {\n        \'fund_code\': fund.get(\'fund_code\', \'\'),\n        \'fund_name\': fund.get(\'fund_name\', \'\'),\n        \'extract_time\': fund.get(\'extract_time\', \'\'),\n        \'base_info\': {\n            \'fund_type\': safe_str(base.get(\'fund_type\')),\n            \'risk_level\': safe_str(base.get(\'risk_level\')),\n            \'assets_size\': safe_str(base.get(\'assets_size\')),\n            \'assets_date\': safe_str(base.get(\'assets_date\')),\n            \'manager\': safe_str(base.get(\'manager\')),\n            \'company\': safe_str(base.get(\'company\')),\n            \'setup_date\': safe_str(base.get(\'setup_date\')),\n        },\n        \'performance\': {\n            \'nav\': safe_float(perf.get(\'nav\')),\n            \'nav_date\': safe_str(perf.get(\'nav_date\')),\n            \'daily_growth_rate\': safe_float(perf.get(\'daily_growth_rate\')),\n            \'return_1m\': safe_float(perf.get(\'1m\')),\n            \'return_3m\': safe_float(perf.get(\'3m\')),\n            \'return_6m\': safe_float(perf.get(\'6m\')),\n            \'return_1y\': safe_float(perf.get(\'1y\')),\n            \'return_3y\': safe_float(perf.get(\'3y\')),\n            \'return_since\': safe_float(perf.get(\'since\')),\n        },\n        \'status\': {\n            \'buy_status\': safe_str(status.get(\'buy_status\')),\n            \'sell_status\': safe_str(status.get(\'sell_status\')),\n            \'buy_fee\': safe_str(status.get(\'buy_fee\')),\n        },\n        \'nav_history\': nav_data,\n    }\n\n\n# ── 主题分类映射 ──────────────────────────────────────────\nTHEME_MAP = {\n    \'科创板\': [\'科创\'],\n    \'创业板\': [\'创业板\'],\n    \'沪深300\': [\'沪深300\'],\n    \'中证500\': [\'中证500\'],\n    \'上证50\': [\'上证50\'],\n    \'纳斯达克\': [\'纳斯达克\', \'纳指\', \'纳斯达\'],\n    \'标普500\': [\'标普500\', \'标普 500\', \'SP500\'],\n    \'恒生科技\': [\'恒生科技\'],\n    \'恒生指数\': [\'恒生\'],\n    \'半导体\': [\'半导体\', \'芯片\'],\n    \'人工智能\': [\'人工智能\', \'AI\'],\n    \'新能源\': [\'新能源\', \'光伏\', \'锂电\', \'风电\', \'储能\'],\n    \'医药\': [\'医药\', \'医疗\', \'生物\', \'制药\', \'中药\'],\n    \'消费\': [\'消费\', \'食品\', \'饮料\', \'白酒\'],\n    \'科技\': [\'科技\', \'TMT\', \'信息\'],\n    \'军工\': [\'军工\', \'国防\'],\n    \'红利\': [\'红利\', \'股息\'],\n    \'债券\': [\'债券\', \'债\', \'纯债\', \'短债\', \'中短债\'],\n    \'可转债\': [\'可转债\', \'可转\'],\n    \'FOF\': [\'FOF\', \'fof\'],\n    \'REITs\': [\'REIT\', \'reit\'],\n    \'原油\': [\'原油\', \'石油\'],\n    \'黄金\': [\'黄金\', \'贵金属\'],\n    \'机器人\': [\'机器人\'],\n    \'低波\': [\'低波\', \'稳健\', \'固收\'],\n}\n\n\ndef classify_theme(name):\n    """根据基金名称分类主题"""\n    results = []\n    for theme, keywords in THEME_MAP.items():\n        for kw in keywords:\n            if kw.lower() in name.lower():\n                results.append(theme)\n                break\n    return results\n\n\n# ── API 路由 ──────────────────────────────────────────────\n\n@app.route(\'/\')\ndef index():\n    return send_from_directory(\'static\', \'index.html\')\n\n\n@app.route(\'/api/funds\')\ndef api_funds():\n    """获取基金列表（支持分页、排序、筛选）"""\n    funds = load_funds()\n    page = request.args.get(\'page\', 1, type=int)\n    per_page = request.args.get(\'per_page\', 50, type=int)\n    sort_by = request.args.get(\'sort_by\', \'return_1y\')\n    sort_order = request.args.get(\'sort_order\', \'desc\')\n    search = request.args.get(\'search\', \'\').strip()\n    fund_type = request.args.get(\'fund_type\', \'\').strip()\n    theme = request.args.get(\'theme\', \'\').strip()\n\n    # 提取摘要\n    summaries = [extract_fund_summary(f) for f in funds]\n\n    # 筛选\n    if search:\n        summaries = [s for s in summaries if search.lower() in s[\'fund_name\'].lower()\n                     or search in s[\'fund_code\']]\n    if fund_type:\n        summaries = [s for s in summaries if fund_type in s[\'fund_type\']]\n    if theme:\n        summaries = [s for s in summaries if theme in classify_theme(s[\'fund_name\'])]\n\n    total = len(summaries)\n\n    # 排序\n    reverse = sort_order == \'desc\'\n    try:\n        summaries.sort(key=lambda x: x.get(sort_by, 0) if isinstance(x.get(sort_by), (int, float)) else 0, reverse=reverse)\n    except Exception:\n        pass\n\n    # 分页\n    start = (page - 1) * per_page\n    end = start + per_page\n    page_data = summaries[start:end]\n\n    return jsonify({\n        \'total\': total,\n        \'page\': page,\n        \'per_page\': per_page,\n        \'total_pages\': math.ceil(total / per_page) if per_page > 0 else 0,\n        \'data\': page_data\n    })\n\n\n@app.route(\'/api/fund/<code>\')\ndef api_fund_detail(code):\n    """获取基金详情"""\n    funds = load_funds()\n    for f in funds:\n        if f.get(\'fund_code\') == code:\n            return jsonify({\'success\': True, \'data\': extract_fund_detail(f)})\n    return jsonify({\'success\': False, \'message\': f\'基金 {code} 未找到\'}), 404\n\n\n@app.route(\'/api/themes\')\ndef api_themes():\n    """获取所有主题及其基金数量"""\n    funds = load_funds()\n    theme_counts = {}\n    for f in funds:\n        name = f.get(\'fund_name\', \'\')\n        themes = classify_theme(name)\n        for t in themes:\n            theme_counts[t] = theme_counts.get(t, 0) + 1\n\n    result = [{\'name\': k, \'count\': v} for k, v in sorted(theme_counts.items(), key=lambda x: -x[1])]\n    return jsonify(result)\n\n\n@app.route(\'/api/ranking\')\ndef api_ranking():\n    """获取热门排行"""\n    funds = load_funds()\n    sort_by = request.args.get(\'sort_by\', \'return_1y\')\n    limit = request.args.get(\'limit\', 100, type=int)\n    fund_type = request.args.get(\'fund_type\', \'\').strip()\n\n    summaries = [extract_fund_summary(f) for f in funds]\n\n    # 筛选有效数据\n    summaries = [s for s in summaries if s.get(sort_by, 0) != 0]\n\n    if fund_type:\n        summaries = [s for s in summaries if fund_type in s[\'fund_type\']]\n\n    # 排序\n    summaries.sort(key=lambda x: x.get(sort_by, 0), reverse=True)\n    top = summaries[:limit]\n\n    return jsonify({\'data\': top, \'sort_by\': sort_by})\n\n\n@app.route(\'/api/stats\')\ndef api_stats():\n    """获取统计数据"""\n    funds = load_funds()\n    total = len(funds)\n\n    fund_types = {}\n    companies = {}\n    risk_levels = {}\n    total_assets = 0\n    assets_count = 0\n\n    for f in funds:\n        base = f.get(\'base_info\', {})\n        ft = safe_str(base.get(\'fund_type\'))\n        fund_types[ft] = fund_types.get(ft, 0) + 1\n\n        company = safe_str(base.get(\'company\'))\n        if company != \'--\':\n            companies[company] = companies.get(company, 0) + 1\n\n        risk = safe_str(base.get(\'risk_level\'))\n        risk_levels[risk] = risk_levels.get(risk, 0) + 1\n\n    return jsonify({\n        \'total\': total,\n        \'fund_types\': dict(sorted(fund_types.items(), key=lambda x: -x[1])[:20]),\n        \'top_companies\': dict(sorted(companies.items(), key=lambda x: -x[1])[:20]),\n        \'risk_levels\': risk_levels,\n    })\n\n\n@app.route(\'/api/search\')\ndef api_search():\n    """快速搜索"""\n    funds = load_funds()\n    q = request.args.get(\'q\', \'\').strip()\n    if not q or len(q) < 1:\n        return jsonify([])\n\n    results = []\n    for f in funds:\n        name = f.get(\'fund_name\', \'\')\n        code = f.get(\'fund_code\', \'\')\n        if q.lower() in name.lower() or q in code:\n            results.append({\n                \'fund_code\': code,\n                \'fund_name\': name,\n                \'fund_type\': safe_str(f.get(\'base_info\', {}).get(\'fund_type\')),\n            })\n            if len(results) >= 20:\n                break\n\n    return jsonify(results)\n\n\nif __name__ == \'__main__\':\n    # 预加载数据\n    print("正在加载数据...")\n    load_funds(force=True)\n    print("数据加载完成，启动服务器...")\n    app.run(host=\'0.0.0.0\', port=5000, debug=True)',
    'static/index.html': '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>基金数据可视化系统 - 同花顺专业版</title>\n<style>\n:root {\n  --bg-primary: #0a0e14;\n  --bg-secondary: #12161e;\n  --bg-card: #181c25;\n  --bg-hover: #1e2330;\n  --border: #2a3040;\n  --border-light: #333a48;\n  --text-primary: #dce1e8;\n  --text-secondary: #8b95a8;\n  --text-muted: #5a6378;\n  --red: #e8553d;\n  --red-bg: rgba(232,85,61,0.12);\n  --green: #78b89a;\n  --green-bg: rgba(51,176,124,0.12);\n  --orange: #d4a233;\n  --blue: #4d94ff;\n  --purple: #9b7ef5;\n  --gold: #e8c547;\n  --cyan: #3cc6c6;\n  --max-red: #e8553d;\n  --mid-red: #f08070;\n  --neutral: #8b95a8;\n  --mid-green: #50c898;\n  --max-green: #78b89a;\n}\n\n* { margin:0; padding:0; box-sizing:border-box; }\n\nbody {\n  font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', \'PingFang SC\', \'Microsoft YaHei\', sans-serif;\n  background: var(--bg-primary);\n  color: var(--text-primary);\n  min-height: 100vh;\n  overflow-x: hidden;\n}\n\n/* ── 顶部导航 ────────────────────────── */\n.header {\n  background: var(--bg-secondary);\n  border-bottom: 1px solid var(--border);\n  padding: 0 16px;\n  height: 52px;\n  display: flex;\n  align-items: center;\n  justify-content: space-between;\n  position: sticky;\n  top: 0;\n  z-index: 100;\n}\n.header-left { display:flex; align-items:center; gap:20px; }\n.logo { \n  font-size:18px; font-weight:700; color:var(--text-primary); letter-spacing:1px;\n  display:flex; align-items:center; gap:6px;\n}\n.logo-icon { width:28px; height:28px; background:var(--red); border-radius:6px; \n  display:flex; align-items:center; justify-content:center; font-size:14px; color:#fff; }\n.nav-links { display:flex; gap:2px; }\n.nav-link {\n  padding:6px 14px; border-radius:4px; cursor:pointer; font-size:13px;\n  color:var(--text-secondary); transition:all 0.2s; border:none; background:none;\n  font-weight:500;\n}\n.nav-link:hover { color:var(--text-primary); background:var(--bg-hover); }\n.nav-link.active { color:#fff; background:var(--red); font-weight:600; }\n.header-right { display:flex; align-items:center; gap:10px; }\n.search-box { position:relative; }\n.search-box input {\n  background:var(--bg-primary); border:1px solid var(--border);\n  border-radius:4px; padding:6px 10px 6px 34px; color:var(--text-primary);\n  font-size:13px; width:240px; outline:none; transition:border-color 0.2s;\n}\n.search-box input:focus { border-color:var(--blue); }\n.search-box .search-icon {\n  position:absolute; left:10px; top:50%; transform:translateY(-50%);\n  color:var(--text-muted); font-size:14px;\n}\n.search-results {\n  position:absolute; top:100%; left:0; right:0;\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:0 0 6px 6px; max-height:340px; overflow-y:auto;\n  display:none; z-index:200; box-shadow:0 8px 30px rgba(0,0,0,0.5);\n}\n.search-results.show { display:block; }\n.search-result-item {\n  padding:8px 12px; cursor:pointer; font-size:12px;\n  border-bottom:1px solid var(--border);\n  display:flex; justify-content:space-between; align-items:center;\n}\n.search-result-item:hover { background:var(--bg-hover); }\n.search-result-item .stype { color:var(--text-muted); font-size:11px; }\n\n/* ── Ticker滚动条 ────────────────────── */\n.ticker-bar {\n  background:var(--bg-secondary); border-bottom:1px solid var(--border);\n  padding:5px 16px; display:flex; align-items:center; gap:20px;\n  font-size:11px; overflow:hidden; white-space:nowrap; height:30px;\n}\n.ticker-scroll {\n  display:flex; gap:20px; animation:tickerScroll 60s linear infinite;\n}\n@keyframes tickerScroll {\n  0% { transform:translateX(0); }\n  100% { transform:translateX(-50%); }\n}\n.ticker-item { display:flex; align-items:center; gap:5px; flex-shrink:0; }\n.ticker-item .tname { color:var(--text-secondary); }\n.ticker-item .tval { font-weight:600; }\n.ticker-item .up { color:var(--red); }\n.ticker-item .down { color:var(--green); }\n\n/* ── 主布局 ──────────────────────────── */\n.main-container { display:flex; height:calc(100vh - 82px); }\n.sidebar {\n  width:200px; background:var(--bg-secondary); border-right:1px solid var(--border);\n  padding:12px 0; overflow-y:auto; flex-shrink:0;\n}\n.content { flex:1; overflow-y:auto; padding:12px 16px; }\n\n/* ── 侧边栏 ──────────────────────────── */\n.sidebar-section { padding:0 12px 8px; }\n.sidebar-title {\n  font-size:11px; color:var(--text-muted); text-transform:uppercase;\n  letter-spacing:1.5px; padding:8px 4px 6px; font-weight:700;\n}\n.theme-list { list-style:none; }\n.theme-item {\n  padding:6px 12px; cursor:pointer; font-size:12px; color:var(--text-secondary);\n  display:flex; justify-content:space-between; align-items:center;\n  transition:all 0.12s; border-radius:3px; margin:1px 0;\n}\n.theme-item:hover { background:var(--bg-hover); color:var(--text-primary); }\n.theme-item.active { background:var(--red-bg); color:var(--red); font-weight:600; }\n.theme-item .count { font-size:10px; color:var(--text-muted); background:var(--bg-primary);\n  padding:1px 7px; border-radius:8px; }\n\n/* ── Dashboard ───────────────────────── */\n.dashboard-grid {\n  display:grid;\n  grid-template-columns:2fr 1fr;\n  gap:12px;\n  margin-bottom:12px;\n}\n.heatmap-card {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:12px;\n}\n.heatmap-card h3 { font-size:13px; color:var(--text-secondary); margin-bottom:10px; font-weight:500; }\n.heatmap-grid { display:flex; flex-wrap:wrap; gap:3px; }\n.heatmap-cell {\n  width:18px; height:18px; border-radius:2px; cursor:pointer;\n  transition:transform 0.15s; position:relative;\n}\n.heatmap-cell:hover { transform:scale(1.8); z-index:10; box-shadow:0 0 8px rgba(0,0,0,0.6); }\n.heatmap-cell .tooltip {\n  display:none; position:absolute; bottom:100%; left:50%;\n  transform:translateX(-50%); background:#000; color:#fff; padding:4px 8px;\n  border-radius:4px; font-size:10px; white-space:nowrap; pointer-events:none;\n  z-index:20;\n}\n.heatmap-cell:hover .tooltip { display:block; }\n\n.top-movers {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:12px;\n}\n.top-movers h3 { font-size:13px; color:var(--text-secondary); margin-bottom:10px; font-weight:500; }\n.mover-tabs { display:flex; gap:4px; margin-bottom:8px; }\n.mover-tab {\n  padding:4px 12px; border-radius:3px; font-size:11px; cursor:pointer;\n  background:var(--bg-primary); color:var(--text-secondary); border:none;\n  transition:all 0.15s;\n}\n.mover-tab.active { background:var(--red); color:#fff; }\n.mover-list { max-height:200px; overflow-y:auto; }\n.mover-row {\n  display:flex; align-items:center; padding:5px 0; font-size:11px;\n  cursor:pointer; gap:6px; border-bottom:1px solid var(--border);\n}\n.mover-row:hover { background:var(--bg-hover); }\n.mover-row .mr { font-weight:700; width:18px; color:var(--text-muted); }\n.mover-row .mn { flex:1; color:var(--blue); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }\n.mover-row .mv { font-weight:600; }\n\n/* ── 统计卡片 ────────────────────────── */\n.stats-row {\n  display:grid;\n  grid-template-columns:repeat(4,1fr);\n  gap:10px;\n  margin-bottom:12px;\n}\n.stat-card {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:14px; cursor:pointer; transition:all 0.2s;\n}\n.stat-card:hover { border-color:var(--text-muted); }\n.stat-card .label { font-size:11px; color:var(--text-muted); margin-bottom:4px; }\n.stat-card .value { font-size:22px; font-weight:700; }\n.stat-card .sub { font-size:10px; color:var(--text-muted); margin-top:2px; }\n\n/* ── 工具栏 ──────────────────────────── */\n.toolbar {\n  display:flex; align-items:center; justify-content:space-between;\n  margin-bottom:10px; gap:10px; flex-wrap:wrap;\n}\n.toolbar-group { display:flex; align-items:center; gap:6px; }\n.toolbar select, .toolbar button {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:4px; padding:5px 10px; color:var(--text-secondary);\n  font-size:12px; cursor:pointer; outline:none; transition:all 0.12s;\n}\n.toolbar select:hover, .toolbar button:hover { border-color:var(--text-muted); }\n.toolbar button.primary { background:var(--red); color:#fff; border-color:var(--red); }\n.toolbar button.primary:hover { opacity:0.85; }\n.toolbar button.outline { background:transparent; border-color:var(--blue); color:var(--blue); }\n\n/* ── 数据表格 ────────────────────────── */\n.table-container {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; overflow:hidden;\n}\ntable { width:100%; border-collapse:collapse; font-size:12px; }\nthead th {\n  background:var(--bg-secondary); padding:8px 10px; text-align:right;\n  color:var(--text-secondary); font-weight:600; white-space:nowrap;\n  border-bottom:2px solid var(--border); cursor:pointer; user-select:none;\n  position:sticky; top:0; z-index:10;\n}\nthead th:first-child, thead th:nth-child(2), thead th:nth-child(3) { text-align:left; }\nthead th:hover { color:var(--text-primary); }\nthead th .sort-arrow { font-size:10px; margin-left:2px; opacity:0; }\nthead th:hover .sort-arrow { opacity:1; }\nthead th .sort-arrow.asc { color:var(--red); opacity:1; }\nthead th .sort-arrow.desc { color:var(--green); opacity:1; }\ntbody td {\n  padding:6px 10px; border-bottom:1px solid var(--border);\n  white-space:nowrap; text-align:right;\n}\ntbody td:first-child, tbody td:nth-child(2), tbody td:nth-child(3) { text-align:left; }\ntbody tr { transition:background 0.12s; cursor:pointer; }\ntbody tr:hover { background:var(--bg-hover); }\ntbody tr.selected { background:rgba(77,148,255,0.08); }\n.fund-name { color:var(--blue); font-weight:500; cursor:pointer; }\n.fund-code { color:var(--text-muted); font-size:10px; }\n.up { color:var(--red); }\n.down { color:var(--green); }\n.flat { color:var(--text-secondary); }\n.rank-badge {\n  display:inline-flex; align-items:center; justify-content:center;\n  width:22px; height:22px; border-radius:3px; font-size:11px; font-weight:700; color:#fff;\n}\n.rank-1, .rank-2, .rank-3 { color:#fff; }\n.rank-1 { background:var(--red); }\n.rank-2 { background:var(--orange); }\n.rank-3 { background:var(--gold); }\n.rank-other { color:var(--text-muted); }\n\n/* ── 迷你走势图 ──────────────────────── */\n.sparkline { width:80px; height:30px; vertical-align:middle; }\n\n/* ── 分页 ────────────────────────────── */\n.pagination {\n  display:flex; align-items:center; justify-content:center;\n  gap:4px; padding:14px 0;\n}\n.pagination button {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:4px; padding:5px 10px; color:var(--text-secondary);\n  cursor:pointer; font-size:12px; transition:all 0.12s;\n}\n.pagination button:hover:not(:disabled) { background:var(--bg-hover); }\n.pagination button:disabled { opacity:0.3; cursor:not-allowed; }\n.pagination button.active { background:var(--red); color:#fff; border-color:var(--red); }\n.pagination .page-info { color:var(--text-muted); font-size:11px; margin:0 6px; }\n\n/* ── 详情弹窗 ────────────────────────── */\n.modal-overlay {\n  position:fixed; top:0; left:0; right:0; bottom:0;\n  background:rgba(0,0,0,0.65); z-index:1000; display:none;\n  align-items:center; justify-content:center; backdrop-filter:blur(4px);\n}\n.modal-overlay.show { display:flex; }\n.modal {\n  background:var(--bg-secondary); border:1px solid var(--border);\n  border-radius:10px; width:92%; max-width:1000px; max-height:88vh;\n  overflow-y:auto; box-shadow:0 20px 60px rgba(0,0,0,0.6);\n}\n.modal-header {\n  display:flex; align-items:center; justify-content:space-between;\n  padding:16px 20px; border-bottom:1px solid var(--border);\n  position:sticky; top:0; background:var(--bg-secondary); z-index:10;\n}\n.modal-header h2 { font-size:16px; }\n.modal-close {\n  background:none; border:1px solid var(--border); border-radius:4px;\n  color:var(--text-secondary); font-size:16px; cursor:pointer;\n  width:28px; height:28px; display:flex; align-items:center; justify-content:center;\n  transition:all 0.12s;\n}\n.modal-close:hover { background:var(--bg-hover); color:var(--text-primary); }\n.modal-body { padding:20px; }\n\n/* ── 详情网格 ────────────────────────── */\n.detail-top {\n  display:flex; align-items:flex-start; gap:16px; margin-bottom:16px;\n}\n.detail-price-box {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:14px 20px; text-align:center; min-width:160px;\n}\n.detail-price-box .nav-val { font-size:28px; font-weight:700; }\n.detail-price-box .nav-chg { font-size:13px; margin-top:4px; }\n.detail-price-box .nav-date { font-size:10px; color:var(--text-muted); margin-top:2px; }\n.detail-info-grid {\n  display:grid; grid-template-columns:repeat(3,1fr); gap:8px; flex:1;\n}\n.detail-info-item {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:4px; padding:8px 12px; text-align:center;\n}\n.detail-info-item .di-label { font-size:10px; color:var(--text-muted); }\n.detail-info-item .di-val { font-size:13px; font-weight:600; margin-top:2px; }\n\n.detail-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; }\n.detail-section {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:14px;\n}\n.detail-section h3 { font-size:13px; color:var(--text-secondary); margin-bottom:10px;\n  padding-bottom:8px; border-bottom:1px solid var(--border); font-weight:500; }\n.detail-row { display:flex; justify-content:space-between; padding:5px 0;\n  font-size:12px; color:var(--text-secondary); }\n.detail-row .detail-val { color:var(--text-primary); font-weight:500; }\n.detail-full { grid-column:1/-1; }\n.chart-controls {\n  display:flex; align-items:center; gap:8px; margin-bottom:10px;\n}\n.chart-controls button {\n  background:var(--bg-primary); border:1px solid var(--border);\n  border-radius:3px; padding:3px 10px; color:var(--text-secondary);\n  font-size:11px; cursor:pointer; transition:all 0.12s;\n}\n.chart-controls button.active { background:var(--blue); color:#fff; border-color:var(--blue); }\n.chart-controls button:hover { border-color:var(--text-muted); }\n.chart-container { position:relative; }\ncanvas { width:100% !important; border-radius:4px; }\n.ma-legend { display:flex; gap:14px; margin-top:8px; font-size:10px; color:var(--text-muted); }\n.ma-legend span { display:flex; align-items:center; gap:4px; }\n.ma-dot { width:8px; height:2px; border-radius:1px; }\n\n.return-bars { display:flex; gap:6px; flex-wrap:wrap; margin-top:6px; }\n.return-bar-item {\n  flex:1; min-width:70px; text-align:center;\n  background:var(--bg-secondary); border-radius:4px; padding:8px 6px;\n  border:1px solid var(--border);\n}\n.return-bar-item .period { font-size:10px; color:var(--text-muted); }\n.return-bar-item .val { font-size:15px; font-weight:700; margin-top:3px; }\n\n/* ── 比较面板 ────────────────────────── */\n.compare-bar {\n  background:var(--bg-secondary); border-bottom:1px solid var(--border);\n  padding:6px 16px; display:none; align-items:center; gap:8px;\n  flex-wrap:wrap;\n}\n.compare-bar.show { display:flex; }\n.compare-tag {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:3px; padding:3px 8px; font-size:11px; display:flex;\n  align-items:center; gap:4px;\n}\n.compare-tag .remove { cursor:pointer; color:var(--text-muted); font-weight:700; }\n.compare-tag .remove:hover { color:var(--red); }\n.compare-bar button { font-size:11px; padding:3px 10px; }\n\n/* ── 排行面板 ────────────────────────── */\n.ranking-panel {\n  display:grid;\n  grid-template-columns:repeat(auto-fill, minmax(300px, 1fr));\n  gap:10px; margin-top:10px;\n}\n.ranking-card {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; overflow:hidden;\n}\n.ranking-card .card-header {\n  padding:10px 14px; border-bottom:1px solid var(--border);\n  font-size:13px; font-weight:600; display:flex; align-items:center; gap:6px;\n}\n.ranking-card .card-header .dot { width:6px; height:6px; border-radius:2px; }\n.ranking-card .card-list { padding:6px 0; }\n.ranking-card .rank-row {\n  display:flex; align-items:center; padding:5px 14px;\n  font-size:11px; cursor:pointer; transition:background 0.1s; gap:6px;\n}\n.ranking-card .rank-row:hover { background:var(--bg-hover); }\n.ranking-card .rank-row .rnum { width:18px; font-weight:700; color:var(--text-muted); }\n.ranking-card .rank-row .rname { flex:1; color:var(--blue); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }\n.ranking-card .rank-row .rval { font-weight:600; }\n\n/* ── 主题页卡片 ──────────────────────── */\n.theme-cards {\n  display:grid;\n  grid-template-columns:repeat(auto-fill, minmax(180px, 1fr));\n  gap:10px; margin-top:10px;\n}\n.theme-card-item {\n  background:var(--bg-card); border:1px solid var(--border);\n  border-radius:6px; padding:16px; cursor:pointer; text-align:center;\n  transition:all 0.15s;\n}\n.theme-card-item:hover { border-color:var(--blue); transform:translateY(-2px); }\n.theme-card-item .tc-icon { font-size:28px; margin-bottom:6px; }\n.theme-card-item .tc-name { font-size:13px; font-weight:600; margin-bottom:2px; }\n.theme-card-item .tc-count { font-size:11px; color:var(--text-muted); }\n\n/* ── 响应式 ──────────────────────────── */\n@media (max-width:768px) {\n  .sidebar { display:none; }\n  .stats-row { grid-template-columns:repeat(2,1fr); }\n  .dashboard-grid { grid-template-columns:1fr; }\n  .ranking-panel { grid-template-columns:1fr; }\n  .detail-grid { grid-template-columns:1fr; }\n  .detail-info-grid { grid-template-columns:repeat(2,1fr); }\n  .search-box input { width:140px; }\n  .theme-cards { grid-template-columns:repeat(2,1fr); }\n}\n\n::-webkit-scrollbar { width:5px; height:5px; }\n::-webkit-scrollbar-track { background:var(--bg-primary); }\n::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px; }\n::-webkit-scrollbar-thumb:hover { background:var(--text-muted); }\n</style>\n</head>\n<body>\n\n<!-- 顶部导航 -->\n<div class="header">\n  <div class="header-left">\n    <div class="logo"><div class="logo-icon">K</div>基金Pro</div>\n    <div class="nav-links">\n      <button class="nav-link active" data-tab="market">📊 行情中心</button>\n      <button class="nav-link" data-tab="dashboard">📈 市场概览</button>\n      <button class="nav-link" data-tab="ranking">🏆 收益排行</button>\n      <button class="nav-link" data-tab="themes">📁 主题分类</button>\n    </div>\n  </div>\n  <div class="header-right">\n    <div class="search-box">\n      <span class="search-icon">🔍</span>\n      <input type="text" id="searchInput" placeholder="搜索基金代码/名称...">\n      <div class="search-results" id="searchResults"></div>\n    </div>\n    <button class="nav-link" style="font-size:12px;" id="compareToggle" title="基金对比">⚖️ 对比</button>\n  </div>\n</div>\n\n<!-- 对比栏 -->\n<div class="compare-bar" id="compareBar">\n  <span style="font-size:11px;color:var(--text-muted);">对比基金:</span>\n  <span id="compareTags"></span>\n  <button class="outline" onclick="doCompare()" style="margin-left:8px;">开始对比</button>\n  <button onclick="clearCompare()" style="background:transparent;border:none;color:var(--text-muted);cursor:pointer;font-size:11px;">清空</button>\n</div>\n\n<!-- 行情条 -->\n<div class="ticker-bar">\n  <div class="ticker-scroll" id="tickerScroll"></div>\n</div>\n\n<!-- 主布局 -->\n<div class="main-container">\n  <div class="sidebar" id="sidebar">\n    <div class="sidebar-title">主题分类</div>\n    <ul class="theme-list" id="themeList"></ul>\n  </div>\n\n  <div class="content" id="mainContent">\n    <!-- 行情中心 -->\n    <div id="tab-market">\n      <div class="stats-row" id="statsRow"></div>\n      <div class="toolbar">\n        <div class="toolbar-group">\n          <select id="sortBy">\n            <option value="return_1y">近1年收益</option>\n            <option value="return_6m">近6月收益</option>\n            <option value="return_3m">近3月收益</option>\n            <option value="return_1m">近1月收益</option>\n            <option value="return_3y">近3年收益</option>\n            <option value="return_since">成立以来</option>\n            <option value="daily_growth">日涨幅</option>\n            <option value="nav">单位净值</option>\n          </select>\n          <select id="sortOrder">\n            <option value="desc">降序</option>\n            <option value="asc">升序</option>\n          </select>\n          <select id="filterType"><option value="">全部类型</option></select>\n          <select id="perPage">\n            <option value="20">20条</option>\n            <option value="50" selected>50条</option>\n            <option value="100">100条</option>\n          </select>\n        </div>\n        <div class="toolbar-group">\n          <button class="primary" onclick="location.reload()">🔄 刷新</button>\n          <button class="primary" id="updateButton" onclick="triggerUpdateData()">🆕 更新数据</button>\n          <span id="updateStatus" style="margin-left:12px;color:var(--green);font-size:12px;line-height:32px;">状态：就绪</span>\n        </div>\n      </div>\n      <div class="table-container">\n        <table>\n          <thead>\n            <tr>\n              <th style="width:42px">#</th>\n              <th style="width:72px">代码</th>\n              <th style="min-width:160px">基金名称</th>\n              <th style="width:65px">类型</th>\n              <th>净值</th>\n              <th>日涨幅</th>\n              <th>近1月</th>\n              <th>近3月</th>\n              <th>近6月</th>\n              <th>近1年</th>\n              <th>近3年</th>\n              <th>走势</th>\n              <th>规模</th>\n              <th>经理</th>\n              <th>申购状态</th>\n              <th>限购金额</th>\n              <th>费率</th>\n            </tr>\n          </thead>\n          <tbody id="fundTableBody"></tbody>\n        </table>\n      </div>\n      <div class="pagination" id="pagination"></div>\n    </div>\n\n    <!-- 市场概览 -->\n    <div id="tab-dashboard" style="display:none;">\n      <div class="dashboard-grid">\n        <div class="heatmap-card">\n          <h3>🔥 热门主题基金分布热力图</h3>\n          <div class="heatmap-grid" id="heatmapGrid"></div>\n        </div>\n        <div class="top-movers">\n          <h3>📈 涨跌龙虎榜</h3>\n          <div class="mover-tabs">\n            <button class="mover-tab active" onclick="loadTopMovers(\'daily_growth\')">日涨幅榜</button>\n            <button class="mover-tab" onclick="loadTopMovers(\'return_1m\')">月涨幅榜</button>\n            <button class="mover-tab" onclick="loadTopMovers(\'return_1y\')">年涨幅榜</button>\n          </div>\n          <div class="mover-list" id="moverList"></div>\n        </div>\n      </div>\n      <div class="ranking-panel" id="dashRankings"></div>\n    </div>\n\n    <!-- 收益排行 -->\n    <div id="tab-ranking" style="display:none;">\n      <div class="toolbar">\n        <div class="toolbar-group">\n          <select id="rankPeriod">\n            <option value="return_1y">近1年收益</option>\n            <option value="return_6m">近6月收益</option>\n            <option value="return_3m">近3月收益</option>\n            <option value="return_1m">近1月收益</option>\n            <option value="return_3y">近3年收益</option>\n            <option value="return_since">成立以来</option>\n            <option value="daily_growth">日涨幅</option>\n          </select>\n          <select id="rankLimit">\n            <option value="50">Top 50</option>\n            <option value="100" selected>Top 100</option>\n            <option value="200">Top 200</option>\n          </select>\n          <select id="rankType"><option value="">全部类型</option></select>\n        </div>\n      </div>\n      <div class="ranking-panel" id="rankingPanel"></div>\n    </div>\n\n    <!-- 主题分类 -->\n    <div id="tab-themes" style="display:none;">\n      <div class="toolbar"><h3 style="font-size:14px;">📁 主题基金分类</h3></div>\n      <div class="theme-cards" id="themeCards"></div>\n    </div>\n  </div>\n</div>\n\n<!-- 详情弹窗 -->\n<div class="modal-overlay" id="modalOverlay">\n  <div class="modal" id="modalContent"></div>\n</div>\n\n<script>\n// ── 全局状态 ────────────────────────────\nconst STATE = {\n  currentTab: \'market\', currentSort: \'return_1y\', currentOrder: \'desc\',\n  currentFilter: \'\', currentTheme: \'\', currentPage: 1, perPage: 50,\n  searchQuery: \'\', compareList: [], chartPeriod: \'all\'\n};\n\nconst PERIOD_NAMES = {\n  \'daily_growth\': \'日涨幅\', \'return_1m\': \'近1月\', \'return_3m\': \'近3月\',\n  \'return_6m\': \'近6月\', \'return_1y\': \'近1年\', \'return_3y\': \'近3年\',\n  \'return_since\': \'成立以来\'\n};\n\nfunction fmt(n,d=2){if(n==null||n===\'\'||isNaN(n))return\'--\';return Number(n).toFixed(d);}\nfunction fpct(n){if(n==null||n===\'\'||isNaN(n))return\'--\';const v=Number(n);return(v>0?\'+\':\'\')+v.toFixed(2)+\'%\';}\nfunction pcls(n){if(n==null||isNaN(n)||n===0)return\'flat\';return n>0?\'up\':\'down\';}\n\nasync function api(url){const r=await fetch(url);return r.json();}\n\nasync function updateApi(url, options){\n  const r = await fetch(url, options);\n  const data = await r.json();\n  return {status:r.status, ok:r.ok, data};\n}\n\nfunction setUpdateStatus(text, color=\'var(--green)\'){\n  const status = document.getElementById(\'updateStatus\');\n  if(status){status.textContent = `状态：${text}`; status.style.color = color;}\n}\n\nasync function triggerUpdateData(){\n  const btn = document.getElementById(\'updateButton\');\n  if(!btn) return;\n  btn.disabled = true;\n  btn.textContent = \'⏳ 更新中...\';\n  setUpdateStatus(\'任务已启动\');\n  const resp = await updateApi(\'/api/update\',{method:\'POST\'});\n  if(resp.status === 409){\n    setUpdateStatus(resp.data.message || \'已有更新任务进行中\', \'var(--orange)\');\n    startUpdatePolling();\n    return;\n  }\n  if(!resp.ok){\n    setUpdateStatus(\'启动失败\', \'var(--red)\');\n    btn.disabled = false;\n    btn.textContent = \'🆕 更新数据\';\n    return;\n  }\n  setUpdateStatus(resp.data.message || \'已开始更新\', \'var(--blue)\');\n  startUpdatePolling();\n}\n\nlet _updatePollTimer = null;\nfunction startUpdatePolling(){\n  if(_updatePollTimer){ clearTimeout(_updatePollTimer); }\n  _updatePollTimer = setTimeout(pollUpdateStatus, 1500);\n}\n\nasync function pollUpdateStatus(){\n  const btn = document.getElementById(\'updateButton\');\n  try{\n    const resp = await updateApi(\'/api/update_status\');\n    const data = resp.data;\n    if(resp.ok){\n      const statusText = data.status === \'running\' ? \'更新中...\' :\n                         data.status === \'completed\' ? \'已完成\' :\n                         data.status === \'error\' ? \'异常结束\' :\n                         \'就绪\';\n      const color = data.status === \'running\' ? \'var(--blue)\' :\n                    data.status === \'completed\' ? \'var(--green)\' :\n                    data.status === \'error\' ? \'var(--red)\' :\n                    \'var(--text-secondary)\';\n      setUpdateStatus(`${statusText} ${data.message||\'\'}`.trim(), color);\n      if(data.status === \'running\'){\n        _updatePollTimer = setTimeout(pollUpdateStatus, 2000);\n      } else {\n        if(btn){ btn.disabled = false; btn.textContent = \'🆕 更新数据\'; }\n        if(data.status === \'completed\'){\n          await loadStats();\n          await loadThemes();\n          await loadFundList();\n          await loadTicker();\n        }\n      }\n    }\n  }catch(e){\n    setUpdateStatus(\'状态查询失败\', \'var(--red)\');\n    if(btn){ btn.disabled = false; btn.textContent = \'🆕 更新数据\'; }\n  }\n}\n\n// ── 初始化 ─────────────────────────────\nasync function init(){\n  await loadStats();\n  await loadThemes();\n  await loadFundList();\n  await loadTicker();\n  if(STATE.currentTab===\'dashboard\') loadDashboard();\n}\n\n// ── Ticker ─────────────────────────────\nasync function loadTicker(){\n  const data=await api(\'/api/ranking?sort_by=daily_growth&limit=30\');\n  const ticker=document.getElementById(\'tickerScroll\');\n  const items=data.data.map(r=>{\n    const cls=r.daily_growth>0?\'up\':\'down\';\n    return `<div class="ticker-item"><span class="tname">${r.fund_name}</span><span class="tval ${cls}">${fpct(r.daily_growth)}</span></div>`;\n  });\n  ticker.innerHTML=items.join(\'\')+items.join(\'\');\n}\n\n// ── 统计卡片 ───────────────────────────\nasync function loadStats(){\n  const data=await api(\'/api/stats\');\n  document.getElementById(\'statsRow\').innerHTML=`\n    <div class="stat-card"><div class="label">📊 基金总数</div><div class="value" style="color:var(--blue);">${data.total.toLocaleString()}</div><div class="sub">全市场基金</div></div>\n    <div class="stat-card"><div class="label">🏢 基金公司</div><div class="value" style="color:var(--purple);">${Object.keys(data.top_companies).length}</div><div class="sub">已覆盖基金公司</div></div>\n    <div class="stat-card"><div class="label">📁 基金类型</div><div class="value" style="color:var(--orange);">${Object.keys(data.fund_types).length}</div><div class="sub">分类数量</div></div>\n    <div class="stat-card"><div class="label">🔄 数据更新</div><div class="value" style="color:var(--green);">实时</div><div class="sub">最新行情数据</div></div>`;\n  const fs=document.getElementById(\'filterType\');\n  Object.keys(data.fund_types).forEach(t=>{fs.innerHTML+=`<option value="${t}">${t} (${data.fund_types[t]})</option>`;});\n  const rt=document.getElementById(\'rankType\');\n  Object.keys(data.fund_types).forEach(t=>{rt.innerHTML+=`<option value="${t}">${t}</option>`;});\n}\n\n// ── 主题列表 ───────────────────────────\nasync function loadThemes(){\n  const data=await api(\'/api/themes\');\n  const list=document.getElementById(\'themeList\');\n  list.innerHTML=\'<li class="theme-item active" onclick="selTheme(\\\'\\\')"><span>📋 全部基金</span></li>\';\n  data.forEach(t=>{\n    list.innerHTML+=`<li class="theme-item" onclick="selTheme(\'${t.name}\')"><span>${t.name}</span><span class="count">${t.count}</span></li>`;\n  });\n}\n\nfunction selTheme(t){\n  STATE.currentTheme=t; STATE.currentPage=1;\n  document.querySelectorAll(\'.theme-item\').forEach(e=>e.classList.toggle(\'active\',e.textContent.includes(t)&&t!==\'\'));\n  if(t===\'\') document.querySelector(\'.theme-item\').classList.add(\'active\');\n  loadFundList();\n}\n\n// ── 基金列表 ───────────────────────────\nasync function loadFundList(){\n  const p=new URLSearchParams({\n    page:STATE.currentPage,per_page:STATE.perPage,\n    sort_by:STATE.currentSort,sort_order:STATE.currentOrder,\n    search:STATE.searchQuery,theme:STATE.currentTheme,fund_type:STATE.currentFilter\n  });\n  const data=await api(\'/api/funds?\'+p.toString());\n  document.getElementById(\'tickerScroll\').style.animationPlayState=\'running\';\n  renderTable(data.data,data.page,data.per_page);\n  renderPagination(data.total,data.page,data.total_pages);\n}\n\nfunction renderTable(rows,page,pp){\n  const tb=document.getElementById(\'fundTableBody\');\n  if(!rows.length){tb.innerHTML=\'<tr><td colspan="17" style="text-align:center;padding:40px;color:var(--text-muted);">暂无匹配数据</td></tr>\';return;}\n  tb.innerHTML=rows.map((r,i)=>{\n    const rank=(page-1)*pp+i+1;\n    let rc=\'rank-other\'; if(rank===1)rc=\'rank-1\'; else if(rank===2)rc=\'rank-2\'; else if(rank===3)rc=\'rank-3\';\n    return `<tr onclick="rowClick(event,\'${r.fund_code}\')" data-code="${r.fund_code}">\n      <td><span class="rank-badge ${rc}">${rank}</span></td>\n      <td><span class="fund-code">${r.fund_code}</span></td>\n      <td><span class="fund-name">${r.fund_name}</span></td>\n      <td>${r.fund_type||\'--\'}</td>\n      <td>${fmt(r.nav,4)}</td>\n      <td class="${pcls(r.daily_growth)}">${fpct(r.daily_growth)}</td>\n      <td class="${pcls(r.return_1m)}">${fpct(r.return_1m)}</td>\n      <td class="${pcls(r.return_3m)}">${fpct(r.return_3m)}</td>\n      <td class="${pcls(r.return_6m)}">${fpct(r.return_6m)}</td>\n      <td class="${pcls(r.return_1y)}">${fpct(r.return_1y)}</td>\n      <td class="${pcls(r.return_3y)}">${fpct(r.return_3y)}</td>\n      <td><canvas class="sparkline" id="spark_${r.fund_code}" width="80" height="30"></canvas></td>\n      <td>${r.assets_size||\'--\'}</td>\n      <td>${r.manager||\'--\'}</td>\n      <td>${r.buy_status||\'--\'}</td>\n      <td>${r.buy_limit||\'--\'}</td>\n      <td>${r.buy_fee||\'--\'}</td>\n    </tr>`;\n  }).join(\'\');\n  // 异步加载迷你走势图\n  setTimeout(()=>{\n    rows.forEach(r=>loadSparkline(r.fund_code));\n  },50);\n}\n\nasync function loadSparkline(code){\n  const canvas=document.getElementById(\'spark_\'+code);\n  if(!canvas) return;\n  try{\n    const d=await api(\'/api/fund/\'+code);\n    if(!d.success||!d.data.nav_history) return;\n    drawSpark(canvas,d.data.nav_history);\n  }catch(e){}\n}\n\nfunction drawSpark(canvas,data){\n  const ctx=canvas.getContext(\'2d\');\n  const vals=data.filter(d=>d.val>0).map(d=>d.val);\n  if(vals.length<2) return;\n  const W=80,H=30;\n  canvas.width=W*2; canvas.height=H*2; ctx.scale(2,2);\n  const min=Math.min(...vals),max=Math.max(...vals),range=max-min||1;\n  const first=vals[0],last=vals[vals.length-1];\n  const color=last>=first?\'#e8553d\':\'#78b89a\';\n  ctx.strokeStyle=color; ctx.lineWidth=1.2;\n  ctx.beginPath();\n  vals.forEach((v,i)=>{\n    const x=(i/(vals.length-1))*W;\n    const y=H-((v-min)/range)*H;\n    if(i===0)ctx.moveTo(x,y); else ctx.lineTo(x,y);\n  });\n  ctx.stroke();\n}\n\nfunction rowClick(e,code){\n  if(e.target.closest(\'.fund-name\')){\n    openDetail(code);\n  }else{\n    toggleCompare(code);\n    document.querySelectorAll(\'tr[data-code="\'+code+\'"]\').forEach(r=>r.classList.toggle(\'selected\'));\n  }\n}\n\n// ── 分页 ───────────────────────────────\nfunction renderPagination(total,page,tp){\n  const pag=document.getElementById(\'pagination\');\n  if(tp<=1){pag.innerHTML=\'\';return;}\n  let h=`<button ${page<=1?\'disabled\':\'\'} onclick="goPage(${page-1})">‹ 上一页</button>`;\n  const ms=7; let s=Math.max(1,page-3),e=Math.min(tp,page+3);\n  if(e-s<ms-1){if(s===1)e=Math.min(tp,s+ms-1);else s=Math.max(1,e-ms+1);}\n  if(s>1)h+=`<button onclick="goPage(1)">1</button><span class="page-info">...</span>`;\n  for(let i=s;i<=e;i++)h+=`<button class="${i===page?\'active\':\'\'}" onclick="goPage(${i})">${i}</button>`;\n  if(e<tp)h+=`<span class="page-info">...</span><button onclick="goPage(${tp})">${tp}</button>`;\n  h+=`<button ${page>=tp?\'disabled\':\'\'} onclick="goPage(${page+1})">下一页 ›</button>`;\n  h+=`<span class="page-info">共 ${total.toLocaleString()} 条</span>`;\n  pag.innerHTML=h;\n}\n\nfunction goPage(p){STATE.currentPage=p;loadFundList();document.getElementById(\'mainContent\').scrollTop=0;}\n\n// ── 对比功能 ───────────────────────────\nfunction toggleCompare(code){\n  const idx=STATE.compareList.indexOf(code);\n  if(idx>=0) STATE.compareList.splice(idx,1);\n  else if(STATE.compareList.length<5) STATE.compareList.push(code);\n  updateCompareBar();\n}\n\nfunction updateCompareBar(){\n  const bar=document.getElementById(\'compareBar\');\n  const tags=document.getElementById(\'compareTags\');\n  if(STATE.compareList.length===0){bar.classList.remove(\'show\');return;}\n  bar.classList.add(\'show\');\n  tags.innerHTML=STATE.compareList.map(c=>`<span class="compare-tag">${c}<span class="remove" onclick="toggleCompare(\'${c}\')">✕</span></span>`).join(\'\');\n}\n\nfunction clearCompare(){STATE.compareList=[];updateCompareBar();document.querySelectorAll(\'tr.selected\').forEach(r=>r.classList.remove(\'selected\'));}\n\nasync function doCompare(){\n  if(STATE.compareList.length<2){alert(\'请至少选择2只基金进行对比\');return;}\n  const codes=STATE.compareList.join(\',\');\n  const overlay=document.getElementById(\'modalOverlay\');\n  const content=document.getElementById(\'modalContent\');\n  overlay.classList.add(\'show\');\n  content.innerHTML=\'<div class="modal-body" style="text-align:center;padding:60px;">加载对比数据...</div>\';\n  const results=[];\n  for(const code of STATE.compareList){\n    const resp=await api(\'/api/fund/\'+code);\n    if(resp.success) results.push(resp.data);\n  }\n  if(results.length<2){content.innerHTML=\'<div class="modal-body">数据不足</div>\';return;}\n  const periods=[{k:\'return_1m\',n:\'近1月\'},{k:\'return_3m\',n:\'近3月\'},{k:\'return_6m\',n:\'近6月\'},{k:\'return_1y\',n:\'近1年\'},{k:\'return_3y\',n:\'近3年\'},{k:\'return_since\',n:\'成立以来\'}];\n  content.innerHTML=`\n    <div class="modal-header"><h2>📊 基金对比 <span style="font-size:12px;color:var(--text-muted);">${results.map(r=>r.fund_name).join(\' vs \')}</span></h2><button class="modal-close" onclick="closeDetail()">✕</button></div>\n    <div class="modal-body">\n      <div class="detail-section detail-full"><h3>📈 收益对比</h3>\n        <div style="overflow-x:auto;">\n          <table style="font-size:12px;">\n            <thead><tr><th>指标</th>${results.map(r=>`<th>${r.fund_name}</th>`).join(\'\')}</tr></thead>\n            <tbody>\n              <tr><td>单位净值</td>${results.map(r=>`<td>${fmt(r.performance.nav,4)}</td>`).join(\'\')}</tr>\n              <tr><td>日涨幅</td>${results.map(r=>`<td class="${pcls(r.performance.daily_growth_rate)}">${fpct(r.performance.daily_growth_rate)}</td>`).join(\'\')}</tr>\n              ${periods.map(p=>`<tr><td>${p.n}</td>${results.map(r=>`<td class="${pcls(r.performance[p.k])}">${fpct(r.performance[p.k])}</td>`).join(\'\')}</tr>`).join(\'\')}\n              <tr><td>基金类型</td>${results.map(r=>`<td>${r.base_info.fund_type}</td>`).join(\'\')}</tr>\n              <tr><td>基金规模</td>${results.map(r=>`<td>${r.base_info.assets_size}</td>`).join(\'\')}</tr>\n            </tbody>\n          </table>\n        </div>\n      </div>\n      <div class="detail-section detail-full chart-container" style="margin-top:14px;">\n        <h3>📉 净值走势对比（归一化）</h3>\n        <canvas id="compareChart"></canvas>\n      </div>\n    </div>`;\n  setTimeout(()=>drawCompareChart(results),150);\n}\n\nfunction drawCompareChart(results){\n  const canvas=document.getElementById(\'compareChart\');\n  if(!canvas)return;\n  const ctx=canvas.getContext(\'2d\');\n  const W=canvas.parentElement.clientWidth-32,H=320;\n  canvas.width=W*2; canvas.height=H*2; canvas.style.width=W+\'px\'; canvas.style.height=H+\'px\';\n  ctx.scale(2,2);\n  const pad={top:20,right:30,bottom:40,left:60},cw=W-pad.left-pad.right,ch=H-pad.top-pad.bottom;\n  const colors=[\'#e8553d\',\'#4d94ff\',\'#78b89a\',\'#e8c547\',\'#9b7ef5\'];\n  // 归一化：以每个基金第一个可用净值为1\n  const series=results.map((r,i)=>{\n    const navs=(r.nav_history||[]).filter(d=>d.val>0);\n    if(navs.length<2)return null;\n    const base=navs[0].val;\n    return {color:colors[i%colors.length],name:r.fund_name,data:navs.map(d=>({date:d.date,val:d.val/base}))};\n  }).filter(s=>s!==null);\n  if(series.length<2)return;\n  let allMin=Infinity,allMax=-Infinity;\n  series.forEach(s=>{s.data.forEach(d=>{if(d.val<allMin)allMin=d.val;if(d.val>allMax)allMax=d.val;});});\n  allMin*=0.98; allMax*=1.02; const range=allMax-allMin||1;\n  function x(i,len){return pad.left+(i/(len-1))*cw;}\n  function y(v){return pad.top+ch-((v-allMin)/range)*ch;}\n  // Grid\n  ctx.strokeStyle=\'rgba(42,48,64,0.5)\';ctx.lineWidth=0.5;\n  for(let i=0;i<=5;i++){const gy=pad.top+(i/5)*ch;ctx.beginPath();ctx.moveTo(pad.left,gy);ctx.lineTo(pad.left+cw,gy);ctx.stroke();\n    ctx.fillStyle=\'#8b95a8\';ctx.font=\'10px sans-serif\';ctx.textAlign=\'right\';\n    ctx.fillText((allMin+range*(5-i)/5).toFixed(3),pad.left-6,gy+3);}\n  // Lines\n  series.forEach(s=>{\n    ctx.beginPath();ctx.strokeStyle=s.color;ctx.lineWidth=2;\n    s.data.forEach((d,i)=>{if(i===0)ctx.moveTo(x(i,s.data.length),y(d.val));else ctx.lineTo(x(i,s.data.length),y(d.val));});\n    ctx.stroke();\n  });\n  // Legend\n  series.forEach((s,i)=>{\n    ctx.fillStyle=s.color;ctx.fillRect(pad.left+i*120,pad.top+ch+12,12,3);\n    ctx.fillStyle=\'#8b95a8\';ctx.font=\'11px sans-serif\';ctx.textAlign=\'left\';\n    ctx.fillText(s.name,pad.left+i*120+16,pad.top+ch+16);\n  });\n}\n\n// ── 详情弹窗 ───────────────────────────\nasync function openDetail(code){\n  const overlay=document.getElementById(\'modalOverlay\');\n  const content=document.getElementById(\'modalContent\');\n  overlay.classList.add(\'show\');\n  content.innerHTML=\'<div class="modal-body" style="text-align:center;padding:60px;">加载中...</div>\';\n  const resp=await api(\'/api/fund/\'+code);\n  if(!resp.success){content.innerHTML=`<div class="modal-body" style="text-align:center;padding:60px;">${resp.message}</div>`;return;}\n  const d=resp.data,p=d.performance,b=d.base_info,s=d.status;\n  content.innerHTML=`\n    <div class="modal-header"><h2>${d.fund_name} <span style="font-size:13px;color:var(--text-muted);">${d.fund_code}</span></h2><button class="modal-close" onclick="closeDetail()">✕</button></div>\n    <div class="modal-body">\n      <div class="detail-top">\n        <div class="detail-price-box">\n          <div class="nav-val">${fmt(p.nav,4)}</div>\n          <div class="nav-chg ${pcls(p.daily_growth_rate)}">${fpct(p.daily_growth_rate)}</div>\n          <div class="nav-date">净值日期: ${p.nav_date}</div>\n        </div>\n        <div class="detail-info-grid">\n          <div class="detail-info-item"><div class="di-label">基金类型</div><div class="di-val">${b.fund_type}</div></div>\n          <div class="detail-info-item"><div class="di-label">风险等级</div><div class="di-val">${b.risk_level}</div></div>\n          <div class="detail-info-item"><div class="di-label">基金规模</div><div class="di-val">${b.assets_size}</div></div>\n          <div class="detail-info-item"><div class="di-label">基金经理</div><div class="di-val">${b.manager}</div></div>\n          <div class="detail-info-item"><div class="di-label">基金公司</div><div class="di-val">${b.company}</div></div>\n          <div class="detail-info-item"><div class="di-label">成立日期</div><div class="di-val">${b.setup_date}</div></div>\n        </div>\n      </div>\n      <div class="detail-grid">\n        <div class="detail-section detail-full">\n          <h3>📈 阶段收益</h3>\n          <div class="return-bars">\n            <div class="return-bar-item"><div class="period">近1月</div><div class="val ${pcls(p.return_1m)}">${fpct(p.return_1m)}</div></div>\n            <div class="return-bar-item"><div class="period">近3月</div><div class="val ${pcls(p.return_3m)}">${fpct(p.return_3m)}</div></div>\n            <div class="return-bar-item"><div class="period">近6月</div><div class="val ${pcls(p.return_6m)}">${fpct(p.return_6m)}</div></div>\n            <div class="return-bar-item"><div class="period">近1年</div><div class="val ${pcls(p.return_1y)}">${fpct(p.return_1y)}</div></div>\n            <div class="return-bar-item"><div class="period">近3年</div><div class="val ${pcls(p.return_3y)}">${fpct(p.return_3y)}</div></div>\n            <div class="return-bar-item"><div class="period">成立以来</div><div class="val ${pcls(p.return_since)}">${fpct(p.return_since)}</div></div>\n          </div>\n        </div>\n        <div class="detail-section">\n          <h3>💰 交易信息</h3>\n          <div class="detail-row"><span>申购状态</span><span class="detail-val">${s.buy_status}</span></div>\n          <div class="detail-row"><span>赎回状态</span><span class="detail-val">${s.sell_status}</span></div>\n          <div class="detail-row"><span>申购费率</span><span class="detail-val">${s.buy_fee}</span></div>\n        </div>\n        ${d.nav_history&&d.nav_history.length>0?`\n        <div class="detail-section detail-full chart-container">\n          <h3>📉 K线风格净值走势</h3>\n          <div class="chart-controls">\n            <button class="active" onclick="changeChartPeriod(\'all\',this)">全部</button>\n            <button onclick="changeChartPeriod(\'90\',this)">近3月</button>\n            <button onclick="changeChartPeriod(\'180\',this)">近半年</button>\n            <button onclick="changeChartPeriod(\'365\',this)">近1年</button>\n          </div>\n          <canvas id="navChart"></canvas>\n          <div class="ma-legend">\n            <span><span class="ma-dot" style="background:#e8c547;"></span> MA5</span>\n            <span><span class="ma-dot" style="background:#4d94ff;"></span> MA10</span>\n            <span><span class="ma-dot" style="background:#9b7ef5;"></span> MA20</span>\n            <span><span class="ma-dot" style="background:#3cc6c6;"></span> MA60</span>\n          </div>\n        </div>`:\'\'}\n      </div>\n    </div>`;\n  if(d.nav_history&&d.nav_history.length>0){\n    setTimeout(()=>drawNavChart(d.nav_history),150);\n  }\n}\n\nfunction closeDetail(){document.getElementById(\'modalOverlay\').classList.remove(\'show\');}\nfunction changeChartPeriod(period,btn){\n  STATE.chartPeriod=period;\n  document.querySelectorAll(\'.chart-controls button\').forEach(b=>b.classList.remove(\'active\'));\n  btn.classList.add(\'active\');\n  // 重绘 - 通过重新获取数据\n  const code=document.querySelector(\'.modal-header h2 span\')?.textContent;\n  if(code) openDetail(code);\n}\n\n// ── K线风格净值图 ──────────────────────\nfunction drawNavChart(navHistory){\n  const canvas=document.getElementById(\'navChart\');\n  if(!canvas)return;\n  const ctx=canvas.getContext(\'2d\');\n  const W=canvas.parentElement.clientWidth-32,H=350;\n  canvas.width=W*2;canvas.height=H*2;canvas.style.width=W+\'px\';canvas.style.height=H+\'px\';\n  ctx.scale(2,2);\n  let data=navHistory.filter(d=>d.val>0);\n  if(STATE.chartPeriod!==\'all\'){\n    const days=parseInt(STATE.chartPeriod);\n    const cutoff=new Date(); cutoff.setDate(cutoff.getDate()-days);\n    data=data.filter(d=>new Date(d.date)>=cutoff);\n  }\n  if(data.length<2)return;\n  const vals=data.map(d=>d.val);\n  const minVal=Math.min(...vals)*0.99, maxVal=Math.max(...vals)*1.01, range=maxVal-minVal||1;\n  const pad={top:24,right:24,bottom:44,left:68};\n  const cw=W-pad.left-pad.right,ch=H-pad.top-pad.bottom;\n  function xi(i){return pad.left+(i/(data.length-1))*cw;}\n  function yv(v){return pad.top+ch-((v-minVal)/range)*ch;}\n\n  // Grid\n  const gridLines=6;\n  for(let i=0;i<=gridLines;i++){\n    const gy=pad.top+(i/gridLines)*ch;\n    ctx.strokeStyle=\'rgba(42,48,64,0.5)\';ctx.lineWidth=0.5;\n    ctx.beginPath();ctx.moveTo(pad.left,gy);ctx.lineTo(pad.left+cw,gy);ctx.stroke();\n    const lv=minVal+(range*(gridLines-i)/gridLines);\n    ctx.fillStyle=\'#8b95a8\';ctx.font=\'10px sans-serif\';ctx.textAlign=\'right\';\n    ctx.fillText(lv.toFixed(3),pad.left-8,gy+3);\n  }\n\n  // Date labels\n  const ds=Math.max(1,Math.floor(data.length/6));\n  for(let i=0;i<data.length;i+=ds){\n    ctx.fillStyle=\'#8b95a8\';ctx.font=\'10px sans-serif\';ctx.textAlign=\'center\';\n    ctx.fillText(data[i].date.slice(5),xi(i),H-pad.bottom+16);\n  }\n\n  // MA calculations\n  function calcMA(period){\n    const result=[]; for(let i=0;i<data.length;i++){\n      if(i<period-1){result.push(null);continue;}\n      let sum=0; for(let j=i-period+1;j<=i;j++) sum+=data[j].val;\n      result.push(sum/period);\n    } return result;\n  }\n  const ma5=calcMA(5),ma10=calcMA(10),ma20=calcMA(20),ma60=calcMA(60);\n\n  // Fill area\n  const grad=ctx.createLinearGradient(0,pad.top,0,pad.top+ch);\n  grad.addColorStop(0,\'rgba(232,85,61,0.18)\');grad.addColorStop(1,\'rgba(232,85,61,0.0)\');\n  ctx.beginPath();ctx.moveTo(xi(0),pad.top+ch);\n  for(let i=0;i<data.length;i++)ctx.lineTo(xi(i),yv(data[i].val));\n  ctx.lineTo(xi(data.length-1),pad.top+ch);ctx.closePath();\n  ctx.fillStyle=grad;ctx.fill();\n\n  // Price line\n  ctx.beginPath();ctx.strokeStyle=\'#e8553d\';ctx.lineWidth=1.8;\n  for(let i=0;i<data.length;i++){if(i===0)ctx.moveTo(xi(i),yv(data[i].val));else ctx.lineTo(xi(i),yv(data[i].val));}\n  ctx.stroke();\n\n  // MA lines\n  const mas=[{d:ma5,c:\'#e8c547\'},{d:ma10,c:\'#4d94ff\'},{d:ma20,c:\'#9b7ef5\'},{d:ma60,c:\'#3cc6c6\'}];\n  mas.forEach(ma=>{\n    ctx.beginPath();ctx.strokeStyle=ma.c;ctx.lineWidth=1;ctx.setLineDash([3,3]);\n    let started=false;\n    for(let i=0;i<ma.d.length;i++){\n      if(ma.d[i]===null)continue;\n      if(!started){ctx.moveTo(xi(i),yv(ma.d[i]));started=true;}\n      else ctx.lineTo(xi(i),yv(ma.d[i]));\n    }\n    ctx.stroke();ctx.setLineDash([]);\n  });\n\n  // Last value dot\n  const lx=xi(data.length-1),ly=yv(data[data.length-1].val);\n  ctx.fillStyle=\'#e8553d\';ctx.beginPath();ctx.arc(lx,ly,4,0,Math.PI*2);ctx.fill();\n  ctx.fillStyle=\'#fff\';ctx.font=\'bold 10px sans-serif\';ctx.textAlign=\'left\';\n  ctx.fillText(data[data.length-1].val.toFixed(4),lx+8,ly+4);\n}\n\n// ── Dashboard ──────────────────────────\nasync function loadDashboard(){\n  loadHeatmap();\n  loadTopMovers(\'daily_growth\');\n  loadDashRankings();\n}\n\nasync function loadHeatmap(){\n  const themes=await api(\'/api/themes\');\n  const grid=document.getElementById(\'heatmapGrid\');\n  const maxCount=Math.max(...themes.map(t=>t.count));\n  grid.innerHTML=themes.map(t=>{\n    const ratio=t.count/maxCount;\n    let color;\n    if(ratio>0.8) color=\'#e8553d\';\n    else if(ratio>0.6) color=\'#f08070\';\n    else if(ratio>0.4) color=\'#d4a233\';\n    else if(ratio>0.2) color=\'#3cc6c6\';\n    else color=\'#4d94ff\';\n    return `<div class="heatmap-cell" style="background:${color};opacity:${0.4+ratio*0.6};" onclick="switchToTheme(\'${t.name}\')"><span class="tooltip">${t.name}: ${t.count}只</span></div>`;\n  }).join(\'\');\n}\n\nasync function loadTopMovers(period){\n  document.querySelectorAll(\'.mover-tab\').forEach(b=>b.classList.remove(\'active\'));\n  event.target.classList.add(\'active\');\n  const data=await api(\'/api/ranking?sort_by=\'+period+\'&limit=20\');\n  document.getElementById(\'moverList\').innerHTML=data.data.map((r,i)=>`\n    <div class="mover-row" onclick="openDetail(\'${r.fund_code}\')">\n      <span class="mr">${i+1}</span><span class="mn">${r.fund_name}</span>\n      <span class="mv ${pcls(r[period])}">${fpct(r[period])}</span>\n    </div>`).join(\'\');\n}\n\nasync function loadDashRankings(){\n  const periods=[\'return_1y\',\'return_1m\',\'daily_growth\'];\n  const names={return_1y:\'近1年\',return_1m:\'近1月\',daily_growth:\'日涨幅\'};\n  const allData=[];\n  for(const p of periods){\n    const data=await api(\'/api/ranking?sort_by=\'+p+\'&limit=10\');\n    allData.push({period:p,name:names[p],data:data.data});\n  }\n  document.getElementById(\'dashRankings\').innerHTML=allData.map((d,i)=>{\n    const colors=[\'#e8553d\',\'#4d94ff\',\'#78b89a\'];\n    return `<div class="ranking-card">\n      <div class="card-header"><span class="dot" style="background:${colors[i]};"></span>${d.name} Top 10</div>\n      <div class="card-list">${d.data.map((r,j)=>`\n        <div class="rank-row" onclick="openDetail(\'${r.fund_code}\')">\n          <span class="rnum">${j+1}</span><span class="rname">${r.fund_name}</span>\n          <span class="rval ${pcls(r[d.period])}">${fpct(r[d.period])}</span>\n        </div>`).join(\'\')}</div></div>`;\n  }).join(\'\');\n}\n\n// ── 收益排行 ───────────────────────────\nasync function loadRanking(){\n  const period=document.getElementById(\'rankPeriod\').value;\n  const limit=document.getElementById(\'rankLimit\').value;\n  const ft=document.getElementById(\'rankType\').value;\n  const data=await api(`/api/ranking?sort_by=${period}&limit=${limit}&fund_type=${ft}`);\n  const panel=document.getElementById(\'rankingPanel\');\n  const colors=[\'#e8553d\',\'#4d94ff\',\'#78b89a\',\'#e8c547\',\'#9b7ef5\',\'#3cc6c6\'];\n  const chunkSize=10;\n  const chunks=[];\n  for(let i=0;i<data.data.length;i+=chunkSize) chunks.push(data.data.slice(i,i+chunkSize));\n  panel.innerHTML=chunks.map((ch,ci)=>`\n    <div class="ranking-card"><div class="card-header">\n      <span class="dot" style="background:${colors[ci%colors.length]};"></span>\n      ${PERIOD_NAMES[period]} Top ${ci*chunkSize+1}-${ci*chunkSize+ch.length}\n    </div><div class="card-list">${ch.map((r,i)=>`\n      <div class="rank-row" onclick="openDetail(\'${r.fund_code}\')">\n        <span class="rnum">${ci*chunkSize+i+1}</span>\n        <span class="rname">${r.fund_name}</span>\n        <span class="rval ${pcls(r[period])}">${fpct(r[period])}</span>\n      </div>`).join(\'\')}</div></div>`).join(\'\');\n}\n\n// ── 主题卡片 ───────────────────────────\nasync function loadThemeCards(){\n  const themes=await api(\'/api/themes\');\n  const icons=[\'🤖\',\'💊\',\'🔋\',\'💻\',\'🏦\',\'📡\',\'🏥\',\'🛢️\',\'🥇\',\'🏠\',\'🚀\',\'⚡\',\'🛡️\',\'💹\',\'📊\'];\n  document.getElementById(\'themeCards\').innerHTML=themes.map((t,i)=>`\n    <div class="theme-card-item" onclick="switchToTheme(\'${t.name}\')">\n      <div class="tc-icon">${icons[i%icons.length]}</div>\n      <div class="tc-name">${t.name}</div>\n      <div class="tc-count">${t.count} 只基金</div>\n    </div>`).join(\'\');\n}\n\nfunction switchToTheme(theme){\n  STATE.currentTheme=theme;STATE.currentPage=1;\n  document.querySelectorAll(\'.theme-item\').forEach(e=>e.classList.toggle(\'active\',(e.textContent.includes(theme)&&theme!==\'\')));\n  switchTab(\'market\');loadFundList();\n  document.getElementById(\'mainContent\').scrollTop=0;\n}\n\n// ── Tab切换 ────────────────────────────\nfunction switchTab(tab){\n  STATE.currentTab=tab;\n  document.querySelectorAll(\'.nav-link\').forEach(e=>e.classList.toggle(\'active\',e.dataset.tab===tab));\n  [\'market\',\'dashboard\',\'ranking\',\'themes\'].forEach(t=>{\n    document.getElementById(\'tab-\'+t).style.display=t===tab?\'block\':\'none\';\n  });\n  if(tab===\'ranking\') loadRanking();\n  if(tab===\'themes\') loadThemeCards();\n  if(tab===\'dashboard\') loadDashboard();\n  if(tab===\'market\') loadFundList();\n}\n\n// ── 搜索 ───────────────────────────────\nlet st=null;\ndocument.getElementById(\'searchInput\').addEventListener(\'input\',function(){\n  clearTimeout(st);const q=this.value.trim();\n  if(!q){document.getElementById(\'searchResults\').classList.remove(\'show\');STATE.searchQuery=\'\';STATE.currentPage=1;loadFundList();return;}\n  st=setTimeout(async()=>{\n    const data=await api(\'/api/search?q=\'+encodeURIComponent(q));\n    const rs=document.getElementById(\'searchResults\');\n    if(data.length){rs.innerHTML=data.map(r=>`<div class="search-result-item" onclick="selSR(\'${r.fund_code}\')"><span>${r.fund_name}</span><span class="stype">${r.fund_code} · ${r.fund_type||\'\'}</span></div>`).join(\'\');rs.classList.add(\'show\');}\n    else{rs.innerHTML=\'<div class="search-result-item" style="color:var(--text-muted);">未找到匹配基金</div>\';rs.classList.add(\'show\');}\n  },250);\n});\nfunction selSR(code){document.getElementById(\'searchResults\').classList.remove(\'show\');document.getElementById(\'searchInput\').value=\'\';STATE.searchQuery=\'\';openDetail(code);}\ndocument.addEventListener(\'click\',e=>{if(!e.target.closest(\'.search-box\'))document.getElementById(\'searchResults\').classList.remove(\'show\');});\n\n// ── 事件绑定 ───────────────────────────\ndocument.querySelectorAll(\'.nav-link\').forEach(b=>b.addEventListener(\'click\',()=>switchTab(b.dataset.tab)));\ndocument.getElementById(\'sortBy\').addEventListener(\'change\',function(){STATE.currentSort=this.value;STATE.currentPage=1;loadFundList();});\ndocument.getElementById(\'sortOrder\').addEventListener(\'change\',function(){STATE.currentOrder=this.value;STATE.currentPage=1;loadFundList();});\ndocument.getElementById(\'filterType\').addEventListener(\'change\',function(){STATE.currentFilter=this.value;STATE.currentPage=1;loadFundList();});\ndocument.getElementById(\'perPage\').addEventListener(\'change\',function(){STATE.perPage=parseInt(this.value);STATE.currentPage=1;loadFundList();});\ndocument.getElementById(\'rankPeriod\').addEventListener(\'change\',loadRanking);\ndocument.getElementById(\'rankLimit\').addEventListener(\'change\',loadRanking);\ndocument.getElementById(\'rankType\').addEventListener(\'change\',loadRanking);\ndocument.getElementById(\'compareToggle\').addEventListener(\'click\',()=>document.getElementById(\'compareBar\').classList.toggle(\'show\'));\ndocument.getElementById(\'modalOverlay\').addEventListener(\'click\',function(e){if(e.target===this)closeDetail();});\ndocument.addEventListener(\'keydown\',function(e){if(e.key===\'Escape\')closeDetail();});\n\n// ── 启动 ───────────────────────────────\ninit();\n</script>\n</body>\n</html>',
}


def load_embedded_module(name, source_name):
    source = EMBEDDED_SOURCES[source_name]
    module = types.ModuleType(name)
    module.__file__ = os.path.join(SCRIPT_DIR, source_name.replace('/', '_'))
    module.__package__ = ''
    module.__dict__.update({
        '__name__': name,
        '__file__': module.__file__,
        '__package__': '',
    })
    exec(source, module.__dict__)
    sys.modules[name] = module
    return module

jijin_system = load_embedded_module('jijin_system', 'jijin_system.py')

# ========================================================
# Fix: monkey-patch jijin_system.run_clean_list / run_crawler
# to ensure they run with the correct working directory.
# Otherwise the crawler writes to CWD-relative paths while
# the Flask API reads from an absolute path, causing "stale data".
# ========================================================
_orig_run_clean = jijin_system.run_clean_list
_orig_run_crawler = jijin_system.run_crawler


def _patched_clean(log):
    import os as _os
    _old = _os.getcwd()
    try:
        _os.chdir(SCRIPT_DIR)
        return _orig_run_clean(log)
    finally:
        _os.chdir(_old)


def _patched_crawl(log, on_progress=None, on_done=None):
    import os as _os
    _old = _os.getcwd()
    try:
        _os.chdir(SCRIPT_DIR)
        return _orig_run_crawler(log, on_progress, on_done)
    finally:
        _os.chdir(_old)


jijin_system.run_clean_list = _patched_clean
jijin_system.run_crawler = _patched_crawl

# Also patch inside the jijin_system module so any internal references are consistent
jijin_system.run_clean_list = _patched_clean
jijin_system.run_crawler = _patched_crawl

desktop_app_module = load_embedded_module('desktop_app', 'desktop_app.py')
app_module = load_embedded_module('app', 'app.py')
app = app_module.app
INDEX_HTML = EMBEDDED_SOURCES['static/index.html']

FUND_DATA_DIR = os.path.join(SCRIPT_DIR, 'fund_data')


# ========================================================
# Fast local data index
# ========================================================
# The latest fund_profile JSON can be hundreds of MB because it contains
# nav_history for every fund. Loading that file on a UI click makes Tkinter
# appear frozen. Keep a SQLite code -> fund-json index so detail/compare views
# can read one fund at a time.
_INDEX_BUILDING = False
_INDEX_LOCK = threading.Lock()
_INDEX_READY_CACHE = {"key": None, "ready": False}


def _latest_fund_profile_file():
    files = glob.glob(os.path.join(SCRIPT_DIR, "fund_data", "fund_profile_*.json"))
    files = [p for p in files if not p.endswith(".tmp") and os.path.getsize(p) > 1024]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def _fund_index_db_path():
    return os.path.join(SCRIPT_DIR, "fund_cache", "fund_index.sqlite")


def _iter_json_array_items(path):
    import json as _json

    decoder = _json.JSONDecoder()
    chunk_size = 1024 * 1024
    with open(path, "r", encoding="utf-8") as f:
        buf = ""
        pos = 0
        eof = False
        started = False
        while True:
            if not eof and len(buf) - pos < chunk_size // 2:
                buf = buf[pos:]
                pos = 0
                chunk = f.read(chunk_size)
                if chunk:
                    buf += chunk
                else:
                    eof = True

            while pos < len(buf) and buf[pos] in " \r\n\t,":
                pos += 1

            if not started:
                if pos >= len(buf):
                    if eof:
                        return
                    continue
                if buf[pos] == "[":
                    started = True
                    pos += 1
                    continue
                raise ValueError("fund_profile JSON must be an array")

            if pos >= len(buf):
                if eof:
                    return
                continue
            if buf[pos] == "]":
                return

            try:
                item, end = decoder.raw_decode(buf, pos)
            except _json.JSONDecodeError:
                if eof:
                    raise
                continue
            pos = end
            yield item


def _fund_index_is_ready(data_path=None):
    import sqlite3 as _sqlite3

    data_path = data_path or _latest_fund_profile_file()
    if not data_path:
        return False
    db_path = _fund_index_db_path()
    if not os.path.exists(db_path):
        return False
    cache_key = (
        os.path.abspath(data_path),
        os.path.getmtime(data_path),
        os.path.getsize(data_path),
        os.path.getmtime(db_path),
        os.path.getsize(db_path),
    )
    if _INDEX_READY_CACHE.get("key") == cache_key:
        return _INDEX_READY_CACHE.get("ready", False)
    try:
        conn = _sqlite3.connect(db_path, timeout=1)
        cur = conn.cursor()
        cur.execute("select value from meta where key='source_path'")
        source_path = cur.fetchone()
        cur.execute("select value from meta where key='source_mtime'")
        source_mtime = cur.fetchone()
        cur.execute("select value from meta where key='source_size'")
        source_size = cur.fetchone()
        conn.close()
        ready = (
            source_path and source_path[0] == os.path.abspath(data_path)
            and source_mtime and source_mtime[0] == str(os.path.getmtime(data_path))
            and source_size and source_size[0] == str(os.path.getsize(data_path))
        )
        _INDEX_READY_CACHE["key"] = cache_key
        _INDEX_READY_CACHE["ready"] = bool(ready)
        return bool(ready)
    except Exception:
        _INDEX_READY_CACHE["key"] = cache_key
        _INDEX_READY_CACHE["ready"] = False
        return False


def _build_fund_index(data_path=None, log=None):
    import json as _json
    import sqlite3 as _sqlite3
    import time as _time

    data_path = data_path or _latest_fund_profile_file()
    if not data_path:
        if log:
            log("未找到基金 JSON，跳过索引构建。")
        return False
    if _fund_index_is_ready(data_path):
        if log:
            log("基金详情索引已是最新。")
        return True

    os.makedirs(os.path.join(SCRIPT_DIR, "fund_cache"), exist_ok=True)
    db_path = _fund_index_db_path()
    tmp_path = db_path + ".tmp"
    if os.path.exists(tmp_path):
        os.remove(tmp_path)

    t0 = _time.time()
    conn = _sqlite3.connect(tmp_path)
    cur = conn.cursor()
    cur.execute("pragma journal_mode=off")
    cur.execute("pragma synchronous=off")
    cur.execute("create table meta (key text primary key, value text)")
    cur.execute("create table funds (code text primary key, name text, detail_json text)")
    count = 0
    batch = []
    for item in _iter_json_array_items(data_path):
        if not isinstance(item, dict):
            continue
        code = str(item.get("fund_code", "")).zfill(6)
        if not code.strip("0"):
            continue
        name = str(item.get("fund_name") or "")
        batch.append((code, name, _json.dumps(item, ensure_ascii=False, separators=(",", ":"))))
        count += 1
        if len(batch) >= 500:
            cur.executemany("insert or replace into funds(code,name,detail_json) values(?,?,?)", batch)
            batch.clear()
    if batch:
        cur.executemany("insert or replace into funds(code,name,detail_json) values(?,?,?)", batch)
    cur.executemany(
        "insert into meta(key,value) values(?,?)",
        [
            ("source_path", os.path.abspath(data_path)),
            ("source_mtime", str(os.path.getmtime(data_path))),
            ("source_size", str(os.path.getsize(data_path))),
            ("built_at", _dt.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("fund_count", str(count)),
        ],
    )
    conn.commit()
    conn.close()
    os.replace(tmp_path, db_path)
    if log:
        log(f"基金详情索引完成：{count:,} 只，用时 {_time.time() - t0:.1f}s。")
    return True


def _get_fund_from_index(code):
    import json as _json
    import sqlite3 as _sqlite3

    code = str(code).zfill(6)
    if not _fund_index_is_ready():
        return None
    try:
        conn = _sqlite3.connect(_fund_index_db_path(), timeout=1)
        cur = conn.cursor()
        cur.execute("select detail_json from funds where code=?", (code,))
        row = cur.fetchone()
        conn.close()
        if row and row[0]:
            return _json.loads(row[0])
    except Exception:
        return None
    return None


def _ensure_fund_index_async(log=None):
    global _INDEX_BUILDING
    data_path = _latest_fund_profile_file()
    if not data_path or _fund_index_is_ready(data_path):
        return
    with _INDEX_LOCK:
        if _INDEX_BUILDING:
            return
        _INDEX_BUILDING = True

    def _worker():
        global _INDEX_BUILDING
        try:
            import subprocess as _subprocess
            if log:
                log("正在后台建立基金详情索引；这是独立进程，不会阻塞界面。")
            proc = _subprocess.Popen(
                [sys.executable, os.path.abspath(__file__), "--build-fund-index", data_path],
                cwd=SCRIPT_DIR,
                stdout=_subprocess.DEVNULL,
                stderr=_subprocess.DEVNULL,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
            proc.wait()
            if log:
                if proc.returncode == 0 and _fund_index_is_ready(data_path):
                    log("基金详情索引已完成，详情/走势/对比将按代码快速读取。")
                else:
                    log("基金详情索引未完成，可稍后再点详情或重新运行。")
        except Exception as exc:
            if log:
                log(f"基金详情索引失败：{exc}")
        finally:
            _INDEX_BUILDING = False

    threading.Thread(target=_worker, daemon=True).start()

# Ensure embedded Flask app reads fund_data from this workspace first
try:
    # prefer fund_data next to this unified_system.py
    candidate = FUND_DATA_DIR
    if os.path.isdir(candidate):
        setattr(app_module, 'DATA_DIR', candidate)
    else:
        # fallback to commonly used locations
        candidate2 = os.path.join(SCRIPT_DIR, '..', 'fund_data')
        if os.path.isdir(candidate2):
            setattr(app_module, 'DATA_DIR', os.path.abspath(candidate2))
        else:
            candidate3 = os.path.join(os.getcwd(), 'fund_data')
            if os.path.isdir(candidate3):
                setattr(app_module, 'DATA_DIR', os.path.abspath(candidate3))
except Exception:
    pass

# Ensure the embedded Tk visualization uses the same data directory as the web API.
try:
    if os.path.isdir(FUND_DATA_DIR):
        setattr(desktop_app_module, '_DATA_DIR', FUND_DATA_DIR)
        setattr(desktop_app_module, 'DATA_DIR', FUND_DATA_DIR)
except Exception:
    pass


def _run_in_script_dir(func, *args, **kwargs):
    old_cwd = os.getcwd()
    try:
        os.chdir(SCRIPT_DIR)
        return func(*args, **kwargs)
    finally:
        os.chdir(old_cwd)

# ========================================================
# Fix 1: Shorten cache TTL (10min → 60s) so API responses
#        reflect fresh data sooner after an update.
# ========================================================
setattr(app_module, 'CACHE_TTL', 60)

# ========================================================
# Fix 2: Monkey-patch _run_update_pipeline to reset the
#        global controller BEFORE each crawl.  Without this,
#        old results accumulate across updates and get
#        merged into the output JSON, causing stale data.
# ========================================================
_orig_update_pipeline = app_module._run_update_pipeline

def _patched_update_pipeline():
    # Reset the shared crawl controller so only THIS run's
    # results end up in the saved JSON file.
    jijin_system.controller.reset()
    return _orig_update_pipeline()

app_module._run_update_pipeline = _patched_update_pipeline

# Remove built-in static route generated by Flask when static_url_path is blank
# This prevents '/' from being intercepted by the static file handler.
try:
    app.static_folder = None
    app.static_url_path = None
    if hasattr(app.url_map, '_rules') and hasattr(app.url_map, '_rules_by_endpoint'):
        rules_to_remove = [rule for rule in list(app.url_map.iter_rules()) if rule.endpoint == 'static']
        for rule in rules_to_remove:
            try:
                app.url_map._rules.remove(rule)
            except ValueError:
                pass
        app.url_map._rules_by_endpoint.pop('static', None)
    app.view_functions.pop('static', None)
except Exception:
    pass

# Serve embedded index and static assets when no external static folder exists
def _guess_mimetype(name: str):
    if name.endswith('.css'):
        return 'text/css'
    if name.endswith('.js'):
        return 'application/javascript'
    if name.endswith('.png'):
        return 'image/png'
    if name.endswith('.jpg') or name.endswith('.jpeg'):
        return 'image/jpeg'
    if name.endswith('.svg'):
        return 'image/svg+xml'
    if name.endswith('.json'):
        return 'application/json'
    return 'text/plain'

def _embedded_index():
    return INDEX_HTML, 200, {'Content-Type': 'text/html; charset=utf-8'}

def _embedded_static(filename):
    key = f'static/{filename}'
    if key in EMBEDDED_SOURCES:
        content = EMBEDDED_SOURCES[key]
        return content, 200, {'Content-Type': _guess_mimetype(filename)}
    return ('', 404)

# Register routes with unique endpoints to avoid overwrite assertions
try:
    app.add_url_rule('/', 'embedded_index', _embedded_index)
    app.add_url_rule('/static/<path:filename>', 'embedded_static', _embedded_static)
except Exception:
    pass

# Log all incoming requests to a file and console to help debug 404s
def _log_request():
    try:
        entry = f"{_dt.now().isoformat()} {request.remote_addr} {request.method} {request.path}\n"
        print('[REQ]', entry.strip())
        with open(os.path.join(SCRIPT_DIR, 'request.log'), 'a', encoding='utf-8') as _f:
            _f.write(entry)

        # Serve embedded root index
        if request.path == '/':
            return INDEX_HTML, 200, {'Content-Type': 'text/html; charset=utf-8'}

        # Serve embedded favicon if present
        if request.path == '/favicon.ico':
            key = 'static/favicon.ico'
            if key in EMBEDDED_SOURCES:
                return EMBEDDED_SOURCES[key], 200, {'Content-Type': 'image/x-icon'}
            return ('', 204)

        # Serve embedded static files under /static/
        if request.path.startswith('/static/'):
            filename = request.path[len('/static/'):]
            key = f'static/{filename}'
            if key in EMBEDDED_SOURCES:
                content = EMBEDDED_SOURCES[key]
                return content, 200, {'Content-Type': _guess_mimetype(filename)}
            return ('', 404)

    except Exception:
        pass

try:
    app.before_request(_log_request)
except Exception:
    pass

# SPA catch-all: serve index.html for non-API, non-static GET paths (supports client-side routing)
def _spa_catch(subpath):
    if request.method != 'GET':
        return ('', 404)
    p = request.path or '/'
    if p.startswith('/api') or p.startswith('/static'):
        return ('', 404)
    return INDEX_HTML, 200, {'Content-Type': 'text/html; charset=utf-8'}

try:
    app.add_url_rule('/<path:subpath>', 'spa_catch', _spa_catch)
except Exception:
    pass

# If the embedded app already registered `index` or `static`, override their view functions
try:
    if 'index' in app.view_functions:
        app.view_functions['index'] = (lambda: (_embedded_index()))
    if 'static' in app.view_functions:
        def _static_override(filename):
            return _embedded_static(filename)
        app.view_functions['static'] = _static_override
except Exception:
    pass

def _install_integrated_result_viewer():
    """Replace Excel-first result handling in FundToolsApp with an integrated viewer."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    def _open_file(path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", path])
            else:
                import subprocess
                subprocess.call(["xdg-open", path])
        except Exception as exc:
            raise RuntimeError(str(exc))

    def _numeric_series(pd, series):
        return pd.to_numeric(
            series.astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.replace("，", "", regex=False)
            .str.replace("分", "", regex=False)
            .str.strip(),
            errors="coerce",
        )

    def _pick_name_column(df):
        candidates = [
            "基金名称", "基金简称", "名称", "fund_name", "产品名称",
            "代码", "基金代码", "fund_code",
        ]
        for col in candidates:
            if col in df.columns:
                return col
        return df.columns[0] if len(df.columns) else None

    def _pick_score_column(pd, df):
        preferred_keywords = [
            "综合得分", "总分", "收益得分", "风险得分", "效率得分", "位置得分",
            "趋势得分", "经理得分", "成本得分", "评分", "得分", "score",
            "近1年", "年化", "收益", "夏普", "回撤",
        ]
        for key in preferred_keywords:
            for col in df.columns:
                if key.lower() in str(col).lower():
                    vals = _numeric_series(pd, df[col])
                    if vals.notna().sum() >= 2:
                        return col, vals
        best_col, best_vals, best_count = None, None, 0
        for col in df.columns:
            vals = _numeric_series(pd, df[col])
            count = vals.notna().sum()
            if count > best_count:
                best_col, best_vals, best_count = col, vals, count
        return best_col, best_vals

    def _build_table(parent, df):
        import tkinter as tk
        from tkinter import ttk

        container = tk.Frame(parent, bg="#111827")
        container.pack(fill="both", expand=True)

        cols = [str(c) for c in df.columns]
        tree = ttk.Treeview(container, columns=cols, show="headings", height=18)
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in cols:
            width = 150
            if "名称" in col:
                width = 220
            elif "代码" in col:
                width = 95
            tree.heading(col, text=col)
            tree.column(col, width=width, anchor="center", stretch=False)

        max_rows = min(len(df), 500)
        for _, row in df.head(max_rows).iterrows():
            vals = []
            for value in row.tolist():
                text = "" if value is None else str(value)
                vals.append(text[:120])
            tree.insert("", "end", values=vals)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        return tree

    def _build_chart(parent, pd, df, sheet_name):
        import tkinter as tk

        chart_frame = tk.Frame(parent, bg="#111827", height=330)
        chart_frame.pack(fill="x", padx=10, pady=(8, 4))
        chart_frame.pack_propagate(False)

        score_col, score_vals = _pick_score_column(pd, df)
        name_col = _pick_name_column(df)
        if not score_col or score_vals is None or score_vals.notna().sum() < 2:
            tk.Label(
                chart_frame,
                text="当前 Sheet 没有足够的数值列，已切换为表格浏览。",
                bg="#111827", fg="#cbd5e1", font=("Microsoft YaHei", 11),
            ).pack(expand=True)
            return

        try:
            import matplotlib
            matplotlib.use("TkAgg")
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
            import matplotlib.font_manager as fm

            font_names = {f.name for f in fm.fontManager.ttflist}
            font_family = "Microsoft YaHei" if "Microsoft YaHei" in font_names else "SimHei"

            plot_df = df.copy()
            plot_df["_score_value_"] = score_vals
            plot_df = plot_df.dropna(subset=["_score_value_"])
            plot_df = plot_df.sort_values("_score_value_", ascending=False).head(15)

            names = (
                plot_df[name_col].astype(str).tolist()
                if name_col in plot_df.columns
                else [str(i + 1) for i in range(len(plot_df))]
            )
            values = plot_df["_score_value_"].tolist()

            fig = Figure(figsize=(9, 3.3), dpi=100, facecolor="#111827")
            ax = fig.add_subplot(111, facecolor="#111827")
            colors = ["#ef4444" if v >= 0 else "#22c55e" for v in values]
            ax.bar(range(len(values)), values, color=colors, alpha=0.92)
            ax.set_title(f"{sheet_name} - {score_col} Top {len(values)}", color="#e5e7eb", fontname=font_family, fontsize=12)
            ax.tick_params(axis="x", colors="#cbd5e1", labelrotation=35, labelsize=8)
            ax.tick_params(axis="y", colors="#cbd5e1", labelsize=8)
            ax.set_xticks(range(len(names)))
            ax.set_xticklabels([n[:12] for n in names], fontname=font_family)
            ax.grid(axis="y", color="#334155", alpha=0.45)
            for spine in ax.spines.values():
                spine.set_color("#334155")
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        except Exception as exc:
            tk.Label(
                chart_frame,
                text=f"图表渲染失败，已保留表格浏览：{exc}",
                bg="#111827", fg="#fca5a5", font=("Microsoft YaHei", 10),
            ).pack(expand=True)

    def _show_result_visualization(self, path):
        import tkinter as tk
        from tkinter import ttk, messagebox
        import pandas as pd

        if not path or not os.path.exists(path):
            messagebox.showerror("结果不存在", f"未找到结果文件：\n{path}")
            return

        try:
            sheets = pd.read_excel(path, sheet_name=None)
        except Exception as exc:
            messagebox.showerror("无法读取结果", f"结果文件读取失败：\n{path}\n\n{exc}")
            return

        win = tk.Toplevel(self.root)
        win.title(f"可视化结果 - {os.path.basename(path)}")
        win.geometry("1280x820")
        win.minsize(980, 640)
        win.configure(bg="#0f172a")
        try:
            win.focus_set()
        except Exception:
            pass

        header = tk.Frame(win, bg="#0f172a")
        header.pack(fill="x", padx=14, pady=(12, 8))

        tk.Label(
            header,
            text="评分结果可视化",
            bg="#0f172a",
            fg="#f8fafc",
            font=("Microsoft YaHei", 17, "bold"),
        ).pack(side="left")
        tk.Label(
            header,
            text=os.path.basename(path),
            bg="#0f172a",
            fg="#94a3b8",
            font=("Microsoft YaHei", 10),
        ).pack(side="left", padx=14)

        def open_excel():
            try:
                _open_file(path)
            except Exception as exc:
                messagebox.showerror("打开失败", str(exc))

        tk.Button(
            header,
            text="打开 Excel 原文件",
            command=open_excel,
            bg="#38bdf8",
            fg="#082f49",
            relief="flat",
            padx=14,
            pady=7,
            cursor="hand2",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(side="right")

        summary = tk.Frame(win, bg="#0f172a")
        summary.pack(fill="x", padx=14, pady=(0, 8))
        total_rows = sum(len(df) for df in sheets.values())
        cards = [
            ("Sheet 数", len(sheets)),
            ("总行数", f"{total_rows:,}"),
            ("文件位置", os.path.dirname(path)),
        ]
        for title, value in cards:
            card = tk.Frame(summary, bg="#1e293b", padx=14, pady=8)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(card, text=title, bg="#1e293b", fg="#94a3b8", font=("Microsoft YaHei", 9)).pack(anchor="w")
            tk.Label(card, text=str(value), bg="#1e293b", fg="#f8fafc", font=("Microsoft YaHei", 12, "bold")).pack(anchor="w")

        notebook = ttk.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        for sheet_name, df in sheets.items():
            page = tk.Frame(notebook, bg="#111827")
            notebook.add(page, text=str(sheet_name)[:18])
            if df.empty:
                tk.Label(page, text="这个 Sheet 没有数据", bg="#111827", fg="#cbd5e1").pack(expand=True)
                continue
            _build_chart(page, pd, df, str(sheet_name))
            table_wrap = tk.Frame(page, bg="#111827")
            table_wrap.pack(fill="both", expand=True, padx=10, pady=(4, 10))
            _build_table(table_wrap, df)

    def _visual_ask_open_excel(self, path):
        try:
            self._show_result_visualization(path)
            self._log(f"已打开可视化结果：{path}")
        except Exception as exc:
            self._log(f"可视化打开失败，回退到 Excel：{exc}")
            try:
                _open_file(path)
            except Exception as open_exc:
                self._log(f"无法自动打开文件: {open_exc}")

    base_cls._show_result_visualization = _show_result_visualization
    base_cls._ask_open_excel = _visual_ask_open_excel


_install_integrated_result_viewer()


def _install_scoring_fixes():
    """Patch scoring helpers for duplicate fund codes in large crawled datasets."""
    try:
        pd = jijin_system.pd
        RETURN_WEIGHTS = jijin_system.RETURN_WEIGHTS
    except Exception:
        return

    def _first_scalar(value):
        if hasattr(value, "dropna"):
            non_na = value.dropna()
            if len(non_na) > 0:
                return non_na.iloc[0]
            return None
        return value

    def _calc_score_no_duplicate_index_bug(df, age_series):
        pct_df = pd.DataFrame(index=df.index)
        for col in df.columns:
            numeric_col = pd.to_numeric(df[col], errors="coerce")
            pct_df[col] = jijin_system.percentile_rank(jijin_system.winsorize(numeric_col))
        scores = {}
        for code in df.index:
            age = _first_scalar(age_series.get(code))
            total_w, score_sum = 0, 0
            for metric, weight in RETURN_WEIGHTS.items():
                if metric not in pct_df:
                    continue
                if not jijin_system.allowed(metric, age):
                    continue
                value = _first_scalar(pct_df.at[code, metric])
                if pd.isna(value):
                    continue
                total_w += weight
                score_sum += value * weight
            if total_w == 0:
                scores[code] = None
            else:
                score = score_sum / total_w * 100
                if total_w < 40:
                    score *= 0.9
                scores[code] = round(score, 2)
        return pd.Series(scores)

    jijin_system.calc_score = _calc_score_no_duplicate_index_bug

    def _compute_long_term_scores_no_duplicate_index_bug(results):
        rows = []
        ages = {}
        for item in results:
            perf = item.get("performance", {}) or {}
            base = item.get("base_info", {}) or {}
            code = item.get("fund_code")
            age = jijin_system.calc_age(base.get("setup_date"), perf.get("nav_date"))
            ages[code] = (age or 0) * 12
            rows.append({
                "fund_code": code,
                "r_3y": jijin_system.parse_pct_to_float(perf.get("3y")),
                "r_5y": jijin_system.parse_pct_to_float(perf.get("5y")),
            })

        df = pd.DataFrame(rows)
        if df.empty or "fund_code" not in df.columns:
            return pd.Series(dtype=float)
        df = df.set_index("fund_code")
        for col in ["r_3y", "r_5y"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        pct_df = pd.DataFrame(index=df.index)
        for col in ["r_3y", "r_5y"]:
            series = df[col]
            if series.dropna().size > 1:
                pct_df[col] = jijin_system._percentile_rank(jijin_system._winsorize_series(series))
            else:
                pct_df[col] = series

        weights = {"r_3y": 45, "r_5y": 55}
        scores = {}
        for code in df.index:
            age_months = _first_scalar(ages.get(code, 0)) or 0
            if age_months < 36:
                scores[code] = None
                continue
            total_w, score_sum = 0.0, 0.0
            for metric, weight in weights.items():
                if metric not in pct_df.columns:
                    continue
                value = _first_scalar(pct_df.at[code, metric])
                if value is None or pd.isna(value):
                    continue
                total_w += weight
                score_sum += float(value) * weight
            scores[code] = None if total_w == 0 else round(score_sum / total_w * 100, 2)
        return pd.Series(scores, dtype=float)

    jijin_system._compute_long_term_scores = _compute_long_term_scores_no_duplicate_index_bug
    jijin_system._compute_long_term_via = lambda results: _compute_long_term_scores_no_duplicate_index_bug(results)

    def _eff_score_no_duplicate_index_bug(metric_df, age_months_series):
        import math as _m
        pct_df = pd.DataFrame(index=metric_df.index)
        for col in ["sharpe", "calmar", "sortino"]:
            if col not in metric_df.columns:
                continue
            numeric_col = pd.to_numeric(metric_df[col], errors="coerce")
            if numeric_col.dropna().shape[0] > 1:
                pct_df[col] = jijin_system._percentile_rank(jijin_system._winsorize_series(numeric_col))
            else:
                pct_df[col] = numeric_col

        scores = {}
        for code in metric_df.index:
            age_m = _first_scalar(age_months_series.get(code))
            if age_m is None or pd.isna(age_m) or float(age_m) < jijin_system.EFF_MIN_AGE_MONTHS:
                scores[code] = None
                continue
            total_w, score_sum = 0.0, 0.0
            for metric, weight in jijin_system.EFF_WEIGHTS.items():
                if metric not in pct_df.columns:
                    continue
                value = _first_scalar(pct_df.at[code, metric])
                if value is None or pd.isna(value):
                    continue
                total_w += weight
                score_sum += float(value) * weight
            if total_w == 0:
                scores[code] = None
            else:
                score = score_sum / total_w * 100
                if float(age_m) < jijin_system.EFF_YOUNG_AGE_MONTHS:
                    score *= jijin_system.EFF_YOUNG_PENALTY
                scores[code] = round(score, 2)
        return pd.Series(scores, dtype=float)

    jijin_system._eff_score = _eff_score_no_duplicate_index_bug

    def _series_get_scalar(series, code):
        if series is None:
            return None
        try:
            if getattr(series, "empty", False):
                return None
            value = series.get(code)
            return _first_scalar(value)
        except Exception:
            return None

    def _df_cell_scalar(df, code, column):
        if df is None or getattr(df, "empty", True) or column not in df.columns:
            return None
        try:
            if code not in df.index:
                return None
            value = df.loc[code, column]
            return _first_scalar(value)
        except Exception:
            return None

    def _patched_run_topic_screen(topic_name, log):
        spec = jijin_system.TOPIC_SPECS.get(topic_name)
        if spec is None:
            log(f"未知专题：{topic_name}")
            return None
        try:
            _, results = jijin_system._load_latest_json(log)
            if not results:
                return None
            pool = [item for item in results if jijin_system._match_topic(item, spec)]
            log(f"专题 [{topic_name}] 匹配到 {len(pool)} 只基金（全市场 {len(results)}）")
            if not pool:
                log("未匹配到任何基金。可能关键词需要调整。")
                return None

            risk_s, risk_metrics_df, mdd_map = jijin_system._compute_risk_scores(pool)
            dd_s = jijin_system._compute_drawdown_scores(pool, precomputed_mdd=mdd_map)
            ret_s = jijin_system._compute_return_scores(pool)
            long_s = jijin_system._compute_long_term_via(pool)
            eff_s = jijin_system._compute_efficiency_scores(pool)
            pos_s = jijin_system._compute_position_scores(pool)
            tim_s, _ = jijin_system._compute_timing_scores_detail(pool)
            mgr_s, _ = jijin_system._compute_manager_scores_detail(pool)
            cost_s, cost_detail = jijin_system._compute_cost_scores_detail(pool)

            def access_score_for(item):
                status = (item.get("status") or {}).get("buy_status") or ""
                return jijin_system._cost_score_access(status)

            score_map = {
                "return": ret_s,
                "risk": risk_s,
                "drawdown": dd_s,
                "long_term": long_s,
                "efficiency": eff_s,
                "position": pos_s,
                "timing": tim_s,
                "manager": mgr_s,
                "cost": cost_s,
            }
            weights = spec["weights"]
            rows = []

            for item in pool:
                code = item.get("fund_code")
                total_w, score_sum = 0.0, 0.0
                dim_values = {}
                for key, weight in weights.items():
                    if key == "access":
                        value = access_score_for(item)
                    else:
                        value = _series_get_scalar(score_map.get(key), code)
                    if value is None or pd.isna(value):
                        dim_values[key] = None
                        continue
                    value = float(value)
                    dim_values[key] = value
                    score_sum += value * weight
                    total_w += weight
                if total_w < 0.5:
                    continue

                row = jijin_system._build_base_row(item)
                row.update({
                    "专题评分": round(score_sum / total_w, 2),
                    "专题": topic_name,
                })
                for key in weights:
                    if key == "access":
                        row["申购状态分"] = round(access_score_for(item), 2)
                    else:
                        value = dim_values.get(key)
                        row[f"{key}分"] = round(float(value), 2) if value is not None and not pd.isna(value) else "--"

                row["申购费率"] = _df_cell_scalar(cost_detail, code, "申购费率") or row.get("申购费率", "--")
                row["规模(亿)"] = _df_cell_scalar(cost_detail, code, "规模(亿)") or row.get("规模(亿)", "--")
                row["买入状态"] = _df_cell_scalar(cost_detail, code, "买入状态") or row.get("买入状态", "--")
                row["年化收益"] = _df_cell_scalar(risk_metrics_df, code, "年化收益") or row.get("年化收益", "--")
                row["最大回撤"] = _df_cell_scalar(risk_metrics_df, code, "最大回撤") or row.get("最大回撤", "--")
                row["夏普比率"] = _df_cell_scalar(risk_metrics_df, code, "夏普比率") or row.get("夏普比率", "--")
                rows.append(row)

            if not rows:
                log("专题池内无法计算评分（数据不足）。")
                return None

            df = pd.DataFrame(rows).sort_values("专题评分", ascending=False, na_position="last")
            top10 = df.head(10).copy()
            output_dir = "fund_excel"
            os.makedirs(output_dir, exist_ok=True)
            timestamp = _dt.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(output_dir, f"{spec['filename']}_{timestamp}.xlsx")

            comments = dict(getattr(jijin_system, "_COMMON_EXPLAIN", {}))
            comments.update({
                "专题评分": "专题总分：按当前专题配置权重聚合，缺失维度自动按有效权重归一化。",
                "专题": f"当前专题 = {topic_name}",
            })
            extra = {"Top 10": top10} if not top10.empty else None
            jijin_system._strategy_write_excel(df, out_path, score_col="专题评分", extra_sheets=extra, comments=comments)

            for _, row in top10.head(3).iterrows():
                log(f"  {row.get('基金名称', '')}  {topic_name}专题分 = {row.get('专题评分', '--')}")
            log(f"{topic_name}专题筛选完成，共 {len(df)} 只基金。")
            log(f"已导出: {out_path}")
            return os.path.abspath(out_path)
        except Exception as exc:
            import traceback
            log(f"{topic_name}专题筛选出错: {exc}")
            log(traceback.format_exc())
            return None

    jijin_system.run_topic_screen = _patched_run_topic_screen


_install_scoring_fixes()


def _install_polished_result_reporter():
    """Use a Tonghuashun-style HTML dashboard for score/screening results."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    def _clean_cell(value):
        try:
            import pandas as pd
            if pd.isna(value):
                return ""
        except Exception:
            pass
        return value.item() if hasattr(value, "item") else value

    def _to_number(value):
        if value is None:
            return None
        try:
            text = str(value).replace("%", "").replace(",", "").replace("，", "").replace("分", "").strip()
            if text in ("", "--", "nan", "None"):
                return None
            return float(text)
        except Exception:
            return None

    def _pick_col(columns, keywords, fallback=None):
        for keyword in keywords:
            for col in columns:
                if keyword.lower() in str(col).lower():
                    return col
        return fallback

    def _make_result_report(self, excel_path):
        import json
        import os
        import html
        from datetime import datetime

        import pandas as pd

        sheets_raw = pd.read_excel(excel_path, sheet_name=None)
        sheets = []
        total_rows = 0
        all_scores = []
        all_types = {}

        for sheet_name, df in sheets_raw.items():
            df = df.fillna("")
            rows = []
            columns = [str(c) for c in df.columns]
            name_col = _pick_col(columns, ["基金名称", "基金简称", "名称", "fund_name"], columns[0] if columns else "")
            code_col = _pick_col(columns, ["基金代码", "代码", "fund_code"], "")
            score_col = _pick_col(columns, ["综合得分", "收益表现评分", "风险得分", "效率得分", "位置得分", "趋势得分", "经理得分", "成本得分", "评分", "得分", "score"], "")
            type_col = _pick_col(columns, ["基金类型", "类型"], "")

            if not score_col:
                best_col, best_count = "", 0
                for col in columns:
                    count = sum(1 for v in df[col].head(300).tolist() if _to_number(v) is not None)
                    if count > best_count:
                        best_col, best_count = col, count
                score_col = best_col

            for _, record in df.iterrows():
                item = {str(col): _clean_cell(record[col]) for col in df.columns}
                score = _to_number(item.get(score_col)) if score_col else None
                item["_name"] = str(item.get(name_col, ""))
                item["_code"] = str(item.get(code_col, ""))
                item["_score"] = score
                item["_type"] = str(item.get(type_col, ""))
                rows.append(item)
                if score is not None:
                    all_scores.append(score)
                if item["_type"]:
                    all_types[item["_type"]] = all_types.get(item["_type"], 0) + 1

            total_rows += len(rows)
            rows.sort(key=lambda x: (x["_score"] is not None, x["_score"] or -10**9), reverse=True)
            sheets.append({
                "name": str(sheet_name),
                "columns": columns,
                "rows": rows,
                "scoreCol": score_col,
                "nameCol": name_col,
                "codeCol": code_col,
                "typeCol": type_col,
            })

        title = os.path.splitext(os.path.basename(excel_path))[0]
        report_dir = os.path.join(SCRIPT_DIR, "fund_visual_reports")
        os.makedirs(report_dir, exist_ok=True)
        report_path = os.path.join(report_dir, f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")

        avg_score = round(sum(all_scores) / len(all_scores), 2) if all_scores else "--"
        max_score = round(max(all_scores), 2) if all_scores else "--"
        min_score = round(min(all_scores), 2) if all_scores else "--"
        top_types = sorted(all_types.items(), key=lambda x: x[1], reverse=True)[:8]
        payload = {
            "title": title,
            "excelPath": os.path.abspath(excel_path),
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": {
                "sheets": len(sheets),
                "rows": total_rows,
                "avgScore": avg_score,
                "maxScore": max_score,
                "minScore": min_score,
                "topTypes": top_types,
            },
            "sheets": sheets,
        }

        data_json = json.dumps(payload, ensure_ascii=False)
        excel_label = html.escape(os.path.abspath(excel_path))

        template = r'''
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>基金评分结果可视化</title>
<style>
:root{
  --bg:#080d18; --panel:#101827; --panel2:#151f31; --line:#263247;
  --text:#e8edf6; --muted:#8f9bb2; --red:#e8553d; --green:#6fb894;
  --gold:#e8b830; --blue:#3b82f6; --cyan:#28c6d3;
}
*{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--text);
  font-family:"Microsoft YaHei",Arial,sans-serif;letter-spacing:0}
.app{min-height:100vh;display:grid;grid-template-columns:250px 1fr}
.side{background:#0b1220;border-right:1px solid var(--line);padding:18px 14px;position:sticky;top:0;height:100vh}
.brand{font-size:20px;font-weight:800;color:var(--gold);margin-bottom:6px}
.sub{font-size:12px;color:var(--muted);line-height:1.7;word-break:break-all}
.nav{margin-top:18px;display:flex;flex-direction:column;gap:8px}
.nav button{border:1px solid transparent;background:transparent;color:var(--muted);
  text-align:left;border-radius:6px;padding:10px 12px;cursor:pointer;font-weight:700}
.nav button.active,.nav button:hover{background:var(--panel2);border-color:var(--line);color:var(--text)}
.main{padding:18px 22px 28px;min-width:0}
.topbar{display:flex;align-items:center;gap:12px;margin-bottom:14px}
.title{font-size:24px;font-weight:900}.tag{font-size:12px;color:#111827;background:var(--gold);
  padding:4px 8px;border-radius:4px;font-weight:800}
.tools{margin-left:auto;display:flex;gap:10px}
input,select{background:#0d1524;border:1px solid var(--line);color:var(--text);
  border-radius:6px;padding:9px 10px;outline:none}
.cards{display:grid;grid-template-columns:repeat(5,minmax(130px,1fr));gap:12px;margin-bottom:14px}
.card{background:linear-gradient(180deg,#172236,#101827);border:1px solid var(--line);
  border-radius:8px;padding:13px 14px}
.label{font-size:12px;color:var(--muted);margin-bottom:6px}.value{font-size:22px;font-weight:900}
.value.red{color:var(--red)}.value.green{color:var(--green)}.value.gold{color:var(--gold)}
.grid{display:grid;grid-template-columns:1.35fr .65fr;gap:14px;margin-bottom:14px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:14px;min-width:0}
.panel h3{margin:0 0 12px;font-size:15px;color:#dce5f5}
canvas{width:100%;height:330px;display:block}
.type-list{display:grid;gap:8px}
.type-row{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;font-size:13px}
.bar{height:7px;background:#243047;border-radius:999px;overflow:hidden;margin-top:5px}
.bar i{display:block;height:100%;background:linear-gradient(90deg,var(--blue),var(--cyan))}
.table-wrap{background:var(--panel);border:1px solid var(--line);border-radius:8px;overflow:hidden}
table{width:100%;border-collapse:collapse;font-size:13px}
thead{position:sticky;top:0;background:#172236;z-index:2}
th,td{border-bottom:1px solid #1d2940;padding:9px 10px;text-align:right;white-space:nowrap}
th{color:#aeb9cd;font-weight:800;cursor:pointer}td:first-child,th:first-child{text-align:left}
tbody tr:hover{background:#162236}.pos{color:var(--red);font-weight:800}.neg{color:var(--green);font-weight:800}
.rank{color:var(--gold);font-weight:900}.empty{padding:60px;text-align:center;color:var(--muted)}
@media(max-width:980px){.app{grid-template-columns:1fr}.side{height:auto;position:relative}.cards{grid-template-columns:repeat(2,1fr)}.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
  <aside class="side">
    <div class="brand">基金数据集成平台</div>
    <div class="sub">同花顺风格结果看板<br>生成时间：<span id="gen"></span><br><br>源文件：<span id="src"></span></div>
    <div class="nav" id="sheetNav"></div>
  </aside>
  <main class="main">
    <div class="topbar">
      <div class="title" id="title"></div><span class="tag">可视化结果</span>
      <div class="tools"><input id="search" placeholder="搜索基金名称/代码/类型"><select id="limit"><option value="50">前50</option><option value="100">前100</option><option value="300">前300</option><option value="999999">全部</option></select></div>
    </div>
    <section class="cards">
      <div class="card"><div class="label">Sheet 数</div><div class="value" id="cSheets"></div></div>
      <div class="card"><div class="label">记录数</div><div class="value gold" id="cRows"></div></div>
      <div class="card"><div class="label">最高分</div><div class="value red" id="cMax"></div></div>
      <div class="card"><div class="label">平均分</div><div class="value" id="cAvg"></div></div>
      <div class="card"><div class="label">最低分</div><div class="value green" id="cMin"></div></div>
    </section>
    <section class="grid">
      <div class="panel"><h3 id="chartTitle">Top 排行</h3><canvas id="barChart"></canvas></div>
      <div class="panel"><h3>基金类型分布</h3><div id="typeList" class="type-list"></div></div>
    </section>
    <section class="table-wrap"><div id="table"></div></section>
  </main>
</div>
<script id="payload" type="application/json">__DATA__</script>
<script>
const DATA=JSON.parse(document.getElementById('payload').textContent);
let current=0, sortKey='', sortDir=-1;
const $=id=>document.getElementById(id);
function fmt(v){return typeof v==='number'?v.toLocaleString('zh-CN',{maximumFractionDigits:2}):v;}
function num(v){if(v===null||v===undefined)return null;const n=parseFloat(String(v).replace(/[%分,，]/g,''));return Number.isFinite(n)?n:null}
function cls(v){const n=num(v); if(n===null)return ''; return n>=0?'pos':'neg'}
function init(){
 $('title').textContent=DATA.title; $('gen').textContent=DATA.generatedAt; $('src').textContent=DATA.excelPath;
 $('cSheets').textContent=DATA.summary.sheets; $('cRows').textContent=fmt(DATA.summary.rows);
 $('cMax').textContent=fmt(DATA.summary.maxScore); $('cAvg').textContent=fmt(DATA.summary.avgScore); $('cMin').textContent=fmt(DATA.summary.minScore);
 $('sheetNav').innerHTML=DATA.sheets.map((s,i)=>`<button class="${i===0?'active':''}" onclick="switchSheet(${i})">${s.name} <span style="float:right">${s.rows.length}</span></button>`).join('');
 $('search').addEventListener('input',render); $('limit').addEventListener('change',render);
 renderTypes(); render();
}
function switchSheet(i){current=i; sortKey=''; document.querySelectorAll('.nav button').forEach((b,j)=>b.classList.toggle('active',j===i)); render();}
function rowsFiltered(){
 const s=DATA.sheets[current], q=$('search').value.trim().toLowerCase(); let rows=s.rows.slice();
 if(q) rows=rows.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q)));
 if(sortKey) rows.sort((a,b)=>{const na=num(a[sortKey]),nb=num(b[sortKey]); if(na!==null&&nb!==null)return (na-nb)*sortDir; return String(a[sortKey]??'').localeCompare(String(b[sortKey]??''),'zh-CN')*sortDir});
 const limit=parseInt($('limit').value,10); return rows.slice(0,limit);
}
function render(){
 const s=DATA.sheets[current], rows=rowsFiltered(); $('chartTitle').textContent=`${s.name} - ${s.scoreCol||'数值'} Top 20`;
 drawChart(rows.slice(0,20),s); renderTable(rows,s);
}
function renderTypes(){
 const max=Math.max(...DATA.summary.topTypes.map(x=>x[1]),1);
 $('typeList').innerHTML=DATA.summary.topTypes.map(([name,count])=>`<div class="type-row"><div>${name}<div class="bar"><i style="width:${count/max*100}%"></i></div></div><b>${count}</b></div>`).join('')||'<div class="empty">暂无类型数据</div>';
}
function drawChart(rows,s){
 const c=$('barChart'),ctx=c.getContext('2d'),dpr=devicePixelRatio||1,w=c.clientWidth,h=c.clientHeight;c.width=w*dpr;c.height=h*dpr;ctx.scale(dpr,dpr);ctx.clearRect(0,0,w,h);
 const vals=rows.map(r=>num(r[s.scoreCol]??r._score)).filter(v=>v!==null); if(!vals.length){ctx.fillStyle='#8f9bb2';ctx.fillText('暂无可绘制数值',40,60);return}
 const max=Math.max(...vals), min=Math.min(0,...vals), range=max-min||1, pad={l:46,r:18,t:16,b:72}, cw=w-pad.l-pad.r,ch=h-pad.t-pad.b, bw=Math.max(8,cw/rows.length*0.62);
 ctx.strokeStyle='#263247';ctx.lineWidth=1;ctx.font='12px Microsoft YaHei';ctx.fillStyle='#8f9bb2';
 for(let i=0;i<=4;i++){const y=pad.t+ch*i/4;ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(w-pad.r,y);ctx.stroke();ctx.fillText((max-range*i/4).toFixed(1),4,y+4)}
 rows.forEach((r,i)=>{const v=num(r[s.scoreCol]??r._score); if(v===null)return; const x=pad.l+cw*(i+.5)/rows.length-bw/2, y=pad.t+ch-(v-min)/range*ch; const bh=pad.t+ch-y; ctx.fillStyle=v>=0?'#e8553d':'#6fb894'; ctx.fillRect(x,y,bw,bh); ctx.save();ctx.translate(x+bw/2,h-12);ctx.rotate(-Math.PI/5);ctx.fillStyle='#aeb9cd';ctx.textAlign='right';ctx.fillText(String(r._name||r._code||i+1).slice(0,8),0,0);ctx.restore();});
}
function renderTable(rows,s){
 if(!rows.length){$('table').innerHTML='<div class="empty">没有匹配数据</div>';return}
 const cols=s.columns;
 $('table').innerHTML=`<table><thead><tr><th>排行</th>${cols.map(c=>`<th onclick="sortBy('${String(c).replace(/'/g,"\\'")}')">${c}</th>`).join('')}</tr></thead><tbody>${rows.map((r,i)=>`<tr><td class="rank">${i+1}</td>${cols.map(c=>`<td class="${cls(r[c])}">${r[c]??''}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}
function sortBy(c){sortDir=sortKey===c?-sortDir:-1; sortKey=c; render();}
init();
</script>
</body>
</html>
'''
        html_text = template.replace("__DATA__", data_json.replace("</script>", "<\\/script>"))
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html_text)
        return os.path.abspath(report_path)

    def _show_polished_result(self, path):
        import os
        import webbrowser
        from tkinter import messagebox

        if not path or not os.path.exists(path):
            messagebox.showerror("结果不存在", f"未找到结果文件：\n{path}")
            return
        report_path = self._make_result_report(path)
        webbrowser.open("file:///" + report_path.replace("\\", "/"))
        self._log(f"已生成同花顺风格可视化看板：{report_path}")

    def _polished_ask_open_excel(self, path):
        try:
            self._show_result_visualization(path)
        except Exception as exc:
            self._log(f"可视化看板生成失败：{exc}")

    base_cls._make_result_report = _make_result_report
    base_cls._show_result_visualization = _show_polished_result
    base_cls._ask_open_excel = _polished_ask_open_excel


_install_polished_result_reporter()


def _install_native_tonghuashun_viewer():
    """Show scoring results inside the desktop app instead of opening a browser."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    COLORS = {
        "bg": "#080d18",
        "side": "#0b1220",
        "panel": "#101827",
        "panel2": "#151f31",
        "line": "#263247",
        "text": "#e8edf6",
        "muted": "#8f9bb2",
        "red": "#e8553d",
        "green": "#6fb894",
        "gold": "#e8b830",
        "blue": "#3b82f6",
        "cyan": "#28c6d3",
    }

    def _num(value):
        if value is None:
            return None
        try:
            text = str(value).replace("%", "").replace(",", "").replace("，", "").replace("分", "").strip()
            if text in ("", "--", "nan", "None"):
                return None
            return float(text)
        except Exception:
            return None

    def _clean(value):
        try:
            import pandas as pd
            if pd.isna(value):
                return ""
        except Exception:
            pass
        return "" if value is None else str(value)

    def _watchlist_path():
        return os.path.join(SCRIPT_DIR, "watchlist_funds.json")

    def _load_watchlist():
        import json
        path = _watchlist_path()
        if not os.path.exists(path):
            return {"默认自选": []}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {"默认自选": []}
        except Exception:
            return {"默认自选": []}

    def _save_watchlist(data):
        import json
        with open(_watchlist_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _add_watchlist(code, name, group="默认自选"):
        data = _load_watchlist()
        items = data.setdefault(group, [])
        code = str(code).zfill(6)
        if code and all(str(item.get("code", "")).zfill(6) != code for item in items):
            items.append({"code": code, "name": name or code})
            _save_watchlist(data)
            return True
        return False

    def _pick_col(columns, keywords, fallback=""):
        for keyword in keywords:
            for col in columns:
                if keyword.lower() in str(col).lower():
                    return col
        return fallback

    def _viewer_profile(title_hint, columns):
        text = str(title_hint).lower()
        col_text = " ".join(str(c) for c in columns)

        def has(*words):
            source = text + " " + col_text
            return any(str(word).lower() in source for word in words)

        common = ["基金代码", "基金名称", "最新净值", "今日涨幅", "近1月", "近3月", "近6月", "近1年", "基金类型", "基金经理"]
        profiles = [
            (has("风险回撤", "风险决策", "回撤", "drawdown", "risk"),
             "风险回撤", "夏普比率",
             ["夏普比率", "最大回撤", "当前回撤", "回撤进度", "年化收益", "波动率", "下行波动", "决策建议", "卡玛比率", "索提诺比率"]),
            (has("风险效率", "efficiency", "sharpe", "calmar", "sortino"),
             "风险效率", "风险效率评分",
             ["风险效率评分", "Sharpe估算", "Calmar估算", "Sortino估算", "年化收益", "最大回撤估算", "成立年限", "规模"]),
            (has("位置估值", "位置", "估值", "position"),
             "位置估值", "位置得分",
             ["位置得分", "估值位置", "当前分位", "近1年分位", "最大回撤", "当前回撤", "最新净值", "规模"]),
            (has("趋势择时", "趋势", "timing"),
             "趋势择时", "趋势得分",
             ["趋势得分", "趋势状态", "均线状态", "动量", "近1月", "近3月", "近6月", "当前回撤"]),
            (has("基金经理", "经理", "manager"),
             "基金经理", "经理得分",
             ["经理得分", "基金经理", "任职年限", "管理规模", "代表基金", "近1年", "近3年", "最大回撤"]),
            (has("交易成本", "成本", "cost", "费率"),
             "交易成本", "成本得分",
             ["成本得分", "申购费率", "管理费率", "托管费率", "销售服务费", "买入状态", "申购状态分", "规模"]),
            (has("收益归因", "归因", "attribution"),
             "收益归因", "归因评分",
             ["归因评分", "主要收益来源", "近1月", "近3月", "近6月", "近1年", "行业主题", "标签"]),
            (has("长期综合", "综合", "composite"),
             "长期综合", "综合得分",
             ["综合得分", "收益表现评分", "风险效率评分", "位置得分", "趋势得分", "经理得分", "成本得分", "近1年", "最大回撤"]),
            (has("回撤震荡", "震荡"),
             "回撤震荡", "回撤震荡评分",
             ["回撤震荡评分", "当前回撤", "最大回撤", "回撤进度", "波动率", "夏普比率", "决策建议", "近1月"]),
            (has("趋势突破", "突破"),
             "趋势突破", "趋势突破评分",
             ["趋势突破评分", "趋势状态", "突破信号", "近1月", "近3月", "近6月", "今日涨幅", "当前回撤"]),
            (has("低波稳健", "低波", "稳健"),
             "低波稳健", "低波稳健评分",
             ["低波稳健评分", "波动率", "最大回撤", "夏普比率", "近1年", "近3年", "规模", "基金经理"]),
            (has("超跌反弹", "超跌", "反弹"),
             "超跌反弹", "超跌反弹评分",
             ["超跌反弹评分", "当前回撤", "最大回撤", "回撤进度", "近1月", "近3月", "今日涨幅", "决策建议"]),
            (has("收益表现", "performance"),
             "收益表现", "收益表现评分",
             ["收益表现评分", "近1月", "近3月", "近6月", "近1年", "近3年", "成立以来", "成立年限"]),
        ]
        for matched, name, primary, cols in profiles:
            if matched:
                return {"name": name, "primary": primary, "priority": ["基金代码", "基金名称"] + cols + common}
        return {"name": "通用结果", "primary": "综合得分", "priority": ["基金代码", "基金名称", "专题评分", "综合得分"] + common}

    _SHEET_CACHE = {}
    _FUND_MAP_CACHE = {"key": None, "value": {}}

    def _prepare_sheets(path):
        import pandas as pd

        cache_key = (os.path.abspath(path), os.path.getmtime(path), os.path.getsize(path))
        cached = _SHEET_CACHE.get(cache_key)
        if cached:
            return cached
        if len(_SHEET_CACHE) >= 3:
            _SHEET_CACHE.clear()

        raw = pd.read_excel(path, sheet_name=None)
        sheets = []
        all_scores = []
        type_counts = {}
        total_rows = 0

        for sheet_name, df in raw.items():
            df = df.fillna("")
            columns = [str(c) for c in df.columns]
            profile = _viewer_profile(os.path.basename(path) + " " + str(sheet_name), columns)
            fallback = columns[0] if columns else ""
            name_col = _pick_col(columns, ["基金名称", "基金简称", "名称", "fund_name"], fallback)
            code_col = _pick_col(columns, ["基金代码", "代码", "fund_code"], "")
            type_col = _pick_col(columns, ["基金类型", "类型"], "")
            score_col = _pick_col(
                columns,
                [profile["primary"], "专题评分", "综合得分", "收益表现评分", "风险效率评分", "风险得分", "效率得分", "位置得分", "趋势得分", "经理得分", "成本得分", "评分", "得分", "score"],
                "",
            )
            if not score_col:
                best_col, best_count = "", 0
                for col in columns:
                    count = sum(1 for value in df[col].head(300).tolist() if _num(value) is not None)
                    if count > best_count:
                        best_col, best_count = col, count
                score_col = best_col

            rows = []
            for record in df.to_dict("records"):
                row = {str(col): _clean(record.get(col, "")) for col in df.columns}
                if "今日涨幅" not in row:
                    row["今日涨幅"] = row.get("日增长率", row.get("日涨跌幅", ""))
                score = _num(row.get(score_col)) if score_col else None
                row["_name"] = row.get(name_col, "")
                row["_code"] = row.get(code_col, "")
                row["_type"] = row.get(type_col, "")
                row["_score"] = score
                rows.append(row)
                if score is not None:
                    all_scores.append(score)
                if row["_type"]:
                    type_counts[row["_type"]] = type_counts.get(row["_type"], 0) + 1

            rows.sort(key=lambda r: (r["_score"] is not None, r["_score"] if r["_score"] is not None else -10**9), reverse=True)
            if "今日涨幅" not in columns and ("日增长率" in columns or "日涨跌幅" in columns):
                columns.append("今日涨幅")
            front_cols = profile["priority"]
            back_cols = ["基金类型", "标签", "基金经理", "成立日期", "净值日期"]
            ordered = []
            for col in front_cols:
                if col in columns and col not in ordered:
                    ordered.append(col)
            for col in columns:
                if col in ("日增长率", "日涨跌幅") and "今日涨幅" in columns:
                    continue
                if col not in ordered and col not in back_cols:
                    ordered.append(col)
            for col in back_cols:
                if col in columns and col not in ordered:
                    ordered.append(col)
            columns = ordered
            total_rows += len(rows)
            sheets.append({
                "name": str(sheet_name),
                "df": df,
                "columns": columns,
                "rows": rows,
                "name_col": name_col,
                "code_col": code_col,
                "type_col": type_col,
                "score_col": score_col,
                "profile": profile,
            })

        summary = {
            "sheet_count": len(sheets),
            "row_count": total_rows,
            "max_score": max(all_scores) if all_scores else None,
            "avg_score": sum(all_scores) / len(all_scores) if all_scores else None,
            "min_score": min(all_scores) if all_scores else None,
            "top_types": sorted(type_counts.items(), key=lambda item: item[1], reverse=True)[:8],
        }
        result = (sheets, summary)
        _SHEET_CACHE[cache_key] = result
        return result

    def _latest_fund_map():
        import glob
        import json

        files = glob.glob(os.path.join(SCRIPT_DIR, "fund_data", "fund_profile_*.json"))
        files += glob.glob(os.path.join("fund_data", "fund_profile_*.json"))
        files = [
            path for path in sorted(set(files), key=os.path.getmtime, reverse=True)
            if not path.endswith(".tmp") and os.path.getsize(path) > 1024
        ]
        cache_key = tuple((os.path.abspath(path), os.path.getmtime(path), os.path.getsize(path)) for path in files[:3])
        if _FUND_MAP_CACHE.get("key") == cache_key:
            return _FUND_MAP_CACHE.get("value", {})

        data = []
        for path in files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                rows = loaded if isinstance(loaded, list) else loaded.get("results", [])
                if rows:
                    data = rows
                    break
            except Exception:
                continue
        if not data:
            return {}
        mapping = {}
        for item in data:
            code = str(item.get("fund_code", "")).zfill(6)
            if code and code not in mapping:
                mapping[code] = item
        _FUND_MAP_CACHE["key"] = cache_key
        _FUND_MAP_CACHE["value"] = mapping
        return mapping

    class _LazyFundMap:
        def __init__(self):
            self._loaded = False
            self._loading = False
            self._data = {}
            self._lock = threading.Lock()

        def preload(self):
            _ensure_fund_index_async()

        def _ensure_loaded(self):
            self._loaded = True

        def get(self, code, default=None):
            fund = _get_fund_from_index(code)
            if fund:
                return fund
            _ensure_fund_index_async()
            return default

    def _fund_code_from_row(row, sheet):
        code = row.get("_code") or row.get(sheet.get("code_col", ""))
        digits = "".join(ch for ch in str(code) if ch.isdigit())
        return digits.zfill(6) if digits else ""

    def _draw_nav_lines(canvas, series_list, normalized=False):
        canvas.delete("all")
        canvas.update_idletasks()
        width = max(canvas.winfo_width(), 760)
        height = max(canvas.winfo_height(), 320)
        pad_l, pad_t, pad_r, pad_b = 58, 24, 24, 46
        chart_w = width - pad_l - pad_r
        chart_h = height - pad_t - pad_b
        colors = [COLORS["red"], COLORS["blue"], COLORS["green"], COLORS["gold"], COLORS["cyan"], "#c084fc"]

        prepared = []
        for idx, item in enumerate(series_list):
            hist = item.get("history") or []
            points = []
            for row in hist:
                try:
                    value = float(row.get("val"))
                    date = str(row.get("date", ""))
                    points.append((date, value))
                except Exception:
                    continue
            if len(points) < 2:
                continue
            points = points[-260:]
            if normalized:
                base = points[0][1] or 1
                points = [(d, v / base) for d, v in points]
            prepared.append({
                "name": item.get("name") or item.get("code") or f"基金{idx + 1}",
                "points": points,
                "color": colors[idx % len(colors)],
            })

        if not prepared:
            canvas.create_text(40, 70, text="暂无净值走势数据", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 12))
            return

        values = [value for series in prepared for _, value in series["points"]]
        min_v, max_v = min(values), max(values)
        if min_v == max_v:
            min_v *= 0.98
            max_v *= 1.02
        span = max(max_v - min_v, 1e-9)

        for i in range(6):
            y = pad_t + chart_h * i / 5
            value = max_v - span * i / 5
            canvas.create_line(pad_l, y, width - pad_r, y, fill=COLORS["line"])
            canvas.create_text(8, y, text=f"{value:.3f}", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 9))

        for series in prepared:
            points = series["points"]
            coords = []
            for idx, (_, value) in enumerate(points):
                x = pad_l + chart_w * idx / max(len(points) - 1, 1)
                y = pad_t + chart_h - (value - min_v) / span * chart_h
                coords.extend([x, y])
            canvas.create_line(*coords, fill=series["color"], width=2)
            lx, ly = coords[-2], coords[-1]
            canvas.create_oval(lx - 3, ly - 3, lx + 3, ly + 3, fill=series["color"], width=0)

        legend_x = pad_l
        for series in prepared[:6]:
            canvas.create_rectangle(legend_x, height - 24, legend_x + 18, height - 20, fill=series["color"], width=0)
            canvas.create_text(legend_x + 24, height - 22, text=str(series["name"])[:14], fill=COLORS["text"], anchor="w", font=("Microsoft YaHei", 9))
            legend_x += 145

    def _history_points(nav_history):
        from datetime import datetime as _datetime
        points = []
        for row in nav_history or []:
            try:
                date_text = str(row.get("date", ""))[:10]
                date_obj = _datetime.strptime(date_text, "%Y-%m-%d").date()
                value = float(row.get("val"))
                if value > 0:
                    points.append((date_obj, date_text, value))
            except Exception:
                continue
        points.sort(key=lambda item: item[0])
        return points

    def _filter_history_period(points, period):
        from datetime import timedelta as _timedelta
        if not points:
            return []
        if period == "all":
            return points
        last_day = points[-1][0]
        if period == "ytd":
            cutoff = last_day.replace(month=1, day=1)
        else:
            days = {"1m": 31, "3m": 92, "6m": 183, "1y": 365, "3y": 365 * 3, "5y": 365 * 5}.get(period)
            cutoff = last_day if not days else last_day - _timedelta(days=days)
        filtered = [item for item in points if item[0] >= cutoff]
        return filtered if len(filtered) >= 2 else points[-min(len(points), 30):]

    def _draw_single_nav_chart(canvas, nav_history, period):
        from datetime import date as _date
        canvas.delete("all")
        canvas.update_idletasks()
        points = _filter_history_period(_history_points(nav_history), period)
        width = max(canvas.winfo_width(), 760)
        height = max(canvas.winfo_height(), 320)
        pad_l, pad_t, pad_r, pad_b = 66, 78, 26, 46
        chart_w = width - pad_l - pad_r
        chart_h = height - pad_t - pad_b
        if len(points) < 2:
            canvas.create_text(40, 70, text="暂无净值走势数据", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 12))
            return

        values = [item[2] for item in points]
        min_v, max_v = min(values), max(values)
        pad = max((max_v - min_v) * 0.08, max_v * 0.01)
        min_v -= pad
        max_v += pad
        span = max(max_v - min_v, 1e-9)

        max_draw_points = max(360, min(1200, int(chart_w // 2) if chart_w > 0 else 760))
        if len(points) > max_draw_points:
            sample_step = max(1, len(points) // max_draw_points)
            draw_points = [(idx, item) for idx, item in enumerate(points) if idx % sample_step == 0]
            if draw_points[-1][0] != len(points) - 1:
                draw_points.append((len(points) - 1, points[-1]))
        else:
            draw_points = list(enumerate(points))

        def x_at(index):
            return pad_l + chart_w * index / max(len(points) - 1, 1)

        def y_at(value):
            return pad_t + chart_h - (value - min_v) / span * chart_h

        peak_val = points[0][2]
        peak_idx = 0
        max_dd = 0.0
        max_peak_idx = 0
        max_trough_idx = 0
        current_peak = points[0][2]
        for idx, (_date_obj, _date_text, value) in enumerate(points):
            if value > peak_val:
                peak_val = value
                peak_idx = idx
            dd = value / peak_val - 1.0 if peak_val else 0.0
            if dd < max_dd:
                max_dd = dd
                max_peak_idx = peak_idx
                max_trough_idx = idx
            current_peak = max(current_peak, value)

        repair_idx = None
        repair_target = points[max_peak_idx][2]
        for idx in range(max_trough_idx + 1, len(points)):
            if points[idx][2] >= repair_target:
                repair_idx = idx
                break
        current_dd = points[-1][2] / current_peak - 1.0 if current_peak else 0.0
        progress = abs(current_dd) / abs(max_dd) * 100 if max_dd < 0 else 0.0
        progress = max(0.0, min(progress, 100.0))
        fall_days = (points[max_trough_idx][0] - points[max_peak_idx][0]).days
        if repair_idx is not None:
            repair_days = (points[repair_idx][0] - points[max_trough_idx][0]).days
            repair_text = f"已修复 {repair_days}天"
        else:
            repair_days = (_date.today() - points[max_trough_idx][0]).days
            repair_text = f"未修复 {repair_days}天"

        chips = [
            ("最大回撤", f"{max_dd * 100:.2f}%"),
            ("回撤进度", f"{progress:.1f}%"),
            ("当前回撤", f"{current_dd * 100:.2f}%"),
            ("修复状态", repair_text),
        ]
        chip_x = pad_l
        for label, value in chips:
            chip_w = 132 if label != "修复状态" else 154
            canvas.create_rectangle(chip_x, 16, chip_x + chip_w, 58, fill=COLORS["panel2"], outline=COLORS["line"])
            canvas.create_text(chip_x + 10, 28, text=label, fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 8))
            chip_color = COLORS["green"] if "回撤" in label or "进度" in label else COLORS["text"]
            canvas.create_text(chip_x + 10, 47, text=value, fill=chip_color, anchor="w", font=("Microsoft YaHei", 10, "bold"))
            chip_x += chip_w + 8

        for i in range(6):
            y = pad_t + chart_h * i / 5
            value = max_v - span * i / 5
            canvas.create_line(pad_l, y, width - pad_r, y, fill=COLORS["line"])
            canvas.create_text(10, y, text=f"{value:.4f}", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 9))

        coords = []
        for idx, (_date_obj, _date_text, value) in draw_points:
            x = x_at(idx)
            y = y_at(value)
            coords.extend([x, y])
        if len(coords) >= 4:
            px, py_peak = x_at(max_peak_idx), y_at(points[max_peak_idx][2])
            tx, ty = x_at(max_trough_idx), y_at(points[max_trough_idx][2])
            canvas.create_rectangle(px, pad_t, tx, pad_t + chart_h, fill="#1f2737", outline="")
            fill_poly = [pad_l, pad_t + chart_h] + coords + [width - pad_r, pad_t + chart_h]
            canvas.create_polygon(*fill_poly, fill="#2b1720", outline="")
            canvas.create_line(*coords, fill=COLORS["red"], width=2)
            canvas.create_line(px, pad_t, px, pad_t + chart_h, fill=COLORS["gold"], dash=(4, 3))
            canvas.create_line(tx, pad_t, tx, pad_t + chart_h, fill=COLORS["green"], dash=(4, 3))
            canvas.create_oval(px - 4, py_peak - 4, px + 4, py_peak + 4, fill=COLORS["gold"], outline="")
            canvas.create_oval(tx - 4, ty - 4, tx + 4, ty + 4, fill=COLORS["green"], outline="")
            canvas.create_text(px + 8, py_peak - 14, text="回撤起点", fill=COLORS["gold"], anchor="w", font=("Microsoft YaHei", 8, "bold"))
            canvas.create_text(tx + 8, ty + 16, text=f"最大回撤 {max_dd * 100:.2f}% | 下跌{fall_days}天", fill=COLORS["green"], anchor="w", font=("Microsoft YaHei", 8, "bold"))
            if repair_idx is not None and repair_idx < len(points):
                rx, ry = x_at(repair_idx), y_at(points[repair_idx][2])
                canvas.create_line(rx, pad_t, rx, pad_t + chart_h, fill=COLORS["blue"], dash=(4, 3))
                canvas.create_oval(rx - 4, ry - 4, rx + 4, ry + 4, fill=COLORS["blue"], outline="")
                canvas.create_text(rx + 8, ry - 14, text="修复完成", fill=COLORS["blue"], anchor="w", font=("Microsoft YaHei", 8, "bold"))
            lx, ly = coords[-2], coords[-1]
            canvas.create_oval(lx - 4, ly - 4, lx + 4, ly + 4, fill=COLORS["red"], outline="")
            canvas.create_text(lx - 8, ly - 14, text=f"{points[-1][2]:.4f}", fill=COLORS["red"], anchor="e", font=("Microsoft YaHei", 9, "bold"))

        tick_count = min(6, len(points))
        step = max(1, (len(points) - 1) // max(tick_count - 1, 1))
        tick_indexes = list(range(0, len(points), step))
        if tick_indexes[-1] != len(points) - 1:
            tick_indexes.append(len(points) - 1)
        for idx in tick_indexes[:7]:
            x = x_at(idx)
            canvas.create_line(x, pad_t + chart_h, x, pad_t + chart_h + 4, fill=COLORS["line"])
            canvas.create_text(x, height - 22, text=points[idx][1][5:], fill=COLORS["muted"], font=("Microsoft YaHei", 8))

        def draw_crosshair(event):
            if not points:
                return
            x = min(max(event.x, pad_l), width - pad_r)
            idx = int(round((x - pad_l) / max(chart_w, 1) * (len(points) - 1)))
            idx = max(0, min(idx, len(points) - 1))
            px, py = x_at(idx), y_at(points[idx][2])
            date_text = points[idx][1]
            nav_value = points[idx][2]
            if idx > 0 and points[idx - 1][2]:
                pct = (nav_value / points[idx - 1][2] - 1) * 100
                pct_text = f"{pct:+.2f}%"
            else:
                pct_text = "--"
            canvas.delete("crosshair")
            canvas.create_line(px, pad_t, px, pad_t + chart_h, fill="#d1d5db", dash=(3, 3), width=1, tags="crosshair")
            canvas.create_line(pad_l, py, width - pad_r, py, fill="#475569", dash=(3, 3), width=1, tags="crosshair")
            canvas.create_oval(px - 5, py - 5, px + 5, py + 5, fill="#ffffff", outline=COLORS["red"], width=2, tags="crosshair")
            tip_w, tip_h = 148, 72
            tip_x = px + 12 if px + 12 + tip_w < width - pad_r else px - tip_w - 12
            tip_y = py - tip_h - 10 if py - tip_h - 10 > pad_t else py + 14
            canvas.create_rectangle(tip_x, tip_y, tip_x + tip_w, tip_y + tip_h,
                                    fill="#0b1220", outline=COLORS["line"], tags="crosshair")
            canvas.create_text(tip_x + 10, tip_y + 16, text=date_text, fill=COLORS["text"],
                               anchor="w", font=("Microsoft YaHei", 9, "bold"), tags="crosshair")
            canvas.create_text(tip_x + 10, tip_y + 38, text=f"净值 {nav_value:.4f}", fill=COLORS["red"],
                               anchor="w", font=("Microsoft YaHei", 10, "bold"), tags="crosshair")
            canvas.create_text(tip_x + 10, tip_y + 58, text=f"涨跌 {pct_text}", fill=COLORS["green"] if pct_text.startswith("-") else COLORS["red"],
                               anchor="w", font=("Microsoft YaHei", 9, "bold"), tags="crosshair")

        canvas.bind("<Motion>", draw_crosshair)
        canvas.bind("<Button-1>", draw_crosshair)
        canvas.bind("<B1-Motion>", draw_crosshair)

    def _draw_drawdown_chart(canvas, nav_history):
        from datetime import date as _date
        canvas.delete("all")
        canvas.update_idletasks()
        points = _history_points(nav_history)
        width = max(canvas.winfo_width(), 760)
        height = max(canvas.winfo_height(), 230)
        pad_l, pad_t, pad_r, pad_b = 66, 28, 26, 42
        chart_w = width - pad_l - pad_r
        chart_h = height - pad_t - pad_b
        if len(points) < 20:
            canvas.create_text(40, 70, text="历史净值不足，暂无法计算回撤", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 12))
            return {}

        peak_val = points[0][2]
        peak_idx = 0
        max_dd = 0.0
        max_peak_idx = 0
        max_trough_idx = 0
        dd_points = []
        for idx, (_d, _t, value) in enumerate(points):
            if value > peak_val:
                peak_val = value
                peak_idx = idx
            dd = value / peak_val - 1.0 if peak_val else 0.0
            dd_points.append(dd)
            if dd < max_dd:
                max_dd = dd
                max_peak_idx = peak_idx
                max_trough_idx = idx

        repair_idx = None
        repair_target = points[max_peak_idx][2]
        for idx in range(max_trough_idx + 1, len(points)):
            if points[idx][2] >= repair_target:
                repair_idx = idx
                break

        current_peak = max(value for _d, _t, value in points)
        current_dd = points[-1][2] / current_peak - 1.0 if current_peak else 0.0
        progress = abs(current_dd) / abs(max_dd) * 100 if max_dd < 0 else 0.0
        progress = max(0.0, min(progress, 100.0))

        min_dd = min(dd_points)
        max_dd_axis = 0.0
        span = max(max_dd_axis - min_dd, 1e-9)
        for i in range(5):
            y = pad_t + chart_h * i / 4
            value = max_dd_axis - span * i / 4
            canvas.create_line(pad_l, y, width - pad_r, y, fill=COLORS["line"])
            canvas.create_text(10, y, text=f"{value * 100:.1f}%", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 9))

        coords = []
        zero_y = pad_t
        for idx, dd in enumerate(dd_points):
            x = pad_l + chart_w * idx / max(len(dd_points) - 1, 1)
            y = pad_t + chart_h - (dd - min_dd) / span * chart_h
            coords.extend([x, y])
        if len(coords) >= 4:
            fill_poly = [pad_l, zero_y] + coords + [width - pad_r, zero_y]
            canvas.create_polygon(*fill_poly, fill="#17261f", outline="")
            canvas.create_line(*coords, fill=COLORS["green"], width=2)

        def x_at(idx):
            return pad_l + chart_w * idx / max(len(dd_points) - 1, 1)

        trough_x = x_at(max_trough_idx)
        trough_y = pad_t + chart_h - (dd_points[max_trough_idx] - min_dd) / span * chart_h
        canvas.create_line(x_at(max_peak_idx), pad_t, x_at(max_peak_idx), pad_t + chart_h, fill=COLORS["gold"], dash=(4, 3))
        canvas.create_line(trough_x, pad_t, trough_x, pad_t + chart_h, fill=COLORS["green"], dash=(4, 3))
        canvas.create_oval(trough_x - 4, trough_y - 4, trough_x + 4, trough_y + 4, fill=COLORS["green"], outline="")
        canvas.create_text(trough_x + 8, trough_y - 16, text=f"最大回撤 {max_dd * 100:.2f}%", fill=COLORS["green"], anchor="w", font=("Microsoft YaHei", 9, "bold"))
        if repair_idx is not None:
            rx = x_at(repair_idx)
            canvas.create_line(rx, pad_t, rx, pad_t + chart_h, fill=COLORS["blue"], dash=(4, 3))
            canvas.create_text(rx + 8, pad_t + 18, text="修复", fill=COLORS["blue"], anchor="w", font=("Microsoft YaHei", 9, "bold"))

        tick_indexes = [0, max_trough_idx, len(points) - 1]
        if repair_idx is not None:
            tick_indexes.insert(2, repair_idx)
        seen = set()
        for idx in tick_indexes:
            if idx in seen:
                continue
            seen.add(idx)
            x = x_at(idx)
            canvas.create_text(x, height - 20, text=points[idx][1], fill=COLORS["muted"], font=("Microsoft YaHei", 8))

        max_days = (points[max_trough_idx][0] - points[max_peak_idx][0]).days
        if repair_idx is not None:
            repair_days = (points[repair_idx][0] - points[max_trough_idx][0]).days
            repair_text = f"已修复，用时 {repair_days} 天"
        else:
            repair_days = (_date.today() - points[max_trough_idx][0]).days
            repair_text = f"未修复，已持续 {repair_days} 天"
        return {
            "最大回撤": f"{max_dd * 100:.2f}%",
            "回撤区间": f"{points[max_peak_idx][1]} -> {points[max_trough_idx][1]}",
            "下跌耗时": f"{max_days} 天",
            "修复状态": repair_text,
            "当前回撤": f"{current_dd * 100:.2f}%",
            "回撤进度": f"{progress:.1f}%",
        }

    def _open_detail_window(parent, row, sheet, fund_map):
        import tkinter as tk

        code = _fund_code_from_row(row, sheet)
        fund = fund_map.get(code, {})
        perf = fund.get("performance", {}) or {}
        base = fund.get("base_info", {}) or {}
        status = fund.get("status", {}) or {}
        name = fund.get("fund_name") or row.get("_name") or code

        win = tk.Toplevel(parent)
        win.title(f"基金详情 - {name}")
        win.geometry("1160x900")
        win.minsize(980, 760)
        win.configure(bg=COLORS["bg"])

        header = tk.Frame(win, bg=COLORS["bg"])
        header.pack(fill="x", padx=18, pady=(16, 8))
        tk.Label(header, text=str(name), bg=COLORS["bg"], fg=COLORS["text"], font=("Microsoft YaHei", 20, "bold")).pack(side="left")
        tk.Label(header, text=code, bg=COLORS["gold"], fg="#111827", padx=8, pady=4, font=("Microsoft YaHei", 10, "bold")).pack(side="left", padx=10)

        cards = tk.Frame(win, bg=COLORS["bg"])
        cards.pack(fill="x", padx=18, pady=(0, 12))
        info_cards = [
            ("单位净值", perf.get("nav", row.get("最新净值", "--")), COLORS["text"]),
            ("今日涨幅", perf.get("daily_growth_rate", row.get("今日涨幅", row.get("日增长率", "--"))), COLORS["red"] if (_num(perf.get("daily_growth_rate", row.get("今日涨幅", row.get("日增长率")))) or 0) >= 0 else COLORS["green"]),
            ("近1年", perf.get("1y", row.get("近1年", "--")), COLORS["red"] if (_num(perf.get("1y", row.get("近1年"))) or 0) >= 0 else COLORS["green"]),
            ("基金规模", base.get("assets_size", row.get("规模", "--")), COLORS["gold"]),
        ]
        for label, value, color in info_cards:
            card = tk.Frame(cards, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=8)
            tk.Label(card, text=label, bg=COLORS["panel"], fg=COLORS["muted"], font=("Microsoft YaHei", 9)).pack(anchor="w", padx=14, pady=(8, 2))
            tk.Label(card, text=str(value or "--"), bg=COLORS["panel"], fg=color, font=("Microsoft YaHei", 17, "bold")).pack(anchor="w", padx=14, pady=(0, 8))

        chart_panel = tk.Frame(win, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        chart_panel.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        chart_header = tk.Frame(chart_panel, bg=COLORS["panel"])
        chart_header.pack(fill="x", padx=14, pady=(12, 0))
        tk.Label(chart_header, text="单位净值走势", bg=COLORS["panel"], fg=COLORS["text"], font=("Microsoft YaHei", 13, "bold")).pack(side="left")
        nav_period = {"value": "1y", "buttons": {}, "redraw_job": None}
        chart = tk.Canvas(chart_panel, height=520, bg=COLORS["panel"], highlightthickness=0)
        chart.pack(fill="both", expand=True, padx=12, pady=8)

        def redraw_nav():
            nav_period["redraw_job"] = None
            for key, button in nav_period["buttons"].items():
                active = key == nav_period["value"]
                button.configure(bg=COLORS["red"] if active else COLORS["panel2"],
                                 fg="#ffffff" if active else COLORS["text"])
            _draw_single_nav_chart(chart, fund.get("nav_history", []), nav_period["value"])

        def schedule_redraw(delay=80):
            job = nav_period.get("redraw_job")
            if job:
                try:
                    win.after_cancel(job)
                except Exception:
                    pass
            nav_period["redraw_job"] = win.after(delay, redraw_nav)

        for key, label in [("1m", "1月"), ("3m", "3月"), ("6m", "6月"), ("1y", "1年"), ("3y", "3年"), ("5y", "5年"), ("ytd", "今年"), ("all", "成立来")]:
            btn = tk.Button(chart_header, text=label, command=lambda k=key: (nav_period.update({"value": k}), schedule_redraw(10)),
                            bg=COLORS["panel2"], fg=COLORS["text"], relief="flat", padx=8, pady=4,
                            cursor="hand2", font=("Microsoft YaHei", 9, "bold"))
            btn.pack(side="right", padx=2)
            nav_period["buttons"][key] = btn
        chart.bind("<Configure>", lambda _event: schedule_redraw(120))
        win.after(180, redraw_nav)

        detail = tk.Frame(win, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        detail.pack(fill="x", padx=18, pady=(0, 16))
        fields = [
            ("基金类型", base.get("fund_type", row.get("基金类型", "--"))),
            ("风险等级", base.get("risk_level", row.get("风险等级", "--"))),
            ("基金经理", base.get("manager", row.get("基金经理", "--"))),
            ("基金公司", base.get("company", row.get("基金公司", "--"))),
            ("成立日期", base.get("setup_date", row.get("成立日期", "--"))),
            ("申购状态", status.get("buy_status", row.get("买入状态", "--"))),
            ("赎回状态", status.get("sell_status", "--")),
            ("手续费", status.get("buy_fee", row.get("申购费率", "--"))),
        ]
        for idx, (label, value) in enumerate(fields):
            cell = tk.Frame(detail, bg=COLORS["panel"])
            cell.grid(row=idx // 4, column=idx % 4, sticky="ew", padx=12, pady=9)
            tk.Label(cell, text=label, bg=COLORS["panel"], fg=COLORS["muted"], font=("Microsoft YaHei", 9)).pack(anchor="w")
            tk.Label(cell, text=str(value or "--"), bg=COLORS["panel"], fg=COLORS["text"], font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")
        for col in range(4):
            detail.columnconfigure(col, weight=1)

    def _open_compare_window(parent, rows, sheet, fund_map):
        import tkinter as tk
        from tkinter import ttk

        series = []
        seen = set()
        compare_rows = []
        for row in rows:
            code = _fund_code_from_row(row, sheet)
            if not code or code in seen:
                continue
            seen.add(code)
            fund = fund_map.get(code, {})
            perf = fund.get("performance", {}) or {}
            base = fund.get("base_info", {}) or {}
            status = fund.get("status", {}) or {}
            compare_rows.append({
                "基金代码": code,
                "基金名称": fund.get("fund_name") or row.get("_name") or code,
                "近1月": perf.get("1m", row.get("近1月", "--")),
                "近3月": perf.get("3m", row.get("近3月", "--")),
                "近6月": perf.get("6m", row.get("近6月", "--")),
                "近1年": perf.get("1y", row.get("近1年", "--")),
                "近3年": perf.get("3y", row.get("近3年", "--")),
                "今日涨幅": perf.get("daily_growth_rate", row.get("今日涨幅", row.get("日增长率", row.get("日涨跌幅", "--")))),
                "最大回撤": row.get("最大回撤", "--"),
                "夏普比率": row.get("夏普比率", "--"),
                "规模": base.get("assets_size", row.get("规模", "--")),
                "基金经理": base.get("manager", row.get("基金经理", "--")),
                "申购状态": status.get("buy_status", row.get("买入状态", "--")),
                "费率": status.get("buy_fee", row.get("申购费率", "--")),
            })
            if fund.get("nav_history"):
                series.append({
                    "name": fund.get("fund_name") or row.get("_name") or code,
                    "code": code,
                    "history": fund.get("nav_history", []),
                })
        win = tk.Toplevel(parent)
        win.title("基金走势对比")
        win.geometry("1120x700")
        win.minsize(900, 560)
        win.configure(bg=COLORS["bg"])
        tk.Label(win, text="基金走势对比（归一化）", bg=COLORS["bg"], fg=COLORS["text"], font=("Microsoft YaHei", 20, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        panel = tk.Frame(win, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        chart = tk.Canvas(panel, bg=COLORS["panel"], highlightthickness=0)
        chart.pack(fill="both", expand=True, padx=12, pady=12)
        _draw_nav_lines(chart, series, normalized=True)

        table_panel = tk.Frame(win, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        table_panel.pack(fill="x", padx=18, pady=(0, 18))
        cols = ["基金代码", "基金名称", "今日涨幅", "近1月", "近3月", "近6月", "近1年", "近3年", "最大回撤", "夏普比率", "规模", "基金经理", "申购状态", "费率"]
        tree = ttk.Treeview(table_panel, columns=cols, show="headings", height=min(8, max(3, len(compare_rows))))
        tree.tag_configure("pos", foreground=COLORS["red"])
        tree.tag_configure("neg", foreground=COLORS["green"])
        for col in cols:
            tree.heading(col, text=col)
            width = 92
            if col == "基金名称":
                width = 220
            elif col == "基金经理":
                width = 130
            tree.column(col, width=width, anchor="center", stretch=False)
        for item in compare_rows:
            tags = []
            n = _num(item.get("近1年"))
            if n is not None:
                tags.append("pos" if n >= 0 else "neg")
            tree.insert("", "end", values=[item.get(col, "") for col in cols], tags=tags)
        tree.pack(fill="x", padx=10, pady=10)

    def _native_show_result(self, path):
        import os
        import tkinter as tk
        from tkinter import ttk, messagebox

        if not path or not os.path.exists(path):
            messagebox.showerror("结果不存在", f"未找到结果文件：\n{path}")
            return

        try:
            sheets, summary = _prepare_sheets(path)
        except Exception as exc:
            messagebox.showerror("读取失败", f"无法读取结果文件：\n{path}\n\n{exc}")
            return

        win = tk.Toplevel(self.root)
        win.title(f"基金结果看板 - {os.path.basename(path)}")
        win.geometry("1360x850")
        win.minsize(1080, 680)
        win.configure(bg=COLORS["bg"])
        try:
            win.focus_set()
        except Exception:
            pass

        state = {"sheet": 0, "query": "", "limit": 200, "sort_col": "", "sort_dir": -1, "tree": None, "row_by_iid": {}, "compare_rows": {}, "render_job": None, "table_job": None}
        fund_map = _LazyFundMap()

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TH.Treeview", background=COLORS["panel"], foreground=COLORS["text"],
                        fieldbackground=COLORS["panel"], rowheight=30, borderwidth=0)
        style.configure("TH.Treeview.Heading", background=COLORS["panel2"], foreground=COLORS["muted"],
                        font=("Microsoft YaHei", 10, "bold"), borderwidth=0)
        style.map("TH.Treeview", background=[("selected", "#243047")], foreground=[("selected", COLORS["text"])])

        root = tk.Frame(win, bg=COLORS["bg"])
        root.pack(fill="both", expand=True)

        side = tk.Frame(root, bg=COLORS["side"], width=250)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        main = tk.Frame(root, bg=COLORS["bg"])
        main.pack(side="left", fill="both", expand=True)

        tk.Label(side, text="基金数据集成平台", bg=COLORS["side"], fg=COLORS["gold"],
                 font=("Microsoft YaHei", 18, "bold")).pack(anchor="w", padx=16, pady=(18, 4))
        tk.Label(side, text="同花顺风格 · 软件内看板", bg=COLORS["side"], fg=COLORS["muted"],
                 font=("Microsoft YaHei", 10)).pack(anchor="w", padx=16)
        tk.Label(side, text=os.path.basename(path), bg=COLORS["side"], fg=COLORS["muted"],
                 wraplength=210, justify="left", font=("Microsoft YaHei", 9)).pack(anchor="w", padx=16, pady=(12, 10))

        nav_frame = tk.Frame(side, bg=COLORS["side"])
        nav_frame.pack(fill="both", expand=True, padx=10, pady=(6, 12))

        header = tk.Frame(main, bg=COLORS["bg"])
        header.pack(fill="x", padx=18, pady=(16, 8))
        title_var = tk.StringVar(value=os.path.splitext(os.path.basename(path))[0])
        tk.Label(header, textvariable=title_var, bg=COLORS["bg"], fg=COLORS["text"],
                 font=("Microsoft YaHei", 22, "bold")).pack(side="left")
        tk.Label(header, text="可视化结果", bg=COLORS["gold"], fg="#111827",
                 font=("Microsoft YaHei", 10, "bold"), padx=9, pady=4).pack(side="left", padx=10)

        search_var = tk.StringVar()
        limit_var = tk.StringVar(value="200")
        compare_count_var = tk.StringVar(value="对比池 0")
        tk.Button(header, text="加入自选", command=lambda: add_selected_watchlist(), bg=COLORS["green"], fg="#062015",
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Button(header, text="走势对比", command=lambda: open_selected_compare(), bg=COLORS["blue"], fg="#ffffff",
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Button(header, text="清空对比", command=lambda: clear_compare_pool(), bg=COLORS["panel2"], fg=COLORS["text"],
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Button(header, text="加入对比", command=lambda: add_selected_to_compare(), bg=COLORS["cyan"], fg="#06252a",
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Button(header, text="查看详情", command=lambda: open_selected_detail(), bg=COLORS["gold"], fg="#111827",
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Label(header, textvariable=compare_count_var, bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=(8, 0))
        tk.Entry(header, textvariable=search_var, bg="#0d1524", fg=COLORS["text"],
                 insertbackground=COLORS["text"], relief="flat", width=32,
                 font=("Microsoft YaHei", 10)).pack(side="right", ipady=8, padx=(8, 0))
        tk.Label(header, text="搜索", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("Microsoft YaHei", 10)).pack(side="right")
        limit_box = ttk.Combobox(header, textvariable=limit_var, values=["200", "500", "全部"],
                                 width=6, state="readonly", font=("Microsoft YaHei", 10))
        limit_box.pack(side="right", padx=(8, 0), ipady=5)
        tk.Label(header, text="显示", bg=COLORS["bg"], fg=COLORS["muted"],
                 font=("Microsoft YaHei", 10)).pack(side="right")

        cards_frame = tk.Frame(main, bg=COLORS["bg"])
        cards_frame.pack(fill="x", padx=18, pady=(2, 12))

        def fmt(value):
            if value is None:
                return "--"
            if isinstance(value, (int, float)):
                return f"{value:,.2f}" if abs(value) < 10000 else f"{value:,.0f}"
            return str(value)

        cards = [
            ("Sheet 数", summary["sheet_count"], COLORS["text"]),
            ("记录数", summary["row_count"], COLORS["gold"]),
            ("最高分", summary["max_score"], COLORS["red"]),
            ("平均分", summary["avg_score"], COLORS["text"]),
            ("最低分", summary["min_score"], COLORS["green"]),
        ]
        for label, value, color in cards:
            card = tk.Frame(cards_frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=8)
            tk.Label(card, text=label, bg=COLORS["panel"], fg=COLORS["muted"], font=("Microsoft YaHei", 9)).pack(anchor="w", padx=14, pady=(8, 2))
            tk.Label(card, text=fmt(value), bg=COLORS["panel"], fg=color, font=("Microsoft YaHei", 18, "bold")).pack(anchor="w", padx=14, pady=(0, 8))

        upper = tk.Frame(main, bg=COLORS["bg"])
        upper.pack(fill="x", padx=18, pady=(0, 12))
        chart_panel = tk.Frame(upper, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        chart_panel.pack(side="left", fill="both", expand=True, padx=(0, 12))
        type_panel = tk.Frame(upper, bg=COLORS["panel"], width=330, highlightbackground=COLORS["line"], highlightthickness=1)
        type_panel.pack(side="left", fill="y")
        type_panel.pack_propagate(False)

        chart_title = tk.Label(chart_panel, text="Top 排行", bg=COLORS["panel"], fg=COLORS["text"],
                               font=("Microsoft YaHei", 12, "bold"))
        chart_title.pack(anchor="w", padx=14, pady=(12, 0))
        chart = tk.Canvas(chart_panel, height=300, bg=COLORS["panel"], highlightthickness=0)
        chart.pack(fill="x", padx=12, pady=8)

        side_title = tk.Label(type_panel, text="策略信息", bg=COLORS["panel"], fg=COLORS["text"],
                              font=("Microsoft YaHei", 12, "bold"))
        side_title.pack(anchor="w", padx=14, pady=(12, 8))
        type_body = tk.Frame(type_panel, bg=COLORS["panel"])
        type_body.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        table_panel = tk.Frame(main, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        table_panel.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        table_frame = tk.Frame(table_panel, bg=COLORS["panel"])
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        nav_buttons = []

        def filtered_rows():
            sheet = sheets[state["sheet"]]
            rows = list(sheet["rows"])
            query = state["query"].lower().strip()
            if query:
                rows = [row for row in rows if any(query in str(value).lower() for value in row.values())]
            if state["sort_col"]:
                col = state["sort_col"]
                direction = state["sort_dir"]
                def sort_key(row):
                    value = _num(row.get(col))
                    return (0, value) if value is not None else (1, str(row.get(col, "")))
                rows.sort(key=sort_key, reverse=direction < 0)
            limit = state.get("limit", 200)
            if limit and limit > 0:
                return rows[:limit]
            return rows

        def draw_types():
            for child in type_body.winfo_children():
                child.destroy()
            sheet = sheets[state["sheet"]]
            profile = sheet.get("profile", {})
            side_title.config(text=f"{profile.get('name', '策略')}信息")
            tk.Label(type_body, text=f"核心指标：{profile.get('primary') or sheet.get('score_col') or '--'}",
                     bg=COLORS["panel"], fg=COLORS["gold"], font=("Microsoft YaHei", 10, "bold")).pack(anchor="w", pady=(0, 8))
            shown = 0
            for col in profile.get("priority", []):
                if col in sheet["columns"] and shown < 6:
                    tk.Label(type_body, text=f"- {col}", bg=COLORS["panel"], fg=COLORS["muted"],
                             font=("Microsoft YaHei", 9)).pack(anchor="w", pady=1)
                    shown += 1
            tk.Frame(type_body, bg=COLORS["line"], height=1).pack(fill="x", pady=10)
            top_types = summary["top_types"]
            if not top_types:
                tk.Label(type_body, text="暂无类型数据", bg=COLORS["panel"], fg=COLORS["muted"]).pack(anchor="w")
                return
            max_count = max(count for _, count in top_types) or 1
            for name, count in top_types:
                row = tk.Frame(type_body, bg=COLORS["panel"])
                row.pack(fill="x", pady=5)
                tk.Label(row, text=str(name)[:18], bg=COLORS["panel"], fg=COLORS["text"],
                         font=("Microsoft YaHei", 10)).pack(side="left")
                tk.Label(row, text=str(count), bg=COLORS["panel"], fg=COLORS["gold"],
                         font=("Microsoft YaHei", 10, "bold")).pack(side="right")
                bar_bg = tk.Frame(type_body, bg="#243047", height=7)
                bar_bg.pack(fill="x", pady=(0, 2))
                bar_bg.pack_propagate(False)
                bar = tk.Frame(bar_bg, bg=COLORS["cyan"], height=7, width=max(8, int(280 * count / max_count)))
                bar.pack(side="left")

        def draw_chart(rows):
            sheet = sheets[state["sheet"]]
            profile = sheet.get("profile", {})
            score_col = sheet["score_col"] or _pick_col(sheet["columns"], [profile.get("primary", ""), "评分", "得分"], "")
            chart.delete("all")
            chart.update_idletasks()
            width = max(chart.winfo_width(), 760)
            height = 300
            values = []
            for row in rows[:20]:
                value = _num(row.get(score_col)) if score_col else row.get("_score")
                if value is not None:
                    values.append((row, value))
            chart_title.config(text=f"{profile.get('name', sheet['name'])} - {score_col or '数值'} Top {len(values)}")
            if not values:
                chart.create_text(40, 60, text="暂无可绘制数值", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 12))
                return
            max_v = max(value for _, value in values)
            min_v = min(0, min(value for _, value in values))
            rng = max(max_v - min_v, 1)
            left, top, bottom, right = 54, 20, 64, 18
            chart_w = width - left - right
            chart_h = height - top - bottom
            for i in range(5):
                y = top + chart_h * i / 4
                chart.create_line(left, y, width - right, y, fill=COLORS["line"])
                label = max_v - rng * i / 4
                chart.create_text(8, y, text=f"{label:.1f}", fill=COLORS["muted"], anchor="w", font=("Microsoft YaHei", 9))
            bar_w = max(10, chart_w / max(len(values), 1) * 0.58)
            for i, (row, value) in enumerate(values):
                x = left + chart_w * (i + 0.5) / len(values)
                y = top + chart_h - (value - min_v) / rng * chart_h
                color = COLORS["red"] if value >= 0 else COLORS["green"]
                chart.create_rectangle(x - bar_w / 2, y, x + bar_w / 2, top + chart_h, fill=color, width=0)
                name = str(row.get("_name") or row.get("_code") or i + 1)[:8]
                chart.create_text(x, height - 34, text=name, fill=COLORS["muted"], angle=35, font=("Microsoft YaHei", 8))
                chart.create_text(x, y - 8, text=f"{value:.1f}", fill=color, font=("Microsoft YaHei", 8, "bold"))

        def render_table(rows):
            old_job = state.get("table_job")
            if old_job:
                try:
                    win.after_cancel(old_job)
                except Exception:
                    pass
                state["table_job"] = None
            for child in table_frame.winfo_children():
                child.destroy()
            sheet = sheets[state["sheet"]]
            columns = ["排行"] + sheet["columns"]
            tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="TH.Treeview", selectmode="extended")
            state["tree"] = tree
            state["row_by_iid"] = {}
            vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.tag_configure("pos", foreground=COLORS["red"])
            tree.tag_configure("neg", foreground=COLORS["green"])
            tree.tag_configure("odd", background="#0f1828")

            def on_heading(col):
                if col == "排行":
                    return
                if state["sort_col"] == col:
                    state["sort_dir"] *= -1
                else:
                    state["sort_col"] = col
                    state["sort_dir"] = -1
                render()

            for col in columns:
                tree.heading(col, text=col, command=lambda c=col: on_heading(c))
                width = 90
                if "名称" in col:
                    width = 220
                elif "代码" in col:
                    width = 96
                elif col == "排行":
                    width = 62
                tree.column(col, width=width, minwidth=width, anchor="center", stretch=False)

            insert_index = {"value": 0}

            def insert_chunk():
                if state.get("tree") is not tree:
                    return
                start = insert_index["value"]
                end = min(start + 60, len(rows))
                for pos in range(start, end):
                    idx = pos + 1
                    row = rows[pos]
                    vals = [idx] + [row.get(col, "") for col in sheet["columns"]]
                    score = row.get(sheet["score_col"], row.get("_score"))
                    tags = ["odd"] if idx % 2 else []
                    change_value = None
                    for change_col in ("今日涨幅", "日增长率", "日涨跌幅", "涨跌幅", "近1月", "近3月", "近6月", "近1年"):
                        if change_col in row:
                            change_value = _num(row.get(change_col))
                            if change_value is not None:
                                break
                    numeric_score = _num(score)
                    color_value = change_value if change_value is not None else numeric_score
                    if color_value is not None:
                        tags.append("pos" if color_value >= 0 else "neg")
                    iid = tree.insert("", "end", values=vals, tags=tags)
                    state["row_by_iid"][iid] = row
                insert_index["value"] = end
                if end < len(rows):
                    state["table_job"] = win.after(8, insert_chunk)
                else:
                    state["table_job"] = None

            def on_double_click(_event=None):
                selection = tree.selection()
                if not selection:
                    return
                row = state["row_by_iid"].get(selection[0])
                if row:
                    _open_detail_window(win, row, sheet, fund_map)

            tree.bind("<Double-1>", on_double_click)

            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            table_frame.rowconfigure(0, weight=1)
            table_frame.columnconfigure(0, weight=1)
            insert_chunk()

        def selected_rows():
            tree = state.get("tree")
            if tree is None:
                return []
            return [state["row_by_iid"].get(iid) for iid in tree.selection() if state["row_by_iid"].get(iid)]

        def open_selected_detail():
            rows = selected_rows()
            if rows:
                _open_detail_window(win, rows[0], sheets[state["sheet"]], fund_map)

        def add_selected_watchlist():
            rows = selected_rows()
            sheet = sheets[state["sheet"]]
            added = 0
            for row in rows:
                code = _fund_code_from_row(row, sheet)
                name = row.get("_name") or row.get(sheet.get("name_col", "")) or code
                if _add_watchlist(code, name):
                    added += 1
            self._log(f"已加入自选 {added} 只；自选文件：{_watchlist_path()}")

        def update_compare_count():
            compare_count_var.set(f"对比池 {len(state['compare_rows'])}")

        def add_selected_to_compare():
            sheet = sheets[state["sheet"]]
            rows = selected_rows()
            for row in rows:
                code = _fund_code_from_row(row, sheet) or str(id(row))
                state["compare_rows"][code] = (row, sheet)
            update_compare_count()

        def clear_compare_pool():
            state["compare_rows"].clear()
            update_compare_count()

        def open_selected_compare():
            if len(state["compare_rows"]) >= 2:
                rows = [row for row, _sheet in state["compare_rows"].values()]
                # Rows in the compare pool already contain resolved codes, so the current
                # sheet is enough for fallback column lookup.
                _open_compare_window(win, rows, sheets[state["sheet"]], fund_map)
                return
            rows = selected_rows()
            if len(rows) >= 2:
                _open_compare_window(win, rows, sheets[state["sheet"]], fund_map)
                return
            if len(rows) == 1:
                add_selected_to_compare()
            self._log("请先多选至少两只基金，或逐个选中后点击【加入对比】。")

        def render():
            state["render_job"] = None
            state["query"] = search_var.get()
            rows = filtered_rows()
            draw_types()
            draw_chart(rows)
            render_table(rows)

        def render_later(delay=180):
            job = state.get("render_job")
            if job:
                try:
                    win.after_cancel(job)
                except Exception:
                    pass
            state["render_job"] = win.after(delay, render)

        def switch_sheet(index):
            state["sheet"] = index
            state["sort_col"] = ""
            state["sort_dir"] = -1
            for i, button in enumerate(nav_buttons):
                button.configure(bg=COLORS["panel2"] if i == index else COLORS["side"],
                                 fg=COLORS["text"] if i == index else COLORS["muted"])
            render()

        for i, sheet in enumerate(sheets):
            button = tk.Button(nav_frame, text=f"{sheet['name']}   {len(sheet['rows'])}", anchor="w",
                               command=lambda idx=i: switch_sheet(idx), relief="flat",
                               bg=COLORS["panel2"] if i == 0 else COLORS["side"],
                               fg=COLORS["text"] if i == 0 else COLORS["muted"],
                               activebackground=COLORS["panel2"], activeforeground=COLORS["text"],
                               cursor="hand2", padx=10, pady=9, font=("Microsoft YaHei", 10, "bold"))
            button.pack(fill="x", pady=3)
            nav_buttons.append(button)

        search_var.trace_add("write", lambda *_: render_later())
        def update_limit(*_):
            text = limit_var.get()
            state["limit"] = 0 if text == "全部" else int(text)
            render()
        limit_var.trace_add("write", update_limit)
        render()
        self._log(f"已在软件内打开同花顺风格看板：{path}")

    def _native_ask_open_excel(self, path):
        try:
            self._show_result_visualization(path)
        except Exception as exc:
            self._log(f"软件内看板打开失败：{exc}")

    base_cls._show_result_visualization = _native_show_result
    base_cls._ask_open_excel = _native_ask_open_excel


_install_native_tonghuashun_viewer()


def _install_threadsafe_tk_updates():
    """Route worker-thread UI updates through a Tk polling loop."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    import threading as _threading

    original_build_ui = base_cls._build_ui
    original_on_progress = getattr(base_cls, "_on_progress", None)

    def _ensure_ui_queue(self):
        if not hasattr(self, "_ui_queue_lock"):
            self._ui_queue_lock = _threading.Lock()
            self._ui_log_queue = []
            self._ui_progress_text = None
            self._ui_polling = False

    def _flush_ui_queue(self):
        _ensure_ui_queue(self)
        try:
            logs = []
            progress_text = None
            with self._ui_queue_lock:
                if self._ui_log_queue:
                    logs = self._ui_log_queue[:80]
                    del self._ui_log_queue[:80]
                progress_text = self._ui_progress_text
                self._ui_progress_text = None

            if logs and hasattr(self, "log_text"):
                self.log_text.configure(state="normal")
                self.log_text.insert("end", "".join(logs))
                try:
                    line_count = int(float(self.log_text.index("end-1c").split(".")[0]))
                    if line_count > 1200:
                        self.log_text.delete("1.0", f"{line_count - 800}.0")
                except Exception:
                    pass
                self.log_text.see("end")
                self.log_text.configure(state="disabled")

            if progress_text and hasattr(self, "lbl_progress"):
                self.lbl_progress.config(text=progress_text)
        except Exception:
            pass
        finally:
            try:
                self.root.after(80, lambda: _flush_ui_queue(self))
            except Exception:
                pass

    def _patched_build_ui(self):
        original_build_ui(self)
        _ensure_ui_queue(self)
        if not getattr(self, "_ui_polling", False):
            self._ui_polling = True
            try:
                self.root.after(80, lambda: _flush_ui_queue(self))
            except Exception:
                pass
        try:
            self.root.after(1200, lambda: _ensure_fund_index_async(self._log))
        except Exception:
            pass

    def _safe_log(self, msg):
        _ensure_ui_queue(self)
        ts = _dt.now().strftime("%H:%M")
        line = f"[{ts}] {msg}\n"
        try:
            with self._ui_queue_lock:
                self._ui_log_queue.append(line)
                if len(self._ui_log_queue) > 2000:
                    self._ui_log_queue = self._ui_log_queue[-1000:]
        except Exception:
            try:
                print(line, end="")
            except Exception:
                pass

    def _safe_on_progress(self, ok, total):
        _ensure_ui_queue(self)
        try:
            with self._ui_queue_lock:
                self._ui_progress_text = f"进度：{ok} / {total} 条"
        except Exception:
            if callable(original_on_progress):
                try:
                    original_on_progress(self, ok, total)
                except Exception:
                    pass

    base_cls._build_ui = _patched_build_ui
    base_cls._log = _safe_log
    base_cls._on_progress = _safe_on_progress


_install_threadsafe_tk_updates()


def _install_streaming_performance_score():
    """Make the performance scorer stream fund JSON instead of loading nav_history for all funds."""
    try:
        calc_age = jijin_system.calc_age
        parse_pct_to_float = jijin_system.parse_pct_to_float
        fmt_pct = jijin_system.fmt_pct
        calc_score = jijin_system.calc_score
        autosize_excel = jijin_system.autosize_excel
        pd = jijin_system.pd
        json = jijin_system.json
        dt = jijin_system.dt
    except Exception:
        return

    def _iter_json_objects(path, log=None):
        decoder = json.loads
        buf = []
        in_obj = False
        in_str = False
        escape = False
        depth = 0
        count = 0
        with open(path, "r", encoding="utf-8") as f:
            while True:
                chunk = f.read(1024 * 1024)
                if not chunk:
                    break
                for ch in chunk:
                    if not in_obj:
                        if ch == "{":
                            in_obj = True
                            in_str = False
                            escape = False
                            depth = 1
                            buf = [ch]
                        continue

                    buf.append(ch)
                    if in_str:
                        if escape:
                            escape = False
                        elif ch == "\\":
                            escape = True
                        elif ch == '"':
                            in_str = False
                        continue

                    if ch == '"':
                        in_str = True
                    elif ch in "{[":
                        depth += 1
                    elif ch in "}]":
                        depth -= 1
                        if depth == 0:
                            raw = "".join(buf)
                            buf = []
                            in_obj = False
                            try:
                                yield decoder(raw)
                                count += 1
                            except Exception as exc:
                                if log and count < 3:
                                    log(f"跳过一条解析失败的数据：{exc}")

    def _latest_profile_file():
        files = glob.glob(os.path.join(SCRIPT_DIR, "fund_data", "fund_profile_*.json"))
        files += glob.glob(os.path.join("fund_data", "fund_profile_*.json"))
        files = [
            path for path in set(files)
            if os.path.exists(path) and not path.endswith(".tmp") and os.path.getsize(path) > 1024
        ]
        return max(files, key=os.path.getmtime) if files else None

    def _streaming_run_performance_score(log):
        try:
            output_dir = "fund_excel"
            latest_file = _latest_profile_file()
            if not latest_file:
                log("未找到 fund_data 目录下的有效 fund_profile JSON，请先爬取数据。")
                return None

            log(f"正在处理: {latest_file}")
            log("收益表现采用低内存流式计算：只读取收益字段，不加载全部历史净值。")

            metric_rows = []
            full_rows = []
            loaded = 0
            for item in _iter_json_objects(latest_file, log):
                if not isinstance(item, dict):
                    continue
                perf = item.get("performance", {}) or {}
                base = item.get("base_info", {}) or {}
                code = item.get("fund_code")
                if not code:
                    continue
                name = item.get("fund_name")
                age = calc_age(base.get("setup_date"), perf.get("nav_date"))

                metric_rows.append({
                    "fund_code": code,
                    "return_1m": parse_pct_to_float(perf.get("1m")),
                    "return_3m": parse_pct_to_float(perf.get("3m")),
                    "return_6m": parse_pct_to_float(perf.get("6m")),
                    "return_1y": parse_pct_to_float(perf.get("1y")),
                    "return_3y": parse_pct_to_float(perf.get("3y")),
                    "return_5y": parse_pct_to_float(perf.get("5y")),
                })
                full_rows.append({
                    "fund_code": code,
                    "基金名称": name,
                    "基金代码": code,
                    "最新净值": perf.get("nav"),
                    "日增长率": perf.get("daily_growth_rate"),
                    "近1月": fmt_pct(perf.get("1m")),
                    "近3月": fmt_pct(perf.get("3m")),
                    "近6月": fmt_pct(perf.get("6m")),
                    "近1年": fmt_pct(perf.get("1y")),
                    "近3年": fmt_pct(perf.get("3y")),
                    "近5年": fmt_pct(perf.get("5y")),
                    "成立以来": fmt_pct(perf.get("since")),
                    "基金类型": base.get("fund_type"),
                    "风险等级": base.get("risk_level"),
                    "规模": base.get("assets_size"),
                    "基金经理": base.get("manager"),
                    "成立日期": base.get("setup_date"),
                    "净值日期": perf.get("nav_date"),
                    "成立年限": age,
                })
                loaded += 1
                if loaded % 5000 == 0:
                    log(f"已读取 {loaded:,} 只基金...")

            if not full_rows:
                log("JSON 文件中无有效数据。")
                return None

            log(f"共加载 {loaded:,} 只基金，开始计算收益表现评分...")
            metric_df = pd.DataFrame(metric_rows).set_index("fund_code")
            full_df = pd.DataFrame(full_rows).set_index("fund_code")
            full_df["收益表现评分"] = calc_score(metric_df, full_df["成立年限"])
            out = full_df.reset_index(drop=True).sort_values(by="收益表现评分", ascending=False)

            os.makedirs(output_dir, exist_ok=True)
            ts = dt.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(output_dir, f"收益表现评分_{ts}.xlsx")
            out.to_excel(out_path, index=False)
            autosize_excel(out_path)

            log(f"评分计算完成！共 {len(out):,} 只基金")
            log(f"已导出美化Excel: {out_path}")
            return os.path.abspath(out_path)
        except Exception as exc:
            import traceback
            log(f"收益表现评分出错: {exc}")
            log(traceback.format_exc())
            return None

    jijin_system.run_performance_score = _streaming_run_performance_score


_install_streaming_performance_score()


def _install_streaming_risk_drawdown():
    """Make risk drawdown scoring stream JSON and export only the dashboard sheet."""
    try:
        json = jijin_system.json
        pd = jijin_system.pd
        dt = jijin_system.dt
        calc_risk = jijin_system._calc_risk_metrics_from_history
        beautify = jijin_system._beautify_risk_excel
    except Exception:
        return

    def _iter_json_objects(path):
        buf = []
        in_obj = False
        in_str = False
        escape = False
        depth = 0
        with open(path, "r", encoding="utf-8") as f:
            while True:
                chunk = f.read(1024 * 1024)
                if not chunk:
                    break
                for ch in chunk:
                    if not in_obj:
                        if ch == "{":
                            in_obj = True
                            in_str = False
                            escape = False
                            depth = 1
                            buf = [ch]
                        continue
                    buf.append(ch)
                    if in_str:
                        if escape:
                            escape = False
                        elif ch == "\\":
                            escape = True
                        elif ch == '"':
                            in_str = False
                        continue
                    if ch == '"':
                        in_str = True
                    elif ch in "{[":
                        depth += 1
                    elif ch in "}]":
                        depth -= 1
                        if depth == 0:
                            raw = "".join(buf)
                            buf = []
                            in_obj = False
                            try:
                                yield json.loads(raw)
                            except Exception:
                                continue

    def _latest_profile_file():
        files = glob.glob(os.path.join(SCRIPT_DIR, "fund_data", "fund_profile_*.json"))
        files += glob.glob(os.path.join("fund_data", "fund_profile_*.json"))
        files = [
            path for path in set(files)
            if os.path.exists(path) and not path.endswith(".tmp") and os.path.getsize(path) > 1024
        ]
        return max(files, key=os.path.getmtime) if files else None

    def _streaming_run_risk_drawdown(log):
        try:
            output_dir = "fund_excel"
            latest_file = _latest_profile_file()
            if not latest_file:
                log("未找到 fund_data 目录下的有效 fund_profile JSON，请先运行【开始爬取】。")
                return None

            log(f"正在处理: {latest_file}")
            log("风险回撤采用低内存流式计算：逐只基金计算，不写入几万张历史Sheet。")

            summary_list = []
            total = 0
            has_hist = 0
            computed = 0
            for item in _iter_json_objects(latest_file):
                if not isinstance(item, dict):
                    continue
                total += 1
                code = str(item.get("fund_code", "")).zfill(6)
                name = item.get("fund_name", "--")
                base = item.get("base_info", {}) or {}
                hist = item.get("nav_history") or []

                row = {
                    "代码": code,
                    "名称": name,
                    "类型": base.get("fund_type", "--"),
                    "规模": base.get("assets_size", "--"),
                    "经理": base.get("manager", "--"),
                }
                if hist:
                    has_hist += 1
                metrics, _df_hist = calc_risk(hist)
                if metrics:
                    row.update(metrics)
                    computed += 1
                summary_list.append(row)

                if total % 1000 == 0:
                    log(f"风险回撤进度：已处理 {total:,} 只，成功计算 {computed:,} 只。")

            if total == 0:
                log("JSON 文件中无有效数据。")
                return None
            if has_hist == 0:
                log("当前 JSON 中未包含历史净值（nav_history）。")
                log("请重新点击【开始爬取】以生成含历史净值的数据文件，再运行本功能。")
                return None

            df_raw = pd.DataFrame(summary_list)
            display_cols = [
                "代码", "名称", "夏普比率", "卡玛比率", "索提诺比率",
                "年化收益", "最大回撤", "当前回撤", "回撤状态", "回撤进度",
                "决策建议", "波动率", "下行波动", "溃疡指数",
                "类型", "规模", "经理",
            ]
            df_final = df_raw[[c for c in display_cols if c in df_raw.columns]].copy()
            if "夏普比率" in df_final.columns:
                df_final = df_final.sort_values("夏普比率", ascending=False, na_position="last")

            os.makedirs(output_dir, exist_ok=True)
            ts = dt.now().strftime("%m%d_%H%M")
            out_path = os.path.join(output_dir, f"基金风险决策看板_{ts}.xlsx")
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df_final.to_excel(writer, sheet_name="决策看板", index=False)
            beautify(out_path)

            log(f"风险与回撤分析完成！共 {len(df_final):,} 只基金，成功计算 {computed:,} 只。")
            log("历史走势请在软件内结果看板中双击基金或使用走势对比查看。")
            log(f"已导出美化Excel: {out_path}")
            return os.path.abspath(out_path)
        except Exception as exc:
            import traceback
            log(f"风险与回撤分析出错: {exc}")
            log(traceback.format_exc())
            return None

    jijin_system.run_risk_drawdown = _streaming_run_risk_drawdown


_install_streaming_risk_drawdown()


def _install_market_home_panel():
    """Add a Tonghuashun-like global market strip to the main desktop homepage."""
    try:
        base_cls = jijin_system.FundToolsApp
        original_build_ui = base_cls._build_ui
    except Exception:
        return

    COLORS = {
        "bg": "#1e1e2e",
        "panel": "#181825",
        "card": "#24263a",
        "line": "#313244",
        "text": "#cdd6f4",
        "muted": "#6c7086",
        "red": "#f38ba8",
        "green": "#7fbf9f",
        "gold": "#f9e2af",
    }

    def _parse_sina_line(raw):
        try:
            body = raw.split('="', 1)[1].rsplit('";', 1)[0]
            parts = body.split(",")
            name = parts[0] or "--"
            nums = []
            for part in parts[1:]:
                try:
                    nums.append(float(part))
                except Exception:
                    pass
            if len(nums) >= 3:
                current = nums[2] if nums[2] != 0 else nums[0]
                prev = nums[1] if nums[1] != 0 else nums[0]
                change = current - prev
                pct = change / prev * 100 if prev else 0
                return name, current, change, pct
            if nums:
                return name, nums[0], 0, 0
        except Exception:
            pass
        return None

    def _fetch_market_indices():
        import requests
        codes = [
            ("上证指数", "sh000001"),
            ("深证成指", "sz399001"),
            ("创业板指", "sz399006"),
            ("恒生指数", "int_hangseng"),
            ("道琼斯", "gb_$dji"),
            ("纳斯达克", "gb_ixic"),
            ("标普500", "gb_inx"),
            ("日经225", "b_NKY"),
        ]
        url = "https://hq.sinajs.cn/list=" + ",".join(code for _, code in codes)
        headers = {
            "Referer": "https://finance.sina.com.cn/",
            "User-Agent": "Mozilla/5.0",
        }
        session = requests.Session()
        session.trust_env = False
        resp = session.get(url, headers=headers, timeout=8)
        resp.encoding = "gbk"
        parsed = {}
        for line in resp.text.splitlines():
            if not line.strip():
                continue
            code = line.split("hq_str_", 1)[-1].split("=", 1)[0]
            parsed[code] = _parse_sina_line(line)
        out = []
        for label, code in codes:
            item = parsed.get(code)
            if item:
                name, current, change, pct = item
                out.append({"label": label, "name": name, "value": current, "change": change, "pct": pct})
            else:
                out.append({"label": label, "name": label, "value": None, "change": None, "pct": None})
        return out

    def _build_market_panel(self):
        import threading
        import tkinter as tk
        from datetime import datetime

        parent = self.root
        frame = tk.Frame(parent, bg=COLORS["bg"], padx=20, pady=6)
        before = getattr(self, "log_frame", None)
        if before is not None:
            frame.pack(fill="x", before=before)
        else:
            frame.pack(fill="x")

        top = tk.Frame(frame, bg=COLORS["bg"])
        top.pack(fill="x")
        tk.Label(top, text="全球市场", bg=COLORS["bg"], fg=COLORS["gold"],
                 font=("微软雅黑", 11, "bold")).pack(side="left", padx=(4, 10))
        status = tk.Label(top, text="等待刷新", bg=COLORS["bg"], fg=COLORS["muted"],
                          font=("微软雅黑", 9))
        status.pack(side="left")
        cards = tk.Frame(frame, bg=COLORS["bg"])
        cards.pack(fill="x", pady=(6, 0))
        card_widgets = []
        labels = ["上证指数", "深证成指", "创业板指", "恒生指数", "道琼斯", "纳斯达克", "标普500", "日经225"]
        for label in labels:
            card = tk.Frame(cards, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=5)
            name = tk.Label(card, text=label, bg=COLORS["card"], fg=COLORS["muted"], font=("微软雅黑", 8))
            name.pack(anchor="w", padx=9, pady=(5, 0))
            value = tk.Label(card, text="--", bg=COLORS["card"], fg=COLORS["text"], font=("Consolas", 13, "bold"))
            value.pack(anchor="w", padx=9)
            pct = tk.Label(card, text="--", bg=COLORS["card"], fg=COLORS["muted"], font=("Consolas", 9, "bold"))
            pct.pack(anchor="w", padx=9, pady=(0, 5))
            card_widgets.append((name, value, pct))

        def apply_data(data):
            for widgets, item in zip(card_widgets, data):
                name_w, value_w, pct_w = widgets
                name_w.config(text=item["label"])
                if item["value"] is None:
                    value_w.config(text="--", fg=COLORS["text"])
                    pct_w.config(text="暂无数据", fg=COLORS["muted"])
                    continue
                color = COLORS["red"] if (item["change"] or 0) >= 0 else COLORS["green"]
                sign = "+" if (item["change"] or 0) >= 0 else ""
                value_w.config(text=f"{item['value']:.2f}", fg=color)
                pct_w.config(text=f"{sign}{item['change']:.2f}  {sign}{item['pct']:.2f}%", fg=color)
            status.config(text="更新时间 " + datetime.now().strftime("%H:%M:%S"), fg=COLORS["muted"])

        def refresh():
            status.config(text="正在刷新...", fg=COLORS["gold"])
            def worker():
                try:
                    data = _fetch_market_indices()
                    parent.after(0, lambda: apply_data(data))
                except Exception as exc:
                    msg = str(exc).splitlines()[0][:80]
                    parent.after(0, lambda m=msg: status.config(text=f"刷新失败：{m}", fg=COLORS["red"]))
            threading.Thread(target=worker, daemon=True).start()

        tk.Button(top, text="刷新大盘", command=refresh, bg=COLORS["gold"], fg="#1e1e2e",
                  relief="flat", padx=10, pady=3, cursor="hand2",
                  font=("微软雅黑", 9, "bold")).pack(side="right", padx=8)
        refresh()

    def _patched_build_ui(self):
        original_build_ui(self)
        _build_market_panel(self)

    base_cls._build_ui = _patched_build_ui


_install_market_home_panel()

def _install_fund_flow_panel():
    """Add market/sector fund-flow dashboard to the desktop homepage."""
    try:
        base_cls = jijin_system.FundToolsApp
        original_build_ui = base_cls._build_ui
    except Exception:
        return

    COLORS = {
        "bg": "#1e1e2e",
        "panel": "#181825",
        "card": "#24263a",
        "line": "#313244",
        "text": "#cdd6f4",
        "muted": "#6c7086",
        "red": "#f38ba8",
        "green": "#7fbf9f",
        "gold": "#f9e2af",
        "blue": "#89b4fa",
        "cyan": "#94e2d5",
    }

    def _to_float(value):
        try:
            if value in (None, "", "-", "--"):
                return None
            return float(value)
        except Exception:
            return None

    def _money_yi(value):
        num = _to_float(value)
        if num is None:
            return "--"
        yi = num / 100000000
        sign = "+" if yi > 0 else ""
        return f"{sign}{yi:.2f}亿"

    def _pct(value):
        num = _to_float(value)
        if num is None:
            return "--"
        sign = "+" if num > 0 else ""
        return f"{sign}{num:.2f}%"

    def _fetch_json(url, params):
        import requests
        headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://data.eastmoney.com/"}
        candidates = [url]
        if "push2.eastmoney.com" in url:
            candidates.append(url.replace("push2.eastmoney.com", "push2delay.eastmoney.com"))
        if url.startswith("https://"):
            candidates.append(url.replace("https://", "http://", 1))

        last_error = None
        for candidate in candidates:
            try:
                session = requests.Session()
                session.trust_env = False
                resp = session.get(candidate, params=params, headers=headers, timeout=12)
                resp.encoding = "utf-8"
                resp.raise_for_status()
                return resp.json()
            except Exception as exc:
                last_error = exc
                continue
        raise last_error

    def _fetch_market_flow():
        import time
        url = "https://push2his.eastmoney.com/api/qt/stock/fflow/daykline/get"
        params = {
            "lmt": "0",
            "klt": "101",
            "secid": "1.000001",
            "secid2": "0.399001",
            "fields1": "f1,f2,f3,f7",
            "fields2": "f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61,f62,f63,f64,f65",
            "ut": "b2884a393a59ad64002292a3e90d46a5",
            "_": int(time.time() * 1000),
        }
        data = _fetch_json(url, params).get("data") or {}
        rows = data.get("klines") or []
        if not rows:
            return {}
        parts = rows[-1].split(",")
        keys = [
            "date", "main_net", "small_net", "mid_net", "large_net", "super_net",
            "main_ratio", "small_ratio", "mid_ratio", "large_ratio", "super_ratio",
            "sh_close", "sh_pct", "sz_close", "sz_pct",
        ]
        latest = dict(zip(keys, parts))
        return latest

    def _fetch_ths_flow(kind="industry", limit=120):
        import pandas as pd
        import requests
        from io import StringIO

        base = "hyzjl" if kind == "industry" else "gnzjl"
        referer = f"https://data.10jqka.com.cn/funds/{base}/"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": referer,
            "Accept-Language": "zh-CN,zh;q=0.9",
        }
        session = requests.Session()
        session.trust_env = False
        rows = []
        pages = max(1, min(4, (limit + 49) // 50))
        for page in range(1, pages + 1):
            if page == 1:
                url = referer
            else:
                url = f"https://data.10jqka.com.cn/funds/{base}/field/tradezdf/order/desc/page/{page}/"
            resp = session.get(url, headers=headers, timeout=15)
            resp.encoding = "gbk"
            resp.raise_for_status()
            tables = pd.read_html(StringIO(resp.text))
            if not tables:
                continue
            df = tables[0]
            if df.shape[1] < 7:
                continue
            for item in df.itertuples(index=False, name=None):
                name = str(item[1]).strip()
                if not name or name == "nan":
                    continue
                inflow_yi = _to_float(item[4]) or 0
                outflow_yi = _to_float(item[5]) or 0
                net_yi = _to_float(item[6])
                if net_yi is None:
                    net_yi = inflow_yi - outflow_yi
                total_yi = inflow_yi + outflow_yi
                rows.append({
                    "code": "",
                    "name": name,
                    "price": _to_float(item[2]) if len(item) > 2 else None,
                    "pct": _to_float(str(item[3] if len(item) > 3 else "").replace("%", "")),
                    "main_net": net_yi * 100000000,
                    "main_ratio": net_yi / total_yi * 100 if total_yi else None,
                    "super_net": 0,
                    "large_net": net_yi * 100000000,
                    "mid_net": 0,
                    "small_net": 0,
                    "inflow": inflow_yi * 100000000,
                    "outflow": outflow_yi * 100000000,
                    "inflow_ratio": inflow_yi / total_yi * 100 if total_yi else None,
                    "outflow_ratio": outflow_yi / total_yi * 100 if total_yi else None,
                    "leader": str(item[8]).strip() if len(item) > 8 else "",
                    "company_count": item[7] if len(item) > 7 else "",
                    "source": "同花顺",
                })
                if len(rows) >= limit:
                    return rows
        return rows

    def _fetch_sector_flow(sector_type="industry", limit=12):
        try:
            focus_keywords = ["半导体", "芯片", "人工智能", "新能源", "光伏", "证券", "医药", "机器人", "CPO"]
            ths_rows = _fetch_ths_flow("industry" if sector_type == "industry" else "concept", max(limit, 120))
            selected, seen = [], set()
            for row in ths_rows[:limit]:
                selected.append(row)
                seen.add(row.get("name"))
            for row in ths_rows:
                name = str(row.get("name", ""))
                if name in seen:
                    continue
                if any(keyword in name for keyword in focus_keywords):
                    selected.append(row)
                    seen.add(name)
                if len(selected) >= limit + 8:
                    break
            return selected
        except Exception:
            pass
        import math
        import time
        focus_keywords = ["半导体", "芯片", "人工智能", "新能源", "光伏", "证券", "医药", "机器人"]
        sector_type_map = {"industry": "2", "concept": "3", "region": "1"}
        url = "https://push2delay.eastmoney.com/api/qt/clist/get"
        params = {
            "pn": "1",
            "pz": "120",
            "po": "1",
            "np": "1",
            "ut": "b2884a393a59ad64002292a3e90d46a5",
            "fltt": "2",
            "invt": "2",
            "fid": "f62",
            "fid0": "f62",
            "fs": f"m:90 t:{sector_type_map.get(sector_type, '2')}",
            "stat": "1",
            "fields": "f12,f14,f2,f3,f62,f184,f66,f69,f72,f75,f78,f81,f84,f87,f204,f205,f124",
            "rt": "52975239",
            "_": int(time.time() * 1000),
        }
        data = _fetch_json(url, params).get("data") or {}
        rows = data.get("diff") or []
        out = []
        parsed = []
        for item in rows:
            main_net = _to_float(item.get("f62")) or 0
            super_net = _to_float(item.get("f66")) or 0
            large_net = _to_float(item.get("f72")) or 0
            mid_net = _to_float(item.get("f78")) or 0
            small_net = _to_float(item.get("f84")) or 0
            inflow = sum(x for x in (super_net, large_net, mid_net, small_net) if x > 0)
            outflow = -sum(x for x in (super_net, large_net, mid_net, small_net) if x < 0)
            total = inflow + outflow
            parsed.append({
                "code": item.get("f12", ""),
                "name": item.get("f14", "--"),
                "price": item.get("f2"),
                "pct": item.get("f3"),
                "main_net": main_net,
                "main_ratio": item.get("f184"),
                "super_net": super_net,
                "large_net": large_net,
                "mid_net": mid_net,
                "small_net": small_net,
                "inflow": inflow,
                "outflow": outflow,
                "inflow_ratio": inflow / total * 100 if total else None,
                "outflow_ratio": outflow / total * 100 if total else None,
            })
        selected = []
        seen = set()
        for row in parsed[:limit]:
            selected.append(row)
            seen.add(row["name"])
        for row in parsed:
            name = str(row.get("name", ""))
            if name in seen:
                continue
            if any(keyword in name for keyword in focus_keywords):
                selected.append(row)
                seen.add(name)
            if len(selected) >= limit + 8:
                break
        out = selected
        return out

    def _fetch_flow_data():
        import json
        errors = []
        industry, concept, market = [], [], {}
        source = "同花顺资金流页面直爬"
        fetched_at = _dt.now().strftime("%Y-%m-%d %H:%M:%S")
        cache_used = False
        try:
            industry = _fetch_sector_flow("industry", 12)
        except Exception as exc:
            errors.append(f"行业资金流失败: {str(exc).splitlines()[0][:40]}")
        try:
            concept = _fetch_sector_flow("concept", 12)
        except Exception as exc:
            errors.append(f"概念资金流失败: {str(exc).splitlines()[0][:40]}")
        try:
            market = _fetch_market_flow()
        except Exception as exc:
            errors.append(f"大盘资金流失败: {str(exc).splitlines()[0][:40]}")

        if not market and industry:
            total_main = sum(float(row.get("main_net") or 0) for row in industry)
            total_super = sum(float(row.get("super_net") or 0) for row in industry)
            total_large = sum(float(row.get("large_net") or 0) for row in industry)
            total_mid = sum(float(row.get("mid_net") or 0) for row in industry)
            total_small = sum(float(row.get("small_net") or 0) for row in industry)
            market = {
                "main_net": total_main,
                "super_net": total_super,
                "large_net": total_large,
                "mid_net": total_mid,
                "small_net": total_small,
                "main_ratio": None,
                "super_ratio": None,
                "large_ratio": None,
                "mid_ratio": None,
                "small_ratio": None,
                "fallback": True,
            }
        if industry or concept:
            try:
                os.makedirs(os.path.join(SCRIPT_DIR, "fund_cache"), exist_ok=True)
                cache_path = os.path.join(SCRIPT_DIR, "fund_cache", "fund_flow_latest.json")
                def _cache_row(row):
                    return {
                        "name": row.get("name", ""),
                        "pct": row.get("pct"),
                        "inflow_yi": (_to_float(row.get("inflow")) or 0) / 100000000,
                        "outflow_yi": (_to_float(row.get("outflow")) or 0) / 100000000,
                        "net_yi": (_to_float(row.get("main_net")) or 0) / 100000000,
                        "net_ratio": row.get("main_ratio"),
                        "inflow_ratio": row.get("inflow_ratio"),
                        "leader": row.get("leader", ""),
                        "company_count": row.get("company_count", ""),
                    }
                with open(cache_path, "w", encoding="utf-8") as f:
                    json.dump({
                        "source": source,
                        "fetched_at": fetched_at,
                        "industry": [_cache_row(row) for row in industry],
                        "concept": [_cache_row(row) for row in concept],
                    }, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        if not industry and not concept:
            try:
                cache_path = os.path.join(SCRIPT_DIR, "fund_cache", "fund_flow_latest.json")
                with open(cache_path, "r", encoding="utf-8") as f:
                    cached = json.load(f)
                def _from_cache(row):
                    inflow = (_to_float(row.get("inflow_yi")) or 0) * 100000000
                    outflow = (_to_float(row.get("outflow_yi")) or 0) * 100000000
                    net = (_to_float(row.get("net_yi")) or 0) * 100000000
                    return {
                        "name": row.get("name", "--"),
                        "pct": row.get("pct"),
                        "main_net": net,
                        "main_ratio": row.get("net_ratio"),
                        "inflow": inflow,
                        "outflow": outflow,
                        "inflow_ratio": row.get("inflow_ratio"),
                        "source": "本地缓存",
                    }
                industry = [_from_cache(row) for row in (cached.get("industry") or [])[:20]]
                concept = [_from_cache(row) for row in (cached.get("concept") or [])[:20]]
                source = cached.get("source") or "本地缓存"
                fetched_at = cached.get("fetched_at") or fetched_at
                cache_used = True
                errors.append(f"实时源失败，已读取本地缓存 {fetched_at}")
                if not market and industry:
                    total_main = sum(float(row.get("main_net") or 0) for row in industry)
                    market = {"main_net": total_main, "super_net": None, "large_net": None,
                              "mid_net": None, "small_net": None, "fallback": True}
            except Exception as exc:
                errors.append(f"缓存读取失败: {str(exc).splitlines()[0][:40]}")
        return {
            "market": market,
            "industry": industry,
            "concept": concept,
            "errors": errors,
            "source": source,
            "fetched_at": fetched_at,
            "cache_used": cache_used,
        }

    def _build_fund_flow_panel(self):
        import threading
        import tkinter as tk
        from tkinter import ttk
        from datetime import datetime

        parent = self.root
        frame = tk.Frame(parent, bg=COLORS["bg"], padx=20, pady=6)
        before = getattr(self, "log_frame", None)
        if before is not None:
            frame.pack(fill="x", before=before)
        else:
            frame.pack(fill="x")

        top = tk.Frame(frame, bg=COLORS["bg"])
        top.pack(fill="x")
        tk.Label(top, text="资金流向", bg=COLORS["bg"], fg=COLORS["gold"],
                 font=("Microsoft YaHei", 12, "bold")).pack(side="left", padx=(4, 10))
        status = tk.Label(top, text="等待刷新", bg=COLORS["bg"], fg=COLORS["muted"],
                          font=("Microsoft YaHei", 9))
        status.pack(side="left")

        body = tk.Frame(frame, bg=COLORS["bg"])
        body.pack(fill="x", pady=(6, 0))

        market = tk.Frame(body, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        market.pack(side="left", fill="both", expand=True, padx=(0, 10))
        tk.Label(market, text="大方向资金", bg=COLORS["panel"], fg=COLORS["text"],
                 font=("Microsoft YaHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 2))
        market_grid = tk.Frame(market, bg=COLORS["panel"])
        market_grid.pack(fill="x", padx=10, pady=(0, 8))
        market_cards = {}

        def add_market_card(key, label):
            card = tk.Frame(market_grid, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 7), ipady=4)
            tk.Label(card, text=label, bg=COLORS["card"], fg=COLORS["muted"], font=("Microsoft YaHei", 8)).pack(anchor="w", padx=8, pady=(5, 0))
            val = tk.Label(card, text="--", bg=COLORS["card"], fg=COLORS["text"], font=("Consolas", 12, "bold"))
            val.pack(anchor="w", padx=8)
            pct = tk.Label(card, text="--", bg=COLORS["card"], fg=COLORS["muted"], font=("Consolas", 9, "bold"))
            pct.pack(anchor="w", padx=8, pady=(0, 5))
            market_cards[key] = (val, pct)

        for key, label in [
            ("main", "主力"),
            ("super", "超大单"),
            ("large", "大单"),
            ("mid", "中单"),
            ("small", "小单"),
        ]:
            add_market_card(key, label)

        right = tk.Frame(body, bg=COLORS["bg"])
        right.pack(side="left", fill="both", expand=True)

        def make_table(parent_frame, title):
            panel = tk.Frame(parent_frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
            panel.pack(side="left", fill="both", expand=True, padx=(0, 10))
            tk.Label(panel, text=title, bg=COLORS["panel"], fg=COLORS["text"],
                     font=("Microsoft YaHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 4))
            cols = ["名称", "涨跌", "主力净流入", "净占比", "流入", "流出", "流入占比"]
            tree = ttk.Treeview(panel, columns=cols, show="headings", height=8)
            tree.tag_configure("pos", foreground=COLORS["red"])
            tree.tag_configure("neg", foreground=COLORS["green"])
            for col in cols:
                tree.heading(col, text=col)
                width = 76
                if col == "名称":
                    width = 108
                elif col in ("主力净流入", "流入占比"):
                    width = 88
                tree.column(col, width=width, anchor="center", stretch=False)
            tree.pack(fill="both", expand=True, padx=8, pady=(0, 8))
            return tree

        industry_tree = make_table(right, "行业资金流")
        concept_tree = make_table(right, "概念资金流")

        def fill_market(data):
            mapping = {
                "main": ("main_net", "main_ratio"),
                "super": ("super_net", "super_ratio"),
                "large": ("large_net", "large_ratio"),
                "mid": ("mid_net", "mid_ratio"),
                "small": ("small_net", "small_ratio"),
            }
            for key, (net_key, ratio_key) in mapping.items():
                val, pct = market_cards[key]
                num = _to_float(data.get(net_key))
                color = COLORS["red"] if (num or 0) >= 0 else COLORS["green"]
                val.config(text=_money_yi(data.get(net_key)), fg=color)
                pct.config(text=_pct(data.get(ratio_key)), fg=color)

        def fill_tree(tree, rows):
            tree.delete(*tree.get_children())
            if not rows:
                tree.insert("", "end", values=["暂无数据", "--", "--", "--", "--", "--", "--"])
                return
            for row in rows:
                net = row.get("main_net") or 0
                values = [
                    row.get("name", "--"),
                    _pct(row.get("pct")),
                    _money_yi(net),
                    _pct(row.get("main_ratio")),
                    _money_yi(row.get("inflow")),
                    _money_yi(row.get("outflow")),
                    _pct(row.get("inflow_ratio")),
                ]
                tree.insert("", "end", values=values, tags=("pos" if net >= 0 else "neg",))

        def apply_data(data):
            fill_market(data.get("market") or {})
            fill_tree(industry_tree, data.get("industry") or [])
            fill_tree(concept_tree, data.get("concept") or [])
            errors = data.get("errors") or []
            fetched_at = str(data.get("fetched_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            short_time = fetched_at[-8:] if len(fetched_at) >= 8 else datetime.now().strftime("%H:%M:%S")
            if data.get("cache_used"):
                status.config(text=f"实时源失败，已显示缓存 {short_time}", fg=COLORS["gold"])
            elif errors:
                status.config(text=f"部分数据源失败，已显示可用数据 {datetime.now().strftime('%H:%M:%S')}", fg=COLORS["gold"])
            else:
                status.config(text=f"{data.get('source', '资金流')} 更新 {short_time}", fg=COLORS["muted"])

        def refresh():
            status.config(text="正在刷新...", fg=COLORS["gold"])
            def worker():
                try:
                    data = _fetch_flow_data()
                    parent.after(0, lambda: apply_data(data))
                except Exception as exc:
                    msg = str(exc).splitlines()[0][:80]
                    parent.after(0, lambda m=msg: status.config(text=f"刷新失败：{m}", fg=COLORS["red"]))
            threading.Thread(target=worker, daemon=True).start()

        tk.Button(top, text="刷新资金流", command=refresh, bg=COLORS["gold"], fg="#1e1e2e",
                  relief="flat", padx=10, pady=3, cursor="hand2",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="right", padx=8)
        refresh()

    def _patched_build_ui(self):
        original_build_ui(self)
        _build_fund_flow_panel(self)

    base_cls._build_ui = _patched_build_ui


_install_fund_flow_panel()


def _install_portfolio_module():
    """Add a local 'My Portfolio' module with import and real-time P/L calculation."""
    try:
        base_cls = jijin_system.FundToolsApp
        original_build_ui = base_cls._build_ui
    except Exception:
        return

    def _portfolio_path():
        return os.path.join(SCRIPT_DIR, "my_portfolio.json")

    def _read_holdings_from_file(path):
        import pandas as pd
        ext = os.path.splitext(path)[1].lower()
        df = pd.read_excel(path) if ext in (".xlsx", ".xls") else pd.read_csv(path, encoding="utf-8-sig")
        df.columns = [str(c).strip() for c in df.columns]

        def pick(keys):
            for key in keys:
                for col in df.columns:
                    if key in str(col):
                        return col
            return None

        code_col = pick(["基金代码", "代码", "fund_code"])
        name_col = pick(["基金名称", "名称", "fund_name"])
        shares_col = pick(["份额", "持有份额", "shares"])
        cost_col = pick(["持仓成本", "成本价", "买入净值", "cost_nav"])
        amount_col = pick(["投入金额", "本金", "持仓成本金额", "amount", "cost_amount"])
        if not code_col:
            raise ValueError("导入文件必须包含“基金代码”列")
        holdings = []
        for _, row in df.iterrows():
            code = "".join(ch for ch in str(row.get(code_col, "")) if ch.isdigit()).zfill(6)
            if not code or code == "000000":
                continue
            shares = _num(row.get(shares_col)) if shares_col else None
            cost_nav = _num(row.get(cost_col)) if cost_col else None
            cost_amount = _num(row.get(amount_col)) if amount_col else None
            if shares is None and cost_amount is not None and cost_nav:
                shares = cost_amount / cost_nav
            if cost_amount is None and shares is not None and cost_nav is not None:
                cost_amount = shares * cost_nav
            holdings.append({
                "code": code,
                "name": str(row.get(name_col, "")).strip() if name_col else "",
                "shares": shares or 0,
                "cost_nav": cost_nav,
                "cost_amount": cost_amount or 0,
            })
        return holdings

    def _save_portfolio(holdings):
        import json
        with open(_portfolio_path(), "w", encoding="utf-8") as f:
            json.dump({"holdings": holdings, "updated_at": _dt.now().strftime("%Y-%m-%d %H:%M:%S")},
                      f, ensure_ascii=False, indent=2)

    def _load_portfolio():
        import json
        try:
            with open(_portfolio_path(), "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("holdings", [])
        except Exception:
            return []

    def _calc_portfolio_rows():
        _ensure_fund_index_async()
        rows, total_cost, total_value, daily_profit = [], 0.0, 0.0, 0.0
        for h in _load_portfolio():
            code = str(h.get("code", "")).zfill(6)
            fund = _get_fund_from_index(code) or {}
            perf = fund.get("performance", {}) or {}
            latest_nav = _num(perf.get("nav"))
            daily_pct = _num(perf.get("daily_growth_rate")) or 0.0
            shares = float(h.get("shares") or 0)
            cost_amount = float(h.get("cost_amount") or 0)
            cost_nav = h.get("cost_nav")
            if cost_amount <= 0 and shares and cost_nav:
                cost_amount = shares * float(cost_nav)
            value = shares * latest_nav if latest_nav is not None else 0.0
            profit = value - cost_amount
            profit_pct = profit / cost_amount * 100 if cost_amount else 0.0
            day_profit = value * daily_pct / 100.0
            total_cost += cost_amount
            total_value += value
            daily_profit += day_profit
            rows.append({
                "基金代码": code,
                "基金名称": fund.get("fund_name") or h.get("name") or code,
                "持有份额": shares,
                "成本净值": cost_nav or "--",
                "最新净值": latest_nav if latest_nav is not None else "--",
                "持仓成本": cost_amount,
                "持仓市值": value,
                "持仓盈亏": profit,
                "收益率": profit_pct,
                "今日涨幅": daily_pct,
                "今日盈亏": day_profit,
            })
        return rows, {
            "cost": total_cost,
            "value": total_value,
            "profit": total_value - total_cost,
            "profit_pct": (total_value - total_cost) / total_cost * 100 if total_cost else 0.0,
            "daily_profit": daily_profit,
        }

    def _open_portfolio_window(self):
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        win = tk.Toplevel(self.root)
        win.title("我的持仓")
        win.geometry("1280x760")
        win.configure(bg="#080d18")
        header = tk.Frame(win, bg="#080d18")
        header.pack(fill="x", padx=18, pady=(16, 8))
        tk.Label(header, text="我的持仓", bg="#080d18", fg="#e8edf6",
                 font=("Microsoft YaHei", 22, "bold")).pack(side="left")

        cards = tk.Frame(win, bg="#080d18")
        cards.pack(fill="x", padx=18, pady=(0, 12))
        card_labels = []
        for label in ["持仓成本", "持仓市值", "持仓盈亏", "收益率", "今日盈亏"]:
            card = tk.Frame(cards, bg="#101827", highlightbackground="#263247", highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=8)
            tk.Label(card, text=label, bg="#101827", fg="#8f9bb2", font=("Microsoft YaHei", 9)).pack(anchor="w", padx=14, pady=(8, 2))
            val = tk.Label(card, text="--", bg="#101827", fg="#e8edf6", font=("Consolas", 18, "bold"))
            val.pack(anchor="w", padx=14, pady=(0, 8))
            card_labels.append(val)

        body = tk.Frame(win, bg="#101827", highlightbackground="#263247", highlightthickness=1)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        cols = ["基金代码", "基金名称", "持有份额", "成本净值", "最新净值", "持仓成本", "持仓市值", "持仓盈亏", "收益率", "今日涨幅", "今日盈亏"]
        tree = ttk.Treeview(body, columns=cols, show="headings")
        tree.tag_configure("pos", foreground="#e8553d")
        tree.tag_configure("neg", foreground="#6fb894")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=220 if col == "基金名称" else 110, anchor="center", stretch=False)
        vsb = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        vsb.grid(row=0, column=1, sticky="ns", pady=10)
        hsb.grid(row=1, column=0, sticky="ew", padx=10)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        def money(v):
            return f"{v:,.2f}"

        def refresh():
            rows, summary = _calc_portfolio_rows()
            values = [money(summary["cost"]), money(summary["value"]), money(summary["profit"]),
                      f"{summary['profit_pct']:.2f}%", money(summary["daily_profit"])]
            colors = ["#e8edf6", "#e8edf6",
                      "#e8553d" if summary["profit"] >= 0 else "#6fb894",
                      "#e8553d" if summary["profit_pct"] >= 0 else "#6fb894",
                      "#e8553d" if summary["daily_profit"] >= 0 else "#6fb894"]
            for lab, val, color in zip(card_labels, values, colors):
                lab.config(text=val, fg=color)
            tree.delete(*tree.get_children())
            for row in rows:
                vals = []
                for col in cols:
                    value = row.get(col, "")
                    if isinstance(value, float):
                        value = f"{value:.2f}%" if col in ("收益率", "今日涨幅") else f"{value:,.2f}"
                    vals.append(value)
                tree.insert("", "end", values=vals, tags=("pos" if row.get("持仓盈亏", 0) >= 0 else "neg",))
            if not rows:
                self._log("暂无持仓。导入文件列建议包含：基金代码、基金名称、持有份额、成本净值或投入金额。")

        def import_file():
            path = filedialog.askopenfilename(title="导入持仓文件", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
            if not path:
                return
            try:
                holdings = _read_holdings_from_file(path)
                _save_portfolio(holdings)
                self._log(f"已导入持仓 {len(holdings)} 条：{path}")
                refresh()
            except Exception as exc:
                messagebox.showerror("导入失败", str(exc))

        tk.Button(header, text="导入持仓", command=import_file, bg="#e8b830", fg="#111827",
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right")
        tk.Button(header, text="刷新收益", command=refresh, bg="#3b82f6", fg="#ffffff",
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold")).pack(side="right", padx=8)
        refresh()

    def _patched_build_ui(self):
        original_build_ui(self)
        try:
            import tkinter as tk
            if hasattr(self, "log_frame"):
                bar = tk.Frame(self.root, bg="#1e1e2e", padx=20, pady=4)
                bar.pack(fill="x", before=self.log_frame)
                tk.Button(bar, text="我的持仓", command=lambda: _open_portfolio_window(self),
                          bg="#e8b830", fg="#1e1e2e", relief="flat", padx=16, pady=6,
                          cursor="hand2", font=("微软雅黑", 10, "bold")).pack(side="left", padx=4)
        except Exception:
            pass

    base_cls._build_ui = _patched_build_ui


_install_portfolio_module()

def _install_precompute_cache():
    """Precompute module result files during one-click runs and reuse them on button clicks."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    import json
    import os
    import threading
    import time
    from concurrent.futures import ThreadPoolExecutor
    from datetime import datetime as _dt

    CACHE_DIR = "fund_cache"
    CACHE_INDEX = os.path.join(CACHE_DIR, "module_results.json")
    MODULE_MAX_WORKERS = int(os.environ.get("FUND_MODULE_WORKERS", "2"))
    MODULE_MAX_WORKERS = max(1, min(4, MODULE_MAX_WORKERS))
    MODULE_EXECUTOR = ThreadPoolExecutor(max_workers=MODULE_MAX_WORKERS, thread_name_prefix="fund-module")
    MODULE_STATE_LOCK = threading.Lock()

    def _ensure_cache_dir():
        os.makedirs(CACHE_DIR, exist_ok=True)

    def _load_index():
        try:
            with open(CACHE_INDEX, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_index(index):
        _ensure_cache_dir()
        tmp_path = CACHE_INDEX + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, CACHE_INDEX)

    def _latest_data_stamp():
        files = [
            path for path in glob.glob(os.path.join("fund_data", "fund_profile_*.json"))
            if not path.endswith(".tmp") and os.path.getsize(path) > 1024
        ]
        if not files:
            return ""
        latest = max(files, key=os.path.getmtime)
        return f"{os.path.abspath(latest)}|{os.path.getmtime(latest)}"

    def _cache_get(key):
        item = _load_index().get(key) or {}
        path = item.get("path")
        if not path or not os.path.exists(path):
            return None
        if item.get("data_stamp") != _latest_data_stamp():
            return None
        return path

    def _cache_set(key, path, title):
        if not path:
            return None
        index = _load_index()
        index[key] = {
            "title": title,
            "path": os.path.abspath(path),
            "data_stamp": _latest_data_stamp(),
            "updated_at": _dt.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        _save_index(index)
        return index[key]["path"]

    module_jobs = [
        ("performance", "收益表现", jijin_system.run_performance_score),
        ("risk", "风险回撤", jijin_system.run_risk_drawdown),
        ("efficiency", "风险效率", jijin_system.run_efficiency_score),
        ("position", "位置估值", jijin_system.run_position_score),
        ("timing", "趋势择时", jijin_system.run_timing_score),
        ("manager", "基金经理", jijin_system.run_manager_score),
        ("cost", "交易成本", jijin_system.run_cost_score),
        ("attribution", "收益归因", jijin_system.run_attribution_score),
        ("composite", "长期综合", jijin_system.run_composite_score),
        ("drawdown_shock", "回撤震荡", jijin_system.run_drawdown_shock_screen),
        ("trend_breakout", "趋势突破", jijin_system.run_trend_breakout_screen),
        ("low_vol_stable", "低波稳健", jijin_system.run_low_vol_stable_screen),
        ("oversold_rebound", "超跌反弹", jijin_system.run_oversold_rebound_screen),
    ]

    def _all_topic_names():
        specs = getattr(jijin_system, "TOPIC_SPECS", {}) or {}
        return list(specs.keys())

    def _run_job_with_cache(key, title, func, log, *, force=False):
        if not force:
            cached = _cache_get(key)
            if cached:
                log(f"{title} 已有预计算缓存，直接打开。")
                return cached
        log(f"开始预计算/刷新：{title}")
        path = func(log)
        if path:
            cached_path = _cache_set(key, path, title)
            log(f"{title} 已缓存：{cached_path}")
            return cached_path
        log(f"{title} 未生成结果。")
        return None

    def _run_job_with_cache_process(key, title, log, *, force=False):
        import subprocess

        if not force:
            cached = _cache_get(key)
            if cached:
                log(f"{title} 已有预计算缓存，直接打开。")
                return cached

        log(f"{title} 将在独立进程中计算，主界面可继续操作。")
        args = [sys.executable, os.path.abspath(__file__), "--run-module", key]
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        proc = subprocess.Popen(
            args,
            cwd=SCRIPT_DIR,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        result_path = None
        line_count = 0
        important_words = ("开始", "完成", "出错", "失败", "已保存", "已缓存", "共加载", "筛选", "处理", "Traceback")
        for raw_line in proc.stdout or []:
            line = raw_line.rstrip()
            if not line:
                continue
            if line.startswith("__RESULT_PATH__="):
                result_path = line.split("=", 1)[1].strip() or None
                continue
            line_count += 1
            if any(word in line for word in important_words) or line_count % 120 == 0:
                log(f"{title}｜{line}")
        code = proc.wait()
        if code != 0:
            log(f"{title} 独立进程退出码 {code}。")
        if result_path and os.path.exists(result_path):
            cached_path = _cache_set(key, result_path, title)
            log(f"{title} 已缓存：{cached_path}")
            return cached_path
        log(f"{title} 未生成结果。")
        return None

    def _ensure_module_state(self):
        if not hasattr(self, "_module_jobs_running"):
            self._module_jobs_running = set()
            self._module_jobs_done = 0

    def _module_status_text(self):
        _ensure_module_state(self)
        count = len(self._module_jobs_running)
        return "就绪" if count <= 0 else f"后台计算 {count} 个模块"

    def _mark_module_start(self, key, title):
        _ensure_module_state(self)
        with MODULE_STATE_LOCK:
            if key in self._module_jobs_running:
                return False
            self._module_jobs_running.add(key)
        try:
            self.root.after(0, lambda: self._update_status(_module_status_text(self), "#f9e2af"))
        except Exception:
            pass
        self._log(f"{title} 已加入后台任务；当前最多并发 {MODULE_MAX_WORKERS} 个模块。")
        return True

    def _mark_module_done(self, key):
        _ensure_module_state(self)
        with MODULE_STATE_LOCK:
            self._module_jobs_running.discard(key)
            self._module_jobs_done = getattr(self, "_module_jobs_done", 0) + 1
        try:
            color = "#6c7086" if not self._module_jobs_running else "#f9e2af"
            self.root.after(0, lambda: self._update_status(_module_status_text(self), color))
        except Exception:
            pass

    def _open_cached_or_compute(self, key, title, func):
        cached = _cache_get(key)
        if cached:
            self._log(f"{title} 使用一键运行缓存，马上加载。")
            self.root.after(0, lambda p=cached: self._ask_open_excel(p))
            return

        if not _mark_module_start(self, key, title):
            self._log(f"{title} 已在后台计算中，请勿重复点击同一模块。")
            return
        self._log(f"{title} 没有可用缓存，正在重新计算；完成后会自动弹窗。")

        def _task():
            try:
                path = _run_job_with_cache_process(key, title, self._log, force=True)
                if path:
                    self.root.after(0, lambda p=path: self._ask_open_excel(p))
                else:
                    self._log(f"{title} 没有生成可展示结果，请查看上方运行日志。")
            except Exception as exc:
                self._log(f"{title} 计算失败：{exc}")
            finally:
                _mark_module_done(self, key)
        MODULE_EXECUTOR.submit(_task)

    def _run_topic_cached(self, topic_name):
        key = f"topic::{topic_name}"
        title = f"{topic_name}专题"
        cached = _cache_get(key)
        if cached:
            self._log(f"{title} 使用一键运行缓存，马上加载。")
            self.root.after(0, lambda p=cached: self._ask_open_excel(p))
            return

        if not _mark_module_start(self, key, title):
            self._log(f"{title} 已在后台计算中，请勿重复点击同一专题。")
            return
        self._log(f"{title} 没有可用缓存，正在重新计算；完成后会自动弹窗。")

        def _task():
            try:
                path = _run_job_with_cache_process(key, title, self._log, force=True)
                if path:
                    self.root.after(0, lambda p=path: self._ask_open_excel(p))
                else:
                    self._log(f"{title} 没有生成可展示结果，请查看上方运行日志。")
            except Exception as exc:
                self._log(f"{title} 计算失败：{exc}")
            finally:
                _mark_module_done(self, key)
        MODULE_EXECUTOR.submit(_task)

    def _precompute_all(self, log):
        _ensure_cache_dir()
        started = time.time()
        log("\n[缓存] 开始预计算所有模块，后续点击按钮将直接加载。")
        done, failed = 0, 0
        for key, title, func in module_jobs:
            try:
                if _run_job_with_cache(key, title, func, log, force=True):
                    done += 1
                else:
                    failed += 1
            except Exception as exc:
                failed += 1
                log(f"{title} 预计算失败：{exc}")

        topics = _all_topic_names()
        log(f"[缓存] 开始预计算专题模块，共 {len(topics)} 个。")
        for topic in topics:
            title = f"{topic}专题"
            try:
                if _run_job_with_cache(
                    f"topic::{topic}",
                    title,
                    lambda lg, t=topic: jijin_system.run_topic_screen(t, lg),
                    log,
                    force=True,
                ):
                    done += 1
                else:
                    failed += 1
            except Exception as exc:
                failed += 1
                log(f"{title} 预计算失败：{exc}")
        log(f"[缓存] 预计算完成：成功 {done} 个，失败 {failed} 个，用时 {time.time() - started:.1f}s。")

    def _run_performance(self):
        _open_cached_or_compute(self, "performance", "收益表现", jijin_system.run_performance_score)

    def _run_risk(self):
        _open_cached_or_compute(self, "risk", "风险回撤", jijin_system.run_risk_drawdown)

    def _run_efficiency(self):
        _open_cached_or_compute(self, "efficiency", "风险效率", jijin_system.run_efficiency_score)

    def _run_position(self):
        _open_cached_or_compute(self, "position", "位置估值", jijin_system.run_position_score)

    def _run_timing(self):
        _open_cached_or_compute(self, "timing", "趋势择时", jijin_system.run_timing_score)

    def _run_manager(self):
        _open_cached_or_compute(self, "manager", "基金经理", jijin_system.run_manager_score)

    def _run_cost(self):
        _open_cached_or_compute(self, "cost", "交易成本", jijin_system.run_cost_score)

    def _run_attribution(self):
        _open_cached_or_compute(self, "attribution", "收益归因", jijin_system.run_attribution_score)

    def _run_composite(self):
        _open_cached_or_compute(self, "composite", "长期综合", jijin_system.run_composite_score)

    def _run_drawdown_shock(self):
        _open_cached_or_compute(self, "drawdown_shock", "回撤震荡", jijin_system.run_drawdown_shock_screen)

    def _run_trend_breakout(self):
        _open_cached_or_compute(self, "trend_breakout", "趋势突破", jijin_system.run_trend_breakout_screen)

    def _run_low_vol_stable(self):
        _open_cached_or_compute(self, "low_vol_stable", "低波稳健", jijin_system.run_low_vol_stable_screen)

    def _run_oversold_rebound(self):
        _open_cached_or_compute(self, "oversold_rebound", "超跌反弹", jijin_system.run_oversold_rebound_screen)

    def _run_scoring_job(self, func):
        name = getattr(func, "__name__", "模块")
        key = f"custom::{name}"
        _open_cached_or_compute(self, key, name, func)

    def _run_topic(self, topic_name):
        _run_topic_cached(self, topic_name)

    def _run_one_click(self):
        import glob
        from tkinter import messagebox
        if self._crawling:
            self._log("爬取正在进行中，请先停止再使用一键运行。")
            return

        has_pool = os.path.exists("target_funds.json")
        has_data = bool(glob.glob(os.path.join("fund_data", "*.json")))

        do_clean = False
        do_crawl = True
        do_backfill_history = False
        if has_pool:
            do_clean = messagebox.askyesno(
                "一键运行",
                "检测到已有 target_funds.json。\n\n"
                "是否重新运行【智能清洗】以刷新基金池？\n"
                "（选【否】将沿用现有 target_funds.json）"
            )
        else:
            do_clean = True

        if has_data:
            do_crawl = messagebox.askyesno(
                "一键运行",
                "检测到 fund_data 中已有爬取结果。\n\n"
                "是否重新运行【开始爬取】以刷新行情数据？\n"
                "（选【否】将直接在现有数据上预计算全部模块）"
            )
        else:
            do_crawl = True

        if has_data and do_crawl:
            do_backfill_history = messagebox.askyesno(
                "历史净值",
                "是否补全每只基金的完整历史净值？\n\n"
                "选【是】：首次会很慢，但可补全完整走势图。\n"
                "选【否】：日常迭代更新，只抓最新净值并合并旧数据，速度更快。\n\n"
                "建议：第一次补全选【是】，以后每天选【否】。"
            )

        confirm = messagebox.askyesno(
            "一键运行",
            f"将按以下步骤自动执行：\n\n"
            f"1. 智能清洗 : {'是' if do_clean else '跳过（沿用现有基金池）'}\n"
            f"2. 开始爬取 : {'是' if do_crawl else '跳过（复用现有数据）'}\n"
            f"3. 历史净值补全 : {'是（首次较慢）' if do_backfill_history else '否（只做日常增量）'}\n"
            f"4. 全部模块预计算 : 是（收益、风险、效率、位置、经理、成本、综合策略、所有专题）\n\n"
            f"完成后，各模块按钮会优先直接加载缓存结果。\n"
            f"预计耗时会比原来更长，但后面点击会快很多。\n是否继续？"
        )
        if not confirm:
            return

        self.ctrl_frame.pack(fill="x", padx=30, pady=(0, 6), before=self.log_frame)
        self._update_status("一键运行中", "#f9e2af")

        def _pipeline():
            latest_path = None
            try:
                self._log("=" * 60)
                self._log("一键运行：开始")
                self._log("=" * 60)
                t_all = time.time()

                if do_clean:
                    self._log("\n[1/4] 智能清洗：获取并筛选基金池")
                    jijin_system.run_clean_list(self._log)
                else:
                    self._log("\n[1/4] 智能清洗：跳过，沿用现有 target_funds.json")

                if not os.path.exists("target_funds.json"):
                    self._log("target_funds.json 不存在，流水线终止。")
                    return

                if do_crawl:
                    self._log("\n[2/4] 开始爬取：并发抓取 profile + 历史净值")
                    jijin_system.controller.reset()
                    self._crawling = True
                    self._paused = False
                    self.root.after(0, lambda: self._set_crawl_ui(running=True))
                    old_backfill_env = os.environ.get("FUND_BACKFILL_HISTORY")
                    os.environ["FUND_BACKFILL_HISTORY"] = "1" if do_backfill_history else "0"
                    try:
                        jijin_system.run_crawler(
                            log=self._log,
                            on_progress=self._on_progress,
                            on_done=None,
                        )
                    finally:
                        if old_backfill_env is None:
                            os.environ.pop("FUND_BACKFILL_HISTORY", None)
                        else:
                            os.environ["FUND_BACKFILL_HISTORY"] = old_backfill_env
                        self._crawling = False
                        self.root.after(0, lambda: self._set_crawl_ui(running=False))
                        self.root.after(0, lambda: self.lbl_progress.config(text=""))
                else:
                    self._log("\n[2/4] 开始爬取：跳过，复用现有 fund_data")

                self._log("\n[3/4] 全部模块预计算")
                _precompute_all(self, self._log)

                latest_path = _cache_get("composite") or _cache_get("performance")
                elapsed = time.time() - t_all
                self._log("\n" + "=" * 60)
                self._log(f"一键运行完成！总耗时 {elapsed:.1f}s")
                self._log("=" * 60)

                if latest_path:
                    self.root.after(0, lambda p=latest_path: self._ask_open_excel(p))

                if self._auto_shutdown.get():
                    self._log("自动关机选项已启用，60 秒后关机...")
                    self._log("如需取消，请在 60 秒内执行：shutdown /a")
                    try:
                        if sys.platform == "win32":
                            os.system('shutdown /s /t 60 /c "基金数据工具-一键运行已完成，系统即将关机"')
                        else:
                            os.system("shutdown -h +1")
                    except Exception as e2:
                        self._log(f"关机命令执行失败: {e2}")
            except Exception as e:
                import traceback
                self._log(f"一键运行出错: {e}")
                self._log(traceback.format_exc())
            finally:
                self.root.after(0, lambda: self._update_status("就绪", "#6c7086"))

        threading.Thread(target=_pipeline, daemon=True).start()

    def _scheduled_run(self):
        self._log("=" * 60)
        self._log("定时任务触发：开始自动刷新数据并预计算全部模块")
        self._log("=" * 60)
        t_all = time.time()
        try:
            if not os.path.exists("target_funds.json"):
                self._log("\n[定时 1/3] 智能清洗")
                jijin_system.run_clean_list(self._log)
            else:
                self._log("\n[定时 1/3] 智能清洗：跳过，沿用现有 target_funds.json")
            if not os.path.exists("target_funds.json"):
                self._log("target_funds.json 不存在，定时任务终止。")
                return

            self._log("\n[定时 2/3] 开始爬取")
            jijin_system.controller.reset()
            jijin_system.run_crawler(log=self._log, on_progress=None, on_done=None)

            self._log("\n[定时 3/3] 全部模块预计算")
            _precompute_all(self, self._log)

            self._log("\n" + "=" * 60)
            self._log(f"定时任务完成！总耗时 {time.time() - t_all:.1f}s")
            self._log("=" * 60)
        except Exception as e:
            import traceback
            self._log(f"定时任务出错: {e}")
            self._log(traceback.format_exc())

        if self._auto_shutdown.get():
            self._log("自动关机选项已启用，60 秒后关机...")
            self._log("如需取消，请在 60 秒内执行：shutdown /a")
            try:
                if sys.platform == "win32":
                    os.system('shutdown /s /t 60 /c "基金数据工具-定时任务已完成，系统即将关机"')
                else:
                    os.system("shutdown -h +1")
            except Exception as e:
                self._log(f"关机命令执行失败: {e}")

    base_cls._run_performance = _run_performance
    base_cls._run_risk = _run_risk
    base_cls._run_scoring_job = _run_scoring_job
    base_cls._run_efficiency = _run_efficiency
    base_cls._run_position = _run_position
    base_cls._run_timing = _run_timing
    base_cls._run_manager = _run_manager
    base_cls._run_cost = _run_cost
    base_cls._run_attribution = _run_attribution
    base_cls._run_composite = _run_composite
    base_cls._run_drawdown_shock = _run_drawdown_shock
    base_cls._run_trend_breakout = _run_trend_breakout
    base_cls._run_low_vol_stable = _run_low_vol_stable
    base_cls._run_oversold_rebound = _run_oversold_rebound
    base_cls._run_topic = _run_topic
    base_cls._run_one_click = _run_one_click
    base_cls._scheduled_run = _scheduled_run
    base_cls._precompute_all_modules = _precompute_all
    base_cls._module_cache_get = staticmethod(_cache_get)


_install_precompute_cache()

def _install_fast_crawler():
    """Speed up fund crawling by reducing per-fund history requests and raising safe concurrency."""
    import json as _json
    import os as _os
    import requests as _requests
    import threading as _threading
    import time as _time
    from concurrent.futures import ThreadPoolExecutor as _ThreadPoolExecutor, as_completed as _as_completed

    max_workers = int(_os.environ.get("FUND_FAST_WORKERS", "16"))
    nav_max_pages = int(_os.environ.get("FUND_NAV_MAX_PAGES", "1"))
    nav_page_size = int(_os.environ.get("FUND_NAV_PAGE_SIZE", "100"))
    bootstrap_nav_pages = int(_os.environ.get("FUND_BOOTSTRAP_NAV_PAGES", "0"))
    bootstrap_nav_hard_cap = int(_os.environ.get("FUND_BOOTSTRAP_NAV_HARD_CAP", "300"))
    history_min_points = int(_os.environ.get("FUND_HISTORY_MIN_POINTS", "500"))
    profile_timeout = int(_os.environ.get("FUND_PROFILE_TIMEOUT", "8"))
    nav_timeout = int(_os.environ.get("FUND_NAV_TIMEOUT", "8"))
    thread_local = _threading.local()

    def _fast_session(headers):
        sess = getattr(thread_local, "session", None)
        if sess is None:
            sess = _requests.Session()
            sess.trust_env = False
            sess.headers.update(headers)
            adapter = _requests.adapters.HTTPAdapter(
                pool_connections=max_workers * 2,
                pool_maxsize=max_workers * 2,
                max_retries=1,
            )
            sess.mount("https://", adapter)
            sess.mount("http://", adapter)
            thread_local.session = sess
        return sess

    def _fast_fetch_nav_history(code, headers, session=None, max_pages=None, page_size=None):
        """Fetch NAV history. max_pages=0 means fetch until empty, with a hard safety cap."""
        max_pages = nav_max_pages if max_pages is None else max_pages
        page_size = nav_page_size if page_size is None else page_size
        page_limit = bootstrap_nav_hard_cap if max_pages <= 0 else max_pages
        sess = session or _fast_session(headers)
        rows_out = []
        h = dict(headers)
        h["Referer"] = f"https://fund.eastmoney.com/{code}.html"
        for page in range(1, page_limit + 1):
            url = f"https://api.fund.eastmoney.com/f10/lsjz?fundCode={code}&pageIndex={page}&pageSize={page_size}"
            try:
                resp = sess.get(url, headers=h, timeout=nav_timeout)
                resp.encoding = "utf-8"
                data = _json.loads(jijin_system.parse_jsonp(resp.text))
                rows = (data.get("Data") or {}).get("LSJZList") or []
                if not rows:
                    break
                for row in rows:
                    value = row.get("LJJZ") or row.get("DWJZ")
                    if value in (None, "", "--"):
                        continue
                    try:
                        rows_out.append({"date": row.get("FSRQ"), "val": float(value)})
                    except Exception:
                        continue
                if len(rows) < page_size:
                    break
            except Exception:
                break
        rows_out.sort(key=lambda item: item["date"] or "")
        return rows_out

    def _merge_nav_history(old_rows, new_rows):
        merged = {}
        for item in (old_rows or []):
            date = str(item.get("date") or "").strip()
            if not date:
                continue
            try:
                merged[date] = {"date": date, "val": float(item.get("val"))}
            except Exception:
                continue
        for item in (new_rows or []):
            date = str(item.get("date") or "").strip()
            if not date:
                continue
            try:
                merged[date] = {"date": date, "val": float(item.get("val"))}
            except Exception:
                continue
        return [merged[k] for k in sorted(merged)]

    def _load_latest_valid_existing(output_dir, log, expected_total=0):
        import glob as _glob
        pattern = _os.path.join(output_dir, "fund_profile_*.json")
        files = sorted(_glob.glob(pattern), key=_os.path.getmtime, reverse=True)
        min_count = int(expected_total * 0.8) if expected_total else 0
        for path in files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = _json.load(f)
                if isinstance(data, list) and data:
                    if min_count and len(data) < min_count:
                        log(f"跳过未完成历史数据：{path}（{len(data):,}/{expected_total:,} 只）")
                        continue
                    log(f"迭代基底：使用上一份有效数据 {path}（{len(data):,} 只）")
                    return {str(item.get("fund_code", "")).zfill(6): item for item in data if item.get("fund_code")}, path
            except Exception as exc:
                log(f"跳过损坏/不可读历史数据：{path}（{type(exc).__name__}）")
        return {}, ""

    def _fast_fetch_one(args):
        idx, total, code, headers, old_fund, backfill_history = args
        code = str(code).zfill(6)
        try:
            session = _fast_session(headers)
            url = f"https://fund.eastmoney.com/{code}.html"
            resp = session.get(url, timeout=profile_timeout)
            resp.encoding = "utf-8"
            if resp.status_code != 200:
                if old_fund:
                    return True, old_fund, f"{code} 复用旧数据"
                return False, None, f"{code} HTTP {resp.status_code}"
            fund_data = jijin_system.FundDecompiler(resp.text, code).parse_all(headers)
            old_history = (old_fund or {}).get("nav_history") or []
            try:
                if old_history and backfill_history and len(old_history) < history_min_points:
                    full_history = _fast_fetch_nav_history(code, headers, session=session, max_pages=bootstrap_nav_pages, page_size=nav_page_size)
                    fund_data["nav_history"] = _merge_nav_history(old_history, full_history)
                elif old_history:
                    latest_history = _fast_fetch_nav_history(code, headers, session=session, max_pages=nav_max_pages, page_size=nav_page_size)
                    fund_data["nav_history"] = _merge_nav_history(old_history, latest_history)
                else:
                    fund_data["nav_history"] = _fast_fetch_nav_history(code, headers, session=session, max_pages=bootstrap_nav_pages, page_size=nav_page_size)
            except Exception:
                fund_data["nav_history"] = old_history
            return True, fund_data, code
        except Exception as exc:
            if old_fund:
                return True, old_fund, f"{code} 复用旧数据"
            return False, None, f"{code} {type(exc).__name__}: {str(exc)[:80]}"

    def _fast_run_crawler(log, on_progress=None, on_done=None):
        old_cwd = _os.getcwd()
        try:
            _os.chdir(SCRIPT_DIR)
            fund_codes_file = "target_funds.json"
            output_dir = "fund_data"
            log("开始数据爬取（迭代模式）...")
            log(f"迭代参数：并发 {max_workers}，已有基金只补最新 {nav_max_pages} 页 x {nav_page_size} 条；新基金完整历史页数={bootstrap_nav_pages or '直到结束'}")

            if not _os.path.exists(fund_codes_file):
                log("未找到 target_funds.json，请先运行【智能清洗】")
                if on_done:
                    on_done()
                return

            jijin_system.controller.reset()

            with open(fund_codes_file, "r", encoding="utf-8") as f:
                tasks = _json.load(f)
            if isinstance(tasks, dict) and "funds" in tasks:
                tasks = tasks["funds"]
            tasks = [str(item).zfill(6) for item in tasks]
            total = len(tasks)
            existing_map, existing_path = _load_latest_valid_existing(output_dir, log, expected_total=total)
            backfill_history = (_os.environ.get("FUND_BACKFILL_HISTORY", "0") == "1")
            log(f"共 {total:,} 只基金，开始并发爬取...")
            log("提示：迭代模式会保留旧文件中的完整 nav_history，每次只追加/更新最新净值点。")
            if backfill_history:
                log(f"历史补全：已开启。历史点少于 {history_min_points} 的基金会补抓完整历史；首次会明显变慢。")
            else:
                log("历史补全：未开启。本次只补最新净值，适合日常更新。")

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept-Language": "zh-CN,zh;q=0.9",
                "Connection": "keep-alive",
                "Referer": "https://fund.eastmoney.com/",
            }

            t0 = _time.time()
            ok = 0
            done = 0
            stopped = False

            with _ThreadPoolExecutor(max_workers=max_workers) as pool:
                future_map = {
                    pool.submit(_fast_fetch_one, (idx, total, code, headers, existing_map.get(code), backfill_history)): (idx, code)
                    for idx, code in enumerate(tasks, 1)
                }
                for future in _as_completed(future_map):
                    if not jijin_system.controller.wait_if_paused():
                        stopped = True
                        for f in future_map:
                            if not f.done():
                                f.cancel()
                        break

                    idx, code = future_map[future]
                    done += 1
                    try:
                        success, data, ret_code = future.result()
                    except Exception as exc:
                        success, data, ret_code = False, None, code
                        log(f"[{done}/{total}] {code} 异常: {exc}")

                    if success and data:
                        jijin_system.controller.add_result(data)
                        ok += 1
                    else:
                        log(f"[{done}/{total}] {ret_code} 失败")

                    now = _time.time()
                    if success and data:
                        rate = done / max(now - t0, 1)
                        remain = (total - done) / max(rate, 0.01)
                        log(f"[{done}/{total}] {ret_code} {data.get('fund_name', '')} | 成功 {ok} | {rate:.1f}只/秒 | 剩余 {remain/60:.1f}分钟")

                    if on_progress:
                        on_progress(ok, total)

            elapsed = _time.time() - t0
            if stopped:
                log(f"\n已停止爬取，已完成 {done}/{total}，成功 {ok} 条，用时 {elapsed:.1f}s")
            else:
                log(f"\n爬取完成！成功 {ok}/{total} 条，用时 {elapsed:.1f}s，平均 {elapsed / max(total, 1):.2f}s/只")

            results = jijin_system.controller.get_results()
            if results:
                code_order = {code: i for i, code in enumerate(tasks)}
                results.sort(key=lambda d: code_order.get(str(d.get("fund_code", "")).zfill(6), 999999))
                _os.makedirs(output_dir, exist_ok=True)
                ts = _dt.now().strftime("%Y%m%d_%H%M%S")
                outfile = _os.path.join(output_dir, f"fund_profile_{ts}.json")
                tmp_outfile = outfile + ".tmp"
                with open(tmp_outfile, "w", encoding="utf-8") as f:
                    _json.dump(results, f, ensure_ascii=False, separators=(",", ":"))
                _os.replace(tmp_outfile, outfile)
                log(f"已保存 JSON 数据：{outfile}")
                try:
                    jijin_system.save_results_to_excel(results, _os.path.join(output_dir, f"fund_profile_{ts}.xlsx"), log=log)
                except Exception as exc:
                    log(f"Excel 导出跳过/失败：{exc}")

            if on_done:
                on_done()
        finally:
            _os.chdir(old_cwd)

    jijin_system.fetch_nav_history = _fast_fetch_nav_history
    jijin_system.fetch_one = _fast_fetch_one
    jijin_system.run_crawler = _fast_run_crawler
    try:
        app_module.run_crawler = _fast_run_crawler
    except Exception:
        pass


_install_fast_crawler()

def _install_compact_main_controls():
    """Make the homepage button/notebook area compact so lower panels stay visible."""
    try:
        base_cls = jijin_system.FundToolsApp
    except Exception:
        return

    def _compact_add_tab_grouped(self, notebook, tab_title, groups):
        import tkinter as tk
        page = tk.Frame(notebook, bg="#1e1e2e", pady=2)
        notebook.add(page, text=tab_title)
        for group_label, btns in groups:
            row = tk.Frame(page, bg="#1e1e2e", pady=1)
            row.pack(fill="x", padx=4, pady=1)
            tk.Label(row, text=group_label,
                     font=("微软雅黑", 8, "bold"),
                     fg="#6c7086", bg="#1e1e2e", width=10, anchor="w").pack(side="left", padx=(2, 5))
            for text, color, cmd in btns:
                tk.Button(row, text=text, font=("微软雅黑", 9, "bold"),
                          width=9, height=1,
                          fg="#1e1e2e", bg=color, relief="flat", cursor="hand2",
                          command=cmd).pack(side="left", padx=2, pady=0)

    def _compact_add_taxonomy_tab(self, notebook, tab_title, sub_tree):
        import tkinter as tk
        page = tk.Frame(notebook, bg="#1e1e2e", pady=2)
        notebook.add(page, text=tab_title)

        color_idx = 0
        palette = getattr(self, "_TAB_PALETTE", [
            "#89b4fa", "#7fbf9f", "#fab387", "#f9e2af", "#cba6f7", "#94e2d5",
            "#74c7ec", "#89dceb", "#b4befe", "#eba0ac", "#f5c2e7", "#f38ba8",
        ])
        for level2, level3_dict in sub_tree.items():
            group_box = tk.LabelFrame(
                page, text=level2,
                font=("微软雅黑", 8, "bold"),
                fg="#f9e2af", bg="#1e1e2e", bd=1,
                labelanchor="nw", padx=4, pady=2)
            group_box.pack(fill="x", padx=4, pady=2)

            for level3, topics in level3_dict.items():
                line = tk.Frame(group_box, bg="#1e1e2e", pady=1)
                line.pack(fill="x")
                tk.Label(line, text=level3,
                         font=("微软雅黑", 8),
                         fg="#a6adc8", bg="#1e1e2e", width=11, anchor="w").pack(side="left", padx=(1, 5))
                for topic in topics:
                    color = palette[color_idx % len(palette)]
                    color_idx += 1
                    tk.Button(line, text=topic, font=("微软雅黑", 9, "bold"),
                              width=9, height=1,
                              fg="#1e1e2e", bg=color, relief="flat", cursor="hand2",
                              command=lambda t=topic: self._run_topic(t)).pack(side="left", padx=1, pady=0)

    original_build_ui = base_cls._build_ui

    def _compact_build_ui(self):
        original_build_ui(self)
        try:
            import tkinter as tk
            from tkinter import ttk
            style = ttk.Style(self.root)
            style.configure("TNotebook.Tab", padding=(8, 2))
            style.configure("TNotebook", tabmargins=(2, 1, 2, 0))

            def walk(widget):
                yield widget
                for child in widget.winfo_children():
                    yield from walk(child)

            for widget in walk(self.root):
                if isinstance(widget, ttk.Notebook):
                    try:
                        widget.configure(height=118)
                    except Exception:
                        pass
                    try:
                        widget.pack_configure(fill="x", expand=False, padx=16, pady=(0, 4))
                    except Exception:
                        pass
                    for tab_id in widget.tabs():
                        try:
                            page = widget.nametowidget(tab_id)
                            page.configure(height=88)
                            page.pack_propagate(False)
                        except Exception:
                            pass
            self.root.update_idletasks()
        except Exception:
            pass

    base_cls._add_tab_grouped = _compact_add_tab_grouped
    base_cls._add_taxonomy_tab = _compact_add_taxonomy_tab
    base_cls._build_ui = _compact_build_ui


_install_compact_main_controls()

def run_web():
    port = 5000
    url = f'http://127.0.0.1:{port}'
    def open_browser():
        webbrowser.open(url)
    threading.Timer(2.0, open_browser).start()
    try:
        # Print effective DATA_DIR used by the embedded app for verification
        data_dir = getattr(app_module, 'DATA_DIR', None)
        print(f"[INFO] embedded app DATA_DIR = {data_dir}")
        get_latest = getattr(app_module, 'get_latest_data_file', None)
        try:
            latest = get_latest() if callable(get_latest) else None
        except Exception as _:
            latest = None
        print(f"[INFO] embedded app latest data file = {latest}")
    except Exception:
        pass
    app.run(host='0.0.0.0', port=port, debug=False)

def run_desktop():
    def _start():
        root = desktop_app_module.tk.Tk()
        desktop_app_module.FundDesktopApp(root)
        root.mainloop()
    _run_in_script_dir(_start)

def run_fund_tools():
    def _start():
        app_instance = jijin_system.FundToolsApp()
        app_instance.root.mainloop()
    _run_in_script_dir(_start)

def run_launcher():
    import tkinter as tk
    win = tk.Tk()
    win.title('统一基金系统启动器')
    win.geometry('680x500')
    win.configure(bg='#1a1a2e')
    frame = tk.Frame(win, bg='#1a1a2e')
    frame.pack(expand=True, fill='both')
    tk.Label(frame, text='基金系统统一入口', fg='#e0e0e0', bg='#1a1a2e', font=('Microsoft YaHei', 18, 'bold')).pack(pady=20)
    status = tk.Label(frame, text='状态：就绪', fg='#dce1e8', bg='#1a1a2e', font=('Microsoft YaHei', 10))
    status.pack(pady=8)

    def trigger_update_from_launcher():
        status.config(text='状态：正在更新...')
        def task():
            try:
                def _update():
                    jijin_system.run_clean_list(lambda msg: None)
                    jijin_system.controller.reset()
                    jijin_system.run_crawler(log=lambda msg: None, on_progress=None, on_done=None)
                _run_in_script_dir(_update)
                status.config(text='状态：更新完成')
            except Exception as e:
                status.config(text=f'状态：更新失败 {e}')
        threading.Thread(target=task, daemon=True).start()

    tk.Button(frame, text='启动 原基金筛选系统', width=24, height=2, command=lambda: [win.destroy(), run_fund_tools()], bg='#8e44ad', fg='#ffffff').pack(pady=8)
    tk.Button(frame, text='启动 Web 可视化', width=24, height=2, command=lambda: [win.destroy(), run_web()], bg='#3498db', fg='#ffffff').pack(pady=8)
    tk.Button(frame, text='启动 桌面可视化', width=24, height=2, command=lambda: [win.destroy(), run_desktop()], bg='#6fb894', fg='#ffffff').pack(pady=8)
    tk.Button(frame, text='一键更新数据', width=24, height=2, command=trigger_update_from_launcher, bg='#e8b830', fg='#1a1a2e').pack(pady=8)
    win.mainloop()


def _run_module_cli(module_key):
    old_cwd = os.getcwd()
    try:
        os.chdir(SCRIPT_DIR)
        module_map = {
            "performance": ("收益表现", jijin_system.run_performance_score),
            "risk": ("风险回撤", jijin_system.run_risk_drawdown),
            "efficiency": ("风险效率", jijin_system.run_efficiency_score),
            "position": ("位置估值", jijin_system.run_position_score),
            "timing": ("趋势择时", jijin_system.run_timing_score),
            "manager": ("基金经理", jijin_system.run_manager_score),
            "cost": ("交易成本", jijin_system.run_cost_score),
            "attribution": ("收益归因", jijin_system.run_attribution_score),
            "composite": ("长期综合", jijin_system.run_composite_score),
            "drawdown_shock": ("回撤震荡", jijin_system.run_drawdown_shock_screen),
            "trend_breakout": ("趋势突破", jijin_system.run_trend_breakout_screen),
            "low_vol_stable": ("低波稳健", jijin_system.run_low_vol_stable_screen),
            "oversold_rebound": ("超跌反弹", jijin_system.run_oversold_rebound_screen),
        }
        if module_key.startswith("topic::"):
            topic_name = module_key.split("::", 1)[1]
            title = f"{topic_name}专题"
            func = lambda log: jijin_system.run_topic_screen(topic_name, log)
        else:
            title, func = module_map.get(module_key, ("", None))
        if not func:
            print(f"未知模块：{module_key}", flush=True)
            print("__RESULT_PATH__=", flush=True)
            return 2

        def _cli_log(message):
            print(str(message), flush=True)

        print(f"开始独立进程计算：{title}", flush=True)
        path = func(_cli_log)
        print(f"独立进程计算完成：{title}", flush=True)
        print(f"__RESULT_PATH__={os.path.abspath(path) if path else ''}", flush=True)
        return 0 if path else 1
    except Exception as exc:
        import traceback
        print(f"模块计算异常：{exc}", flush=True)
        print(traceback.format_exc(), flush=True)
        print("__RESULT_PATH__=", flush=True)
        return 1
    finally:
        os.chdir(old_cwd)


if __name__ == '__main__':
    if "--build-fund-index" in sys.argv:
        arg_index = sys.argv.index("--build-fund-index")
        data_arg = sys.argv[arg_index + 1] if len(sys.argv) > arg_index + 1 else None
        ok = _build_fund_index(data_arg, log=lambda msg: print(msg, flush=True))
        sys.exit(0 if ok else 1)
    if "--run-module" in sys.argv:
        arg_index = sys.argv.index("--run-module")
        module_arg = sys.argv[arg_index + 1] if len(sys.argv) > arg_index + 1 else ""
        sys.exit(_run_module_cli(module_arg))
    run_fund_tools()
