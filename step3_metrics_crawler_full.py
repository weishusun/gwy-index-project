# -*- coding: utf-8 -*-
"""
Step3：民办本科院校 基础办学指标 + NLP 文本字段 爬虫（Excel 版本，最终优化版）

功能：
1）从 step2_private_undergrad_with_urls_selenium.xlsx 读取各学校官网及扩展链接。
2）按学校循环，访问若干“有用的”页面，提取：
    - 数值类基础办学指标（学生数、教师数、面积等）
    - 学校简介文本（profile_text_snippet，用于 NLP）
    - 办学定位标签（positioning_keywords，规则抽取）
3）将结果写回 step3_private_undergrad_metrics_2025.xlsx（每行 = 1 所学校）。
4）支持：
    - URL 过滤（去掉教务、登录、IP、无关新闻）
    - HTML 本地缓存
    - 断点续跑（metrics_status = 'ok' 的学校默认跳过）
    - 定期保存 Excel

依赖：
    pip install pandas requests beautifulsoup4 lxml openpyxl
"""

import os
import re
import time
import random
from datetime import datetime
from typing import Dict, List, Optional

import requests
from requests.exceptions import SSLError
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin


# ========= 配置区域 =========

# Step2 输出（包含 official_site + 各类 *_url_*）
URLS_FILE = "step2_private_undergrad_with_urls_selenium.xlsx"
# Step3 初始化后的指标表
METRICS_FILE = "step3_private_undergrad_metrics_2025.xlsx"

# HTML 缓存目录
HTML_CACHE_DIR = "html_cache_step3"

# 每处理多少所学校保存一次 Excel
SAVE_EVERY_N_SCHOOLS = 5
# 单校最多访问多少个页面（放得比较宽，先多爬一些）
MAX_PAGES_PER_SCHOOL = 10
# 简介截断长度（字符）
MAX_TEXT_SNIPPET_LEN = 800

# 请求相关
MAX_RETRIES = 2
REQUEST_TIMEOUT = 15
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )
}


# ========= 通用工具函数 =========

def ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def safe_int(num_str: str) -> Optional[int]:
    """
    将带中文单位/逗号/空格的数字字符串转为 int。
    示例："3,456人"、"约1.2万名"、"80万平方米"。
    """
    if not isinstance(num_str, str):
        return None

    s = num_str.replace(",", "").replace("，", "").replace(" ", "")

    # 处理 “1.2万”
    m = re.match(r"([\d\.]+)\s*万", s)
    if m:
        try:
            return int(float(m.group(1)) * 10000)
        except ValueError:
            return None

    s = re.sub(r"[^\d]", "", s)
    if not s:
        return None

    try:
        return int(s)
    except ValueError:
        return None


def cache_key_for_url(url: str) -> str:
    """根据 URL 生成简单文件名"""
    import hashlib
    h = hashlib.md5(url.encode("utf-8")).hexdigest()
    return h + ".html"


def normalize_special_url(url: str) -> Optional[str]:
    """
    针对一些“奇怪但可以修”的链接做规范化：
    - 丢掉 javascript: / mailto: / tel:
    - 把 http://id.sanyau.edu.cn:9092/... 替换成 https://id.sanyau.edu.cn/
    其它情况基本原样返回。
    """
    if not isinstance(url, str):
        return None
    u = url.strip()
    if not u:
        return None

    lower = u.lower()

    # 1) 伪协议直接丢弃
    if lower.startswith("javascript:") or lower.startswith("mailto:") or lower.startswith("tel:"):
        return None

    # 2) 三亚学院的统一身份认证：换成可访问的 https 根域名
    if "id.sanyau.edu.cn:9092" in u:
        return "https://id.sanyau.edu.cn/"

    return u


def fetch_html(
    url: str,
    session: requests.Session,
    use_cache: bool = True,
    max_retries: int = MAX_RETRIES,
    timeout: int = REQUEST_TIMEOUT,
) -> Optional[str]:
    """
    获取网页 HTML，带重试、随机等待和本地缓存。
    专门对 SSLError 再用 verify=False 重试一次，兼容证书链不全的网站。
    """
    if not url or not isinstance(url, str):
        return None

    url = url.strip()
    if not url:
        return None

    ensure_dir(HTML_CACHE_DIR)
    cache_path = os.path.join(HTML_CACHE_DIR, cache_key_for_url(url))

    # 先尝试从缓存读取
    if use_cache and os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception:
            pass

    for attempt in range(1, max_retries + 1):
        try:
            try:
                # 第一次正常校验证书
                resp = session.get(url, headers=HEADERS, timeout=timeout)
            except SSLError as ssl_err:
                print(f"[WARN] {url} SSL 证书校验失败，尝试使用 verify=False 重新请求：{ssl_err}")
                resp = session.get(url, headers=HEADERS, timeout=timeout, verify=False)

            ct = resp.headers.get("Content-Type", "")
            if resp.status_code == 200 and "text/html" in ct:
                resp.encoding = resp.apparent_encoding or resp.encoding
                html = resp.text

                # 写缓存
                try:
                    with open(cache_path, "w", encoding="utf-8", errors="ignore") as f:
                        f.write(html)
                except Exception:
                    pass

                # 简单随机等待，降低被封风险
                time.sleep(random.uniform(0.8, 2.0))
                return html
            else:
                print(f"[WARN] {url} 状态码 {resp.status_code} / Content-Type={ct}")
        except Exception as e:
            print(f"[WARN] 请求失败 {url}，重试 {attempt}/{max_retries}，原因：{e}")
            time.sleep(2 * attempt)

    print(f"[ERROR] 多次重试后仍无法访问：{url}")
    return None


def discover_about_links(html: str, base_url: str) -> List[str]:
    """
    从首页 HTML 中自动发现“学校概况 / 学校简介 / 信息公开 / 校情总览”等链接。
    """
    if not html:
        return []

    soup = BeautifulSoup(html, "lxml")
    found_urls = []
    keywords = [
        "学校概况", "学院概况", "学校简介", "学院简介",
        "学校介绍", "学院介绍", "校情总览",
        "信息公开", "学校概览"
    ]

    for a in soup.find_all("a", href=True):
        text = (a.get_text() or "").strip()
        if any(kw in text for kw in keywords):
            href = a["href"]
            full_url = href if href.startswith("http") else urljoin(base_url, href)
            found_urls.append(full_url)

    return list(dict.fromkeys(found_urls))


def normalize_whitespace(text: str) -> str:
    """压缩多余空白"""
    if not isinstance(text, str):
        return ""
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ========= URL 过滤 =========

def is_useful_for_profile_or_metrics(url: str) -> bool:
    """
    判断 URL 是否“有可能”对概况 / 简介 / 办学指标有用。
    - 排除教务系统、登录、VPN、邮箱等；
    - 排除 IP 地址形式的内部系统；
    - 排除普通新闻页（带年份路径但无 gk/jianjie/about 等关键字）。
    """
    if not isinstance(url, str):
        return False
    u = url.strip().lower()
    if not u:
        return False

    # 必须是 http/https 开头
    if not (u.startswith("http://") or u.startswith("https://")):
        return False

    # 1) 排除系统/登录类
    bad_keywords = [
        "jwgl", "jwglxt", "jwc", "ids", "caslogin", "login",
        "vpn", "webvpn", "sso", "oa.", "mail.", "webmail",
        "ecard", "xsgl", "student", "teacher", "portal"
    ]
    if any(bk in u for bk in bad_keywords):
        return False

    # 2) 排除 IP 地址形式
    if re.match(r"^https?://\d{1,3}(\.\d{1,3}){3}", u):
        return False

    # 3) 新闻路径（/2024/xx/xx/），一般不需要
    if re.search(r"/20\d{2}/\d{1,2}/", u):
        if not any(kw in u for kw in ["gk", "gaikuang", "jianjie", "about", "xxgk"]):
            return False

    return True


# ========= NLP 相关 =========

def extract_profile_text(html: str) -> str:
    """
    从页面 HTML 中提取“最像学校概况/简介”的一段文本。
    返回未截断的长文本。
    """
    if not html:
        return ""

    soup = BeautifulSoup(html, "lxml")

    # 1) 根据标题寻找块
    title_keywords = ["学校概况", "学院概况", "学校简介", "学院简介", "学校介绍", "学院介绍", "校情总览"]
    for h_tag in ["h1", "h2", "h3", "h4", "h5"]:
        for h in soup.find_all(h_tag):
            title = (h.get_text() or "").strip()
            if any(kw in title for kw in title_keywords):
                texts = []
                for sib in h.find_all_next(["p", "div"], limit=40):
                    t = (sib.get_text(" ", strip=True) or "")
                    if len(t) < 50:
                        continue
                    texts.append(t)
                full = " ".join(texts)
                if len(full) > 150:
                    return full

    # 2) 找最长的一段 <p>
    candidate = ""
    for p in soup.find_all("p"):
        t = (p.get_text(" ", strip=True) or "")
        if len(t) > len(candidate):
            candidate = t
    if len(candidate) >= 150:
        return candidate

    # 3) 退化为整页文本
    return soup.get_text(" ", strip=True)


def extract_positioning_keywords(text: str) -> str:
    """
    从简介文本中抽取办学定位/特征标签。
    返回：逗号分隔字符串。
    """
    if not isinstance(text, str):
        return ""

    labels = []
    t = text
    tl = text.lower()

    rules = {
        "应用型": ["应用型", "应用技术型", "应用型本科", "实践教学", "产教融合", "校企合作"],
        "国际化": ["国际化", "international", "海外交流", "境外交流", "海外高校", "联合培养", "中外合作办学"],
        "东盟": ["东盟", "asean", "一带一路", "rcep", "澜湄合作", "泛北部湾"],
        "外语特色": ["外国语", "外语学院", "翻译", "口译", "笔译", "小语种", "语言类"],
        "商科/经管": ["商学院", "经济管理", "经管", "工商管理", "会计学", "金融学", "市场营销"],
        "师范": ["师范", "教师教育", "教育学院", "教师培养"],
        "医学健康": ["医学院", "护理学", "医学", "康复", "健康管理"],
        "信息技术": ["信息工程", "计算机科学", "软件工程", "大数据", "人工智能", "网络技术"],
    }

    for label, kws in rules.items():
        for kw in kws:
            if kw in t or kw.lower() in tl:
                labels.append(label)
                break

    labels = list(dict.fromkeys(labels))
    return ", ".join(labels)


# ========= 数值指标解析 =========

def parse_basic_metrics_from_page(html: str) -> Dict[str, Optional[int]]:
    """
    从页面文本中提取基础办学数值指标。
    """
    metrics: Dict[str, Optional[int]] = {}
    if not html:
        return metrics

    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(" ", strip=True)

    # 1. 建校年份
    m = re.search(r"(始建于|建校于|创建于|创办于)\s*(\d{4})\s*年", text)
    if m:
        year = int(m.group(2))
        if 1900 < year < 2100:
            metrics["founded_year"] = year

    # 2. 在校生人数
    patterns_students = [
        r"(现有在校生|在校生|在校学生|全日制在校生)[^0-9万]{0,10}([\d,，\.万]+)\s*人",
        r"(全日制本专科生|全日制本科生)[^0-9万]{0,10}([\d,，\.万]+)\s*人",
    ]
    for pat in patterns_students:
        m = re.search(pat, text)
        if m:
            v = safe_int(m.group(2))
            if v and v > 0:
                metrics["students_total"] = v
                break

    # 3. 教师 / 专任教师
    patterns_teachers = [
        r"(专任教师)[^0-9万]{0,10}([\d,，\.万]+)\s*人",
        r"(教师|教职工)[^0-9万]{0,10}([\d,，\.万]+)\s*人",
    ]
    for pat in patterns_teachers:
        m = re.search(pat, text)
        if m:
            v = safe_int(m.group(2))
            if v and v > 0:
                key = "fulltime_teachers" if "专任" in m.group(1) else "teachers_total"
                if key not in metrics or metrics[key] is None:
                    metrics[key] = v

    # 4. 校区数量
    m = re.search(r"(设有|拥有|现有)[^校区]{0,5}(\d+|一|二|两|三|四|五|六|七|八|九|十)\s*个?校区", text)
    if m:
        num_str = m.group(2)
        num_map = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
                   "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
        metrics["campus_count"] = int(num_str) if num_str.isdigit() else num_map.get(num_str, None)

    # 5. 学院数量
    m = re.search(r"(设有|下设|建有)[^学院]{0,5}(\d+)\s*个?(二级)?学院", text)
    if m:
        metrics["college_count"] = int(m.group(2))

    # 6. 本科专业数量
    m = re.search(r"(开设|设有|拥有)[^专业]{0,5}(\d+)\s*个?(本科)?专业", text)
    if m:
        metrics["major_count"] = int(m.group(2))

    # 7. 占地面积（亩/平方米）
    m = re.search(r"占地面积[^0-9万]{0,10}([\d,，\.万]+)\s*亩", text)
    if m:
        v = safe_int(m.group(1))
        if v and v > 0:
            metrics["campus_area_mu"] = v

    m = re.search(r"占地面积[^0-9万]{0,10}([\d,，\.万]+)\s*万?\s*平方米", text)
    if m:
        v = safe_int(m.group(1))
        if v and v > 0:
            metrics["campus_area_m2"] = v

    # 8. 馆藏图书
    m = re.search(r"(馆藏|藏书|图书馆)[^0-9万]{0,10}([\d,，\.万]+)\s*册", text)
    if m:
        v = safe_int(m.group(2))
        if v and v > 0:
            metrics["library_books"] = v

    # 9. 实验室/实训基地数量
    m = re.search(r"(实验室|实训室|实训基地)[^0-9]{0,8}(\d+)\s*个", text)
    if m:
        metrics["labs_count"] = int(m.group(2))

    # 10. 师生比
    m = re.search(r"师生比\s*为?\s*([0-9\.]+)\s*[:：]\s*1", text)
    if m:
        try:
            metrics["student_teacher_ratio"] = float(m.group(1))
        except ValueError:
            pass

    return metrics


def merge_metrics(
    target: Dict[str, Optional[int]],
    new_metrics: Dict[str, Optional[int]],
) -> Dict[str, Optional[int]]:
    """仅在原值为空时写入新指标值"""
    for k, v in new_metrics.items():
        if v is None:
            continue
        if k not in target or target[k] is None:
            target[k] = v
    return target


# ========= 主流程 =========

def main():
    print("=== Step3 基础办学指标 + NLP 文本字段爬虫（最终优化版）启动 ===")

    if not os.path.exists(URLS_FILE):
        raise FileNotFoundError(f"找不到 {URLS_FILE}，请确认 Step2 已完成。")
    if not os.path.exists(METRICS_FILE):
        raise FileNotFoundError(f"找不到 {METRICS_FILE}，请确认 Step3 初始化脚本已运行。")

    df_urls = pd.read_excel(URLS_FILE)
    df_metrics = pd.read_excel(METRICS_FILE)

    if "school_name" not in df_urls.columns or "school_name" not in df_metrics.columns:
        raise ValueError("两个文件中都必须包含 'school_name' 列。")

    df_urls = df_urls.set_index("school_name")
    df_metrics = df_metrics.set_index("school_name")

    # 这些列未来都要写入字符串，统一设为 object，避免 float64 冲突
    text_cols = [
        "metrics_status",
        "last_crawled_at",
        "profile_page_url",
        "profile_text_snippet",
        "positioning_keywords",
    ]

    for col in text_cols:
        if col not in df_metrics.columns:
            # 新建时直接用 object dtype
            df_metrics[col] = pd.Series([None] * len(df_metrics), dtype="object")
        else:
            # 已经存在的列强制转为 object
            df_metrics[col] = df_metrics[col].astype("object")

    # 数值指标列
    metric_cols = [
        "founded_year",
        "students_total",
        "teachers_total",
        "fulltime_teachers",
        "campus_count",
        "college_count",
        "major_count",
        "campus_area_mu",
        "campus_area_m2",
        "library_books",
        "labs_count",
        "student_teacher_ratio",
    ]
    for col in metric_cols:
        if col not in df_metrics.columns:
            df_metrics[col] = None

    session = requests.Session()
    # 不使用系统中的 HTTP(S)_PROXY（避免 127.0.0.1:7890 超时）
    session.trust_env = False

    ensure_dir(HTML_CACHE_DIR)

    total_schools = len(df_urls)
    processed = 0
    changed_since_last_save = 0

    for school_name, row in df_urls.iterrows():
        processed += 1
        print(f"\n>>> [{processed}/{total_schools}] 处理学校：{school_name}")

        status = df_metrics.at[school_name, "metrics_status"]
        if isinstance(status, str) and status.strip().lower() == "ok":
            print("  已标记为 ok，跳过（如需重爬可手动清空 metrics_status）。")
            continue

        # -------- 组装 URL 列表（先放到 raw_urls，再统一规范化+过滤） --------
        raw_urls: List[str] = []
        official = row.get("official_site", None)
        official_norm: Optional[str] = None
        if isinstance(official, str) and official.strip():
            official_norm = normalize_special_url(official.strip())
            if official_norm:
                raw_urls.append(official_norm)

        info_like_cols = []
        other_cols = []
        for col in df_urls.columns:
            if not isinstance(col, str):
                continue
            if col.startswith("info_url_") or col.startswith("disclosure_url_"):
                info_like_cols.append(col)
            elif col.endswith("_url") or "_url_" in col:
                other_cols.append(col)

        for col in info_like_cols + other_cols:
            val = row.get(col, None)
            if isinstance(val, str) and val.strip():
                nv = normalize_special_url(val.strip())
                if nv:
                    raw_urls.append(nv)

        # 去重 + URL 过滤
        candidate_urls: List[str] = []
        for u in raw_urls:
            nu = normalize_special_url(u)
            if not nu:
                continue
            if is_useful_for_profile_or_metrics(nu):
                if nu not in candidate_urls:
                    candidate_urls.append(nu)

        if not candidate_urls:
            print("  没有任何可用 URL，标记为 missing。")
            df_metrics.at[school_name, "metrics_status"] = "missing"
            df_metrics.at[school_name, "last_crawled_at"] = datetime.now()
            changed_since_last_save += 1
            continue

        metrics_collected: Dict[str, Optional[int]] = {}
        profile_page_url: Optional[str] = df_metrics.at[school_name, "profile_page_url"] or None
        profile_text: Optional[str] = df_metrics.at[school_name, "profile_text_snippet"] or None
        pages_visited = 0

        for url in candidate_urls:
            if pages_visited >= MAX_PAGES_PER_SCHOOL:
                print("  已达到 MAX_PAGES_PER_SCHOOL 限制，停止继续访问。")
                break

            print(f"  访问页面：{url}")
            html = fetch_html(url, session=session)
            if not html:
                continue

            pages_visited += 1

            # 如果是官网首页，顺便发现概况链接并加入 raw_urls（然后过滤再进入候选）
            if official_norm and url == official_norm and pages_visited == 1:
                discovered = discover_about_links(html, base_url=url)
                if discovered:
                    print(f"  自动发现 {len(discovered)} 个概况相关链接。")
                    for durl in discovered:
                        nd = normalize_special_url(durl)
                        if nd and nd not in raw_urls:
                            raw_urls.append(nd)
                    for durl in discovered:
                        nd = normalize_special_url(durl)
                        if nd and is_useful_for_profile_or_metrics(nd) and nd not in candidate_urls:
                            candidate_urls.append(nd)

            # 数值指标
            metrics_page = parse_basic_metrics_from_page(html)
            metrics_collected = merge_metrics(metrics_collected, metrics_page)

            # 简介文本（只要还没拿到 profile_text，就尝试）
            if not profile_text:
                url_lower = url.lower()
                prefer = any(kw in url_lower for kw in ["gk", "gaikuang", "jianjie", "about", "xxgk"])
                pt = extract_profile_text(html)
                pt = normalize_whitespace(pt)
                if prefer and len(pt) > 150:
                    profile_text = pt
                    profile_page_url = url
                elif not prefer and len(pt) > 250:
                    profile_text = pt
                    profile_page_url = url

        # 写回数值指标
        if metrics_collected:
            for k, v in metrics_collected.items():
                if k in df_metrics.columns:
                    df_metrics.at[school_name, k] = v

        # 写回 NLP 字段
        if profile_text:
            profile_text = str(profile_text)
            snippet = profile_text[:MAX_TEXT_SNIPPET_LEN]
            df_metrics.at[school_name, "profile_page_url"] = profile_page_url
            df_metrics.at[school_name, "profile_text_snippet"] = snippet
            df_metrics.at[school_name, "positioning_keywords"] = extract_positioning_keywords(profile_text)

        # 状态判定
        has_any_metric = any(
            df_metrics.at[school_name, col] not in [None, ""]
            and not (isinstance(df_metrics.at[school_name, col], float) and pd.isna(df_metrics.at[school_name, col]))
            for col in metric_cols
        )
        has_profile = bool(profile_text)

        if has_any_metric or has_profile:
            df_metrics.at[school_name, "metrics_status"] = "ok" if has_any_metric else "partial"
        else:
            df_metrics.at[school_name, "metrics_status"] = "missing"

        df_metrics.at[school_name, "last_crawled_at"] = datetime.now().isoformat()
        changed_since_last_save += 1
    
        # 定期保存
        if changed_since_last_save >= SAVE_EVERY_N_SCHOOLS:
            print("  达到保存阈值，写回 Excel...")
            df_metrics.reset_index().to_excel(METRICS_FILE, index=False)
            changed_since_last_save = 0

    print("\n=== 全部学校处理完毕，写回最终 Excel ===")
    df_metrics.reset_index().to_excel(METRICS_FILE, index=False)
    print("完成。")


if __name__ == "__main__":
    main()
