import os
import time
import random
import re
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

import urllib3
from urllib3.exceptions import InsecureRequestWarning

# 关闭 verify=False 带来的烦人警告
urllib3.disable_warnings(InsecureRequestWarning)


INPUT_FILE = "step2_private_undergrad_with_urls_selenium.xlsx"
OUTPUT_FILE = "step2_private_undergrad_with_urls_selenium.xlsx"  # 在同一文件上追加列

HEADERS_LIST = [
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0 Safari/537.36"
        ),
        "Accept-Language": "zh-CN,zh;q=0.9",
    },
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) "
            "Gecko/20100101 Firefox/120.0"
        ),
        "Accept-Language": "zh-CN,zh;q=0.9",
    },
]


def rand_headers():
    return random.choice(HEADERS_LIST)


def sleep_random(a=1.0, b=3.0):
    time.sleep(random.uniform(a, b))


def get_html(url, timeout=15, max_retries=3):
    """
    通用请求函数：带重试 + 随机 UA + 随机等待。
    关闭证书验证（verify=False），避免部分官网证书异常导致失败。
    """
    for i in range(max_retries):
        try:
            resp = requests.get(
                url,
                headers=rand_headers(),
                timeout=timeout,
                allow_redirects=True,
                verify=False,
            )
            if resp.status_code == 200 and "text/html" in resp.headers.get(
                "Content-Type", ""
            ):
                resp.encoding = resp.apparent_encoding
                return resp.text
        except Exception as e:
            print(f"[WARN] 请求失败 {url} 尝试 {i + 1}/{max_retries}，错误：{e}")
        sleep_random(1.5, 4.0)
    return None


def is_baike_url(url: str) -> bool:
    return "baike.baidu.com" in url


# ===== 链接挖掘规则 =====

# 学校概况 / 简介类
INFO_TEXT_KEYWORDS = [
    "学校概况", "学院概况", "学校简介", "学院简介", "学校概述",
    "学校介绍", "学院介绍", "学校一览", "基本情况", "学校信息",
    "学校历史", "学校沿革", "发展历程"
]
INFO_HREF_KEYWORDS = [
    "xxgk", "gaikuang", "gaikuo", "jianjie", "about", "profile"
]

# 信息公开类
DISCLOSURE_TEXT_KEYWORDS = [
    "信息公开", "校务公开", "政务公开", "办学信息公开", "教育信息公开"
]
DISCLOSURE_HREF_KEYWORDS = [
    "xxgk", "xxgkml", "xxgk_list", "xxgk.jsp", "xxgk.htm", "xxgk.do"
]

# 就业 / 就业质量报告
EMPLOY_TEXT_KEYWORDS = [
    "就业信息", "就业工作", "招生就业", "就业质量报告",
    "毕业生就业", "就业网", "就业办", "就业创业"
]
EMPLOY_HREF_KEYWORDS = [
    "jiuye", "jyxx", "zsjyc", "zhaoshengjiuye", "jyzd", "jyzl", "jygl"
]


def find_candidate_links(html: str,
                         base_url: str,
                         text_keywords,
                         href_keywords,
                         max_links=5):
    """
    在页面中查找符合特定关键词的链接，返回去重后的完整 URL 列表。
    """
    soup = BeautifulSoup(html, "html.parser")
    urls = []
    seen = set()

    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        href = a["href"]

        if not href:
            continue

        hit = False
        # 文本命中
        if any(kw in text for kw in text_keywords):
            hit = True
        # href 命中
        if any(kw in href for kw in href_keywords):
            hit = True

        if not hit:
            continue

        full_url = urljoin(base_url, href)
        # 简单过滤一下明显无效链接
        if full_url.startswith("javascript:") or full_url.startswith("#"):
            continue

        if full_url not in seen:
            seen.add(full_url)
            urls.append(full_url)
            if len(urls) >= max_links:
                break

    return urls


# ===== 保存与断点逻辑 =====

LINK_COLS = [
    "info_url_1", "info_url_2", "info_url_3",
    "disclosure_url_1", "disclosure_url_2",
    "employment_url_1", "employment_url_2", "employment_url_3",
]


def safe_save(df: pd.DataFrame):
    """
    实时保存：优先写 OUTPUT_FILE，如被 Excel 占用，则写备份文件。
    """
    try:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"[SAVE] 已写入 {OUTPUT_FILE}")
    except PermissionError:
        tmp_name = OUTPUT_FILE.replace(".xlsx", "_links_backup.xlsx")
        df.to_excel(tmp_name, index=False)
        print(f"[WARN] {OUTPUT_FILE} 被占用，已写入备份 {tmp_name}")


def row_links_already_good_enough(row) -> bool:
    """
    简单断点逻辑：
    - 如果这一行已经有至少 2 个非空链接（任意类别），认为“足够”了，下次跳过。
    """
    cnt = 0
    for col in LINK_COLS:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            cnt += 1
    return cnt >= 2


def main():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"未找到输入文件：{INPUT_FILE}")

    print(f"[INFO] 读取 {INPUT_FILE} ...")
    df = pd.read_excel(INPUT_FILE)

    # 确保有 school_name 和 official_site
    if "school_name" not in df.columns:
        raise ValueError("缺少列 'school_name'，请检查 Step1/Step2 的输出。")
    if "official_site" not in df.columns:
        raise ValueError("缺少列 'official_site'，请先完成 Step2 官网采集。")

    # 确保新列存在
    for col in LINK_COLS:
        if col not in df.columns:
            df[col] = None
        else:
            # 强制转为字符串类型，避免后续赋值 warning
            df[col] = df[col].astype("object")

    total = len(df)

    # 可选：只调试前 N 所学校
    DEBUG_N = None  # 比如先设成 10 测试，确认效果后改回 None
    if DEBUG_N is not None:
        iter_df = df.head(DEBUG_N)
    else:
        iter_df = df

    for idx, row in iter_df.iterrows():
        school = str(row.get("school_name", ""))
        url = str(row.get("official_site", "")).strip()

        if not url:
            print(f"[SKIP] {idx+1}/{total} {school} 无官网 URL")
            continue

        if is_baike_url(url):
            print(f"[SKIP] {idx+1}/{total} {school} 使用百科链接，暂不扩展子页面：{url}")
            continue

        # 如果本行已经有不少链接了，就跳过（断点续跑）
        if row_links_already_good_enough(row):
            print(f"[SKIP] {idx+1}/{total} {school} 已有多个扩展链接，跳过")
            continue

        print(f"\n[INFO] 采集扩展链接 {idx+1}/{total}：{school}  -> {url}")

        html = get_html(url)
        if not html:
            print(f"[WARN] 无法获取官网页面：{url}")
            safe_save(df)
            continue

        # 学校概况 / 简介类
        info_links = find_candidate_links(
            html, url, INFO_TEXT_KEYWORDS, INFO_HREF_KEYWORDS, max_links=3
        )
        if info_links:
            print(f"  [INFO] 学校概况类链接：")
            for i, link in enumerate(info_links, start=1):
                print(f"         info_url_{i}: {link}")
                col = f"info_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 信息公开类
        disc_links = find_candidate_links(
            html, url, DISCLOSURE_TEXT_KEYWORDS, DISCLOSURE_HREF_KEYWORDS, max_links=2
        )
        if disc_links:
            print(f"  [INFO] 信息公开类链接：")
            for i, link in enumerate(disc_links, start=1):
                print(f"         disclosure_url_{i}: {link}")
                col = f"disclosure_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 就业类 / 就业质量报告类
        emp_links = find_candidate_links(
            html, url, EMPLOY_TEXT_KEYWORDS, EMPLOY_HREF_KEYWORDS, max_links=3
        )
        if emp_links:
            print(f"  [INFO] 就业类链接：")
            for i, link in enumerate(emp_links, start=1):
                print(f"         employment_url_{i}: {link}")
                col = f"employment_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 实时保存
        safe_save(df)
        sleep_random(1.0, 3.0)

    print("\n[DONE] 扩展链接采集完成。")


if __name__ == "__main__":
    main()
