import os
import time
import random
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

import urllib3
from urllib3.exceptions import InsecureRequestWarning

# 关闭 verify=False 带来的警告
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


# ===== 链接挖掘工具函数 =====

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
        # 过滤明显无效链接
        if full_url.startswith("javascript:") or full_url.startswith("#"):
            continue

        if full_url not in seen:
            seen.add(full_url)
            urls.append(full_url)
            if len(urls) >= max_links:
                break

    return urls


# ============ 新增的几类链接规则 ============

# 1) 国际合作 / 国际交流
INTL_COOP_TEXT = [
    "国际交流", "国际合作", "对外合作", "合作院校",
    "国际教育学院", "国际学院", "国际部", "外事办公室"
]
INTL_COOP_HREF = [
    "gjhz", "gjjl", "hezuo", "cooperation",
    "international", "global", "waishi", "intldept"
]

# 2) 出国留学 / 留学项目
STUDYABROAD_TEXT = [
    "留学项目", "出国留学", "出国项目", "境外学习",
    "海外学习", "留学通道", "2+2", "3+1"
]
STUDYABROAD_HREF = [
    "liuxue", "studyabroad", "goabroad", "overseas",
    "lxxm", "lx", "2+2", "3+1"
]

# 3) 海外实习 / 境外实践
OVERSEA_PRACTICE_TEXT = [
    "海外实习", "境外实习", "海外实践", "境外实践",
    "海外研修", "境外研修"
]
OVERSEA_PRACTICE_HREF = [
    "shixi", "intern", "practice", "overseas",
    "internship", "practicum"
]

# 4) 东盟 / 区域合作 / 一带一路
ASEAN_TEXT = [
    "东盟", "ASEAN", "一带一路", "区域合作",
    "区域研究", "东南亚研究", "中国—东盟"
]
ASEAN_HREF = [
    "asean", "dongmeng", "yidaiyilu", "region",
    "dongnanya", "sea_study"
]

# 5) 专业设置 / 语种 / 学科专业
MAJOR_PROGRAM_TEXT = [
    "专业设置", "本科专业", "学科专业", "专业一览",
    "人才培养方案", "外国语学院", "语言学院"
]
MAJOR_PROGRAM_HREF = [
    "zysz", "zhuanye", "major", "program",
    "subject", "discipline", "zyml"
]


# ===== 保存与断点逻辑 =====

NEW_LINK_COLS = [
    "intl_coop_url_1", "intl_coop_url_2", "intl_coop_url_3",
    "studyabroad_url_1", "studyabroad_url_2", "studyabroad_url_3",
    "overseas_practice_url_1", "overseas_practice_url_2",
    "asean_url_1", "asean_url_2",
    "major_program_url_1", "major_program_url_2",
]


def safe_save(df: pd.DataFrame):
    """
    实时保存：优先写 OUTPUT_FILE，如被 Excel 占用，则写备份文件。
    """
    try:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"[SAVE] 已写入 {OUTPUT_FILE}")
    except PermissionError:
        tmp_name = OUTPUT_FILE.replace(".xlsx", "_extra_links_backup.xlsx")
        df.to_excel(tmp_name, index=False)
        print(f"[WARN] {OUTPUT_FILE} 被占用，已写入备份 {tmp_name}")


def row_new_links_good_enough(row) -> bool:
    """
    简单断点逻辑：
    - 如果这一行在 NEW_LINK_COLS 中已有 >= 3 个非空链接，就认为“够用”，下次跳过。
    """
    cnt = 0
    for col in NEW_LINK_COLS:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            cnt += 1
    return cnt >= 3


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

    # 确保新增列存在 & 类型为 object
    for col in NEW_LINK_COLS:
        if col not in df.columns:
            df[col] = None
        df[col] = df[col].astype("object")

    total = len(df)

    # 可选：调试模式，只跑前 N 所学校
    DEBUG_N = None  # 比如先设成 10 调试，确认效果后改回 None
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
            print(f"[SKIP] {idx+1}/{total} {school} 使用百科链接，暂不扩展 extra 链接：{url}")
            continue

        # 断点续跑：如果新列已经有不少链接，就跳过
        if row_new_links_good_enough(row):
            print(f"[SKIP] {idx+1}/{total} {school} 已有多条 extra 链接，跳过")
            continue

        print(f"\n[INFO] 扩展 extra 链接 {idx+1}/{total}：{school}  -> {url}")

        html = get_html(url)
        if not html:
            print(f"[WARN] 无法获取官网页面：{url}")
            safe_save(df)
            continue

        # 1) 国际合作 / 国际交流
        intl_links = find_candidate_links(
            html, url, INTL_COOP_TEXT, INTL_COOP_HREF, max_links=3
        )
        if intl_links:
            print("  [INFO] 国际合作 / 国际交流链接：")
            for i, link in enumerate(intl_links, start=1):
                print(f"         intl_coop_url_{i}: {link}")
                col = f"intl_coop_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 2) 出国留学 / 留学项目
        sa_links = find_candidate_links(
            html, url, STUDYABROAD_TEXT, STUDYABROAD_HREF, max_links=3
        )
        if sa_links:
            print("  [INFO] 留学 / 出国项目链接：")
            for i, link in enumerate(sa_links, start=1):
                print(f"         studyabroad_url_{i}: {link}")
                col = f"studyabroad_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 3) 海外实习 / 境外实践
        op_links = find_candidate_links(
            html, url, OVERSEA_PRACTICE_TEXT, OVERSEA_PRACTICE_HREF, max_links=2
        )
        if op_links:
            print("  [INFO] 海外实习 / 实践链接：")
            for i, link in enumerate(op_links, start=1):
                print(f"         overseas_practice_url_{i}: {link}")
                col = f"overseas_practice_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 4) 东盟 / 区域合作
        asean_links = find_candidate_links(
            html, url, ASEAN_TEXT, ASEAN_HREF, max_links=2
        )
        if asean_links:
            print("  [INFO] 东盟 / 区域合作链接：")
            for i, link in enumerate(asean_links, start=1):
                print(f"         asean_url_{i}: {link}")
                col = f"asean_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 5) 专业设置 / 语种 / 学科专业
        major_links = find_candidate_links(
            html, url, MAJOR_PROGRAM_TEXT, MAJOR_PROGRAM_HREF, max_links=2
        )
        if major_links:
            print("  [INFO] 专业设置 / 语种相关链接：")
            for i, link in enumerate(major_links, start=1):
                print(f"         major_program_url_{i}: {link}")
                col = f"major_program_url_{i}"
                if col in df.columns:
                    df.at[idx, col] = link

        # 实时保存
        safe_save(df)
        sleep_random(1.0, 3.0)

    print("\n[DONE] extra 扩展链接采集完成。")


def run_step2_extra_info() -> None:
    """Collect extra links for each school record."""
    main()


if __name__ == "__main__":
    run_step2_extra_info()
