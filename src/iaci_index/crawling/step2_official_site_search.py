# step2_selenium_full.py
# 功能：为民办本科院校自动发现官网 URL（支持断点续跑）

import time
import random
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ================== 基本配置 ==================

# Step1 输出的民办本科 Excel（输入）
INPUT_FILE = "step1_private_undergrad.xlsx"

# 本脚本的结果文件（输出 & 断点续跑用）
OUTPUT_FILE = "step2_private_undergrad_with_urls_selenium.xlsx"

# 你的 chromedriver 路径 —— 必须改成你自己的
CHROMEDRIVER_PATH = r"E:\gwydata\pythonProject\drivers\chromedriver.exe"

# requests 解析跳转用的请求头
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0 Safari/537.36"
    )
}

# ================== 工具函数 ==================


def resolve_real_url(u: str) -> str:
    """
    把百度 link?url=... 这样的跳转链接解析成真实官网；
    如果不是 baidu 域名，直接返回；
    如果解析失败，就返回原始链接兜底。
    """
    if not u:
        return ""

    # 已经不是 baidu 域名，基本可以视为真实官网
    if "baidu.com" not in u:
        return u

    try:
        r = requests.get(
            u, headers=HEADERS, timeout=5, allow_redirects=True
        )
        return r.url
    except Exception as e:
        print("  ⚠️ 解析跳转失败，先保留百度链接：", e)
        return u


def init_driver():
    """初始化 Selenium 浏览器"""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option(
        "excludeSwitches", ["enable-automation"]
    )
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # 可以伪装一下 UA（可选）
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0 Safari/537.36"
    )

    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # 让 webdriver 标志变为 undefined，降低被识别为自动化的概率
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            """
        },
    )

    # 适当设置页面加载超时，防止某些页面过长时间无响应
    driver.set_page_load_timeout(15)

    return driver


def search_official_site(driver, school_name: str) -> str:
    """
    用 Selenium 打开百度搜索结果页，
    在页面中找一个最像官网的链接 href（不在 Selenium 中跳转），
    返回这个 href（可能是百度跳转，也可能已经是真实官网）。
    """
    query = f"{school_name} 官网"
    search_url = "https://www.baidu.com/s?wd=" + quote(query)

    print("  搜索 URL:", search_url)
    driver.get(search_url)
    # 如出现验证码，可在这里手动处理后回车（可取消注释）：
    # input("⚠️ 如出现百度验证，请在浏览器中处理后回车继续：")

    time.sleep(3.0)  # 等页面稳定

    # 抓所有标题里的链接（桌面版百度通常在 h3/h2 下）
    links = driver.find_elements(By.CSS_SELECTOR, "h3 a, h2 a")
    if not links:
        print("  ⛔ 没找到任何标题链接")
        return ""

    # 过滤掉明显不是官网的结果
    bad_keywords = ["百度百科", "百度知道", "贴吧", "知乎", "微博", "豆瓣"]

    candidate_href = None

    for a in links:
        try:
            text = a.text.strip()
            href = a.get_attribute("href") or ""
        except Exception:
            continue

        if not href:
            continue
        if any(bad in text for bad in bad_keywords):
            continue

        # 标题里包含学校名 / “官网”，优先认为是官网
        if (school_name[:2] in text) or ("官网" in text) or (school_name in text):
            candidate_href = href
            print("  选择链接标题：", text)
            break

    # 如果没有匹配到，就退而求其次，使用第一个结果
    if not candidate_href:
        first = links[0]
        candidate_href = first.get_attribute("href") or ""
        print("  回退：使用第一个结果链接")

    print("  初步候选 href:", candidate_href)
    return candidate_href


# ================== 主流程 ==================


def main():
    # 1. 读取数据：如果已有结果文件，从结果文件接着跑；否则从 Step1 文件开始
    if Path(OUTPUT_FILE).exists():
        print(f"🔁 检测到已有结果文件：{OUTPUT_FILE}，将从中断处继续。")
        df = pd.read_excel(OUTPUT_FILE)
    else:
        print(f"🆕 未发现结果文件，从 {INPUT_FILE} 开始新一轮采集。")
        df = pd.read_excel(INPUT_FILE)
        if "official_site" not in df.columns:
            df["official_site"] = ""

    # 2. 确保有 school_name 列
    if "school_name" not in df.columns:
        raise ValueError(
            f"列 'school_name' 不在当前数据中，请检查 {INPUT_FILE} / {OUTPUT_FILE}。"
        )

    # 3. 初始化浏览器
    driver = init_driver()

    # 4. 遍历学校，支持断点续跑
    total = len(df)
    for idx, row in df.iterrows():
        school = str(row["school_name"])

        # 已经有官网的跳过
        if isinstance(row.get("official_site", ""), str) and row.get(
            "official_site", ""
        ).strip():
            print(
                f"[跳过] {idx + 1}/{total} {school} 已有官网：{row['official_site']}"
            )
            continue

        print(f"\n=== {idx + 1}/{total}: 正在处理 {school} ===")

        # 4.1 先从百度结果页拿到一个候选 href
        try:
            raw_url = search_official_site(driver, school)
        except Exception as e:
            print(f"❌ 搜索 {school} 失败: {e}")
            raw_url = ""

        # 4.2 再用 requests 在后台解析真实官网（跟踪 302）
        url = resolve_real_url(raw_url)

        # 4.3 写入 DataFrame，并立即保存到 OUTPUT_FILE
        df.at[idx, "official_site"] = url
        print(f"➡️  记录官网：{url}")

        df.to_excel(OUTPUT_FILE, index=False)

        # 4.4 随机等待，模拟真人操作，降低被风控几率
        time.sleep(random.uniform(5.0, 10.0))

    driver.quit()
    print(f"\n✅ Selenium 版本采集完成！结果已保存到：{OUTPUT_FILE}")


def run_step2() -> None:
    """Run Selenium crawling to collect official site and search results."""
    main()


if __name__ == "__main__":
    run_step2()
