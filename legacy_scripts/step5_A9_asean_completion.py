"""
Step5-A9：自动补全东盟（ASEAN）合作指标 + 计算 ARII（ASEAN 区域国际化指数）

输入：step5_A8_major_fixed_IACI.xlsx（上一阶段的结果）
输出：step5_A9_asean_features.xlsx（增加东盟相关字段 + ARII）
"""

import os
import json
import time
import random
from typing import Dict, Any, List

import pandas as pd
import numpy as np
from openai import OpenAI

# ========= 1. 配置区 =========

INPUT_FILE = "step5_A8_major_fixed_IACI.xlsx"
OUTPUT_FILE = "step5_A9_asean_features.xlsx"

# 建议用环境变量，或者你自己改成和 step4C 一样的方式
MOONSHOT_API_KEY = "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy"
MOONSHOT_API_BASE = "https://api.moonshot.cn/v1"
KIMI_MODEL_NAME = "kimi-k2-turbo-preview"

# 控制请求间隔，防止过快
SLEEP_BASE = 3
SLEEP_JITTER = 2

# 需要 Kimi 补全的东盟相关字段
TARGET_METRICS: Dict[str, Dict[str, str]] = {
    "asean_partner_countries_count": {
        "desc": "学校与多少个东盟国家（东盟10国）有正式合作或交流关系（请给出大致整数，若无则为0）。",
    },
    "asean_partner_universities_count": {
        "desc": "学校与东盟高校的合作院校数量（交换、联合培养、合作办学等，给出大致整数，若无则为0）。",
    },
    "asean_program_count": {
        "desc": "与东盟相关的项目数量（例如东盟交换项目、研学、暑期学校、联合培养等，粗略估算一个整数，若无则为0）。",
    },
}

client = OpenAI(
    api_key=MOONSHOT_API_KEY,
    base_url=MOONSHOT_API_BASE,
)

# ========= 2. 工具函数 =========


def minmax(s: pd.Series) -> pd.Series:
    s = pd.to_numeric(s, errors="coerce").fillna(0).astype(float)
    min_v, max_v = s.min(), s.max()
    if max_v == min_v:
        return pd.Series(0.0, index=s.index)
    return (s - min_v) / (max_v - min_v)


def parse_json(text: str) -> Dict[str, Any]:
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass

    s = text.find("{")
    e = text.rfind("}")
    if s != -1 and e != -1:
        try:
            return json.loads(text[s:e + 1])
        except Exception:
            pass

    return {}


def build_row_prompt(school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]) -> str:
    """
    构造让 Kimi 抽取东盟合作指标的 prompt
    """
    ctx_json = json.dumps(row_context, ensure_ascii=False)

    metric_desc = "\n".join(
        [f"- {m}: {TARGET_METRICS[m]['desc']}" for m in missing_metrics]
    )

    return f"""
你是一名严谨的教育数据研究助手，需要通过联网搜索评估一所中国民办本科高校在“东盟区域国际化”方面的情况。

【学校名称】
{school_name}

【已知基础信息（JSON）】
{ctx_json}

【你的任务】
1. 访问该校官网（包括：国际合作与交流处、国际学院、港澳台与国际教育、海外合作项目介绍等栏目）。
2. 重点关注与“东盟/ASEAN/东南亚”有关的合作，包括：合作高校、合作国家、交换生项目、联合培养项目、研学项目等。
3. 尽量用“保守但不低估”的方式给出以下3个字段的大致整数：

{metric_desc}

【重要说明】
- 只统计东盟10国：文莱、柬埔寨、印度尼西亚、老挝、马来西亚、缅甸、菲律宾、新加坡、泰国、越南。
- 如果无法确认某个指标，请尽量给出保守估计（0 或 1），不要随意编造极大数值。
- 如果学校官网有明确的东盟合作项目或东盟国家列表，请以官网内容为准。

【输出格式】
只输出一个 JSON 对象，字段必须是：
asean_partner_countries_count, asean_partner_universities_count, asean_program_count

示例：
{{
  "asean_partner_countries_count": 5,
  "asean_partner_universities_count": 12,
  "asean_program_count": 8
}}
"""


def kimi_complete_row(school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]) -> Dict[str, Any]:
    prompt = build_row_prompt(school_name, row_context, missing_metrics)

    resp = client.chat.completions.create(
        model=KIMI_MODEL_NAME,
        messages=[
            {"role": "system", "content": "你是一个严谨的数据提取助手，只输出 JSON。"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        max_tokens=1200,
        extra_body={"enable_search": True},
    )

    content = resp.choices[0].message.content
    if isinstance(content, list):
        # 兼容分段返回
        content = "".join(getattr(p, "text", "") for p in content)

    data = parse_json(content)
    if not isinstance(data, dict):
        return {}
    return data


# ========= 3. 主流程 =========


def main():
    print(">>> Step5-A9 - 东盟指标补全 + ARII 计算 开始")

    df = pd.read_excel(INPUT_FILE)
    print(f"读取数据：{df.shape[0]} 行 × {df.shape[1]} 列")

    if "school_name" not in df.columns:
        raise ValueError("缺少 school_name 列，无法按学校逐行处理。")

    # 如果没有东盟相关列，则先创建
    for col in ["asean_partner_countries_count", "asean_partner_universities_count", "asean_program_count"]:
        if col not in df.columns:
            df[col] = None

    # 作为上下文提供给 Kimi 的列（尽量多给一点信息）
    context_cols = [
        "school_name",
        "province",
        "city",
        "location",
        "official_site",
        # 下面这些列名需要你根据自己的表结构微调：
        "intl_coop_url_1",
        "intl_coop_url_2",
        "intl_coop_url_3",
        "overseas_exchange_url_1",
        "overseas_exchange_url_2",
    ]
    context_cols = [c for c in context_cols if c in df.columns]

    total = len(df)

    for idx, row in df.iterrows():
        school_name = str(row["school_name"])
        print(f"\n[{idx + 1}/{total}] 学校：{school_name}")

        # 判断这一行哪些东盟字段还缺
        missing = []
        for m in TARGET_METRICS.keys():
            v = row.get(m, None)
            if pd.isna(v) or v is None or v == "":
                missing.append(m)

        if not missing:
            print("  -> 东盟相关字段已存在，跳过。")
            continue

        # 构造上下文
        row_ctx = row[context_cols].to_dict()
        print(f"  缺失字段：{missing}")

        try:
            result = kimi_complete_row(school_name, row_ctx, missing)
        except Exception as e:
            print(f"  ⚠ 调用 Kimi 失败：{e}")
            continue

        print(f"  返回 JSON：{result}")

        for m in TARGET_METRICS.keys():
            if m in result and result[m] is not None and result[m] != "":
                df.at[idx, m] = result[m]

        # 特别打印一下广西外国语学院的情况
        if "广西外国语" in school_name:
            print("  >>> 当前广西外国语学院东盟指标：")
            print({k: df.at[idx, k] for k in TARGET_METRICS.keys()})

        # 控制请求频率
        time.sleep(SLEEP_BASE + random.random() * SLEEP_JITTER)

        # 每 20 行保存一次中间结果
        if (idx + 1) % 20 == 0:
            print("  保存中间结果...")
            df.to_excel(OUTPUT_FILE, index=False)

    # ========= 4. 计算 ARII =========
    print("\n>>> 所有学校东盟原始字段抓取完成，开始计算 ARII")

    for col in ["asean_partner_countries_count", "asean_partner_universities_count", "asean_program_count"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["asean_countries_norm"] = minmax(df["asean_partner_countries_count"])
    df["asean_univs_norm"] = minmax(df["asean_partner_universities_count"])
    df["asean_programs_norm"] = minmax(df["asean_program_count"])

    # ARII 权重可根据需要微调，这里先给一个合理方案
    df["ARII"] = (
        0.4 * df["asean_univs_norm"] +
        0.3 * df["asean_countries_norm"] +
        0.3 * df["asean_programs_norm"]
    )

    # 简单看看 ARII Top 20
    df_sorted = df.sort_values(by="ARII", ascending=False)
    print("\n=== ARII Top 20 学校（东盟区域国际化指数） ===")
    print(df_sorted[["school_name", "ARII",
                     "asean_partner_countries_count",
                     "asean_partner_universities_count",
                     "asean_program_count"]].head(20).to_string(index=False))

    # 看看广西外国语学院
    gx = df_sorted[df_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)]
    print("\n=== 广西外国语学院的东盟指数情况 ===")
    if gx.empty:
        print("未找到包含“广西外国语”的 school_name 记录。")
    else:
        print(gx[["school_name", "ARII",
                  "asean_partner_countries_count",
                  "asean_partner_universities_count",
                  "asean_program_count"]].to_string(index=False))

    # 保存最终带 ARII 的表
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存带有东盟指标与 ARII 的表格到：{OUTPUT_FILE}")
    print(">>> Step5-A9 - 完成")


if __name__ == "__main__":
    main()
