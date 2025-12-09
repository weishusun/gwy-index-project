"""
Step5-A10-LLM：用 Kimi 联网补全文本国际化信息，并计算 TLI（文本国际化指数）

输入：step5_A9_asean_features.xlsx（已含 LRI, ICI, ARII）
输出：step5_A10_tli_llm_features.xlsx（新增 llm_intl_* 字段和 TLI）
"""

import os
import time
import random
import json
from typing import Dict, Any, List

import pandas as pd
from openai import OpenAI

INPUT_FILE = "step5_A9_asean_features.xlsx"
OUTPUT_FILE = "step5_A10_tli_llm_features.xlsx"

# 你可以改成跟 step4C 一样的方式，从环境变量读取
MOONSHOT_API_KEY = "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy"
MOONSHOT_API_BASE = "https://api.moonshot.cn/v1"
KIMI_MODEL_NAME = "kimi-k2-turbo-preview"

SLEEP_BASE = 3
SLEEP_JITTER = 2

INTL_KEYWORDS = [
    "国际化", "国际", "全球", "全球化", "国际视野",
    "外国语", "外语", "多语种", "多语言", "跨文化", "跨国",
    "国际合作", "海外交流", "境外交流", "访学", "交换生",
    "留学生", "国际学生",
    "东盟", "东南亚", "RCEP", "一带一路",
]

client = OpenAI(
    api_key=MOONSHOT_API_KEY,
    base_url=MOONSHOT_API_BASE,
)


def minmax(s: pd.Series) -> pd.Series:
    s = s.astype(float)
    if s.max() == s.min():
        return s * 0
    return (s - s.min()) / (s.max() - s.min())


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


def build_prompt(school_name: str, context: Dict[str, Any]) -> str:
    ctx_json = json.dumps(context, ensure_ascii=False)
    return f"""
你是一名教育数据研究助手，需要通过联网搜索了解一所中国民办本科高校的“国际化办学定位”。

【学校名称】
{school_name}

【已知基础信息（JSON）】
{ctx_json}

【任务目标】
1. 访问学校官网以及与“国际交流、国际合作、东盟、国际学院、港澳台与国际教育”等相关页面。
2. 综合公开信息，给出该校在“国际化办学”方面的简要描述，以及若干关键词。
3. 特别关注该校是否强调：
   - 多语种 / 外国语特色
   - 国际合作、出国交流、留学生培养
   - 东盟 / 东南亚 / 一带一路等区域布局

【需要你输出的字段】
只输出一个 JSON 对象，字段如下：
- llm_intl_summary: 用 80~150 字总结学校在国际化办学方面的定位与特点（中文）。
- llm_intl_keywords: 提炼 6~12 个与国际化相关的关键词，用顿号或逗号分隔。
- llm_asean_keywords: 如果有与东盟/东南亚相关的内容，也给出 3~8 个关键词；若基本没有，则输出空字符串 ""。

【输出示例】
{{
  "llm_intl_summary": "学校以多语种、应用型外语人才培养为特色，强调服务区域开放发展，构建了涵盖英语、东盟语种等多语言的人才培养体系，并与多国高校开展长期合作交流。",
  "llm_intl_keywords": "国际化办学, 多语种, 外国语特色, 国际合作, 海外交流, 留学生培养, 一带一路",
  "llm_asean_keywords": "东盟, 东南亚, 泰国高校合作, 越南高校合作"
}}
"""


def kimi_fetch_tli_info(school_name: str, context: Dict[str, Any]) -> Dict[str, Any]:
    prompt = build_prompt(school_name, context)

    resp = client.chat.completions.create(
        model=KIMI_MODEL_NAME,
        messages=[
            {"role": "system", "content": "你是一个严谨的数据提取助手，只输出 JSON。"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.2,
        max_tokens=1000,
        extra_body={"enable_search": True},
    )

    content = resp.choices[0].message.content
    if isinstance(content, list):
        content = "".join(getattr(p, "text", "") for p in content)

    data = parse_json(content)
    if not isinstance(data, dict):
        return {}
    return data


def compute_tli_from_text(summary: str, intl_kw: str, asean_kw: str) -> float:
    text = ""
    for part in [summary, intl_kw, asean_kw]:
        if isinstance(part, str):
            text += part + "\n"

    if not text.strip():
        return 0.0

    total = 0
    hits = 0
    for kw in INTL_KEYWORDS:
        c = text.count(kw)
        total += c
        if c > 0:
            hits += 1

    score = total + 0.5 * hits
    return float(score)


def main():
    print(">>> Step5-A10-LLM - 用大模型补全文本信息并计算 TLI 开始")

    df = pd.read_excel(INPUT_FILE)
    print(f"读取：{df.shape[0]} 行 × {df.shape[1]} 列")

    if "school_name" not in df.columns:
        raise ValueError("缺少 school_name 列")

    # 作为上下文给模型的字段（不强制都有）
    context_cols = [
        "school_name", "province", "city", "location",
        "official_site", "intl_coop_url_1", "intl_coop_url_2",
        "intl_coop_url_3", "studyabroad_url_1", "studyabroad_url_2",
        "asean_url_1", "asean_url_2",
    ]
    context_cols = [c for c in context_cols if c in df.columns]

    # 如果之前已经有 llm_intl_summary，就不重复请求
    if "llm_intl_summary" not in df.columns:
        df["llm_intl_summary"] = None
    if "llm_intl_keywords" not in df.columns:
        df["llm_intl_keywords"] = None
    if "llm_asean_keywords" not in df.columns:
        df["llm_asean_keywords"] = None

    total = len(df)

    for idx, row in df.iterrows():
        school_name = str(row["school_name"])
        print(f"\n[{idx+1}/{total}] 学校：{school_name}")

        if isinstance(row.get("llm_intl_summary", None), str) and row["llm_intl_summary"].strip():
            print("  -> 已有 llm_intl_summary，跳过调用。")
            continue

        ctx = row[context_cols].to_dict()

        try:
            result = kimi_fetch_tli_info(school_name, ctx)
        except Exception as e:
            print(f"  ⚠ 调用 Kimi 失败：{e}")
            continue

        print(f"  返回JSON：{result}")

        for k in ["llm_intl_summary", "llm_intl_keywords", "llm_asean_keywords"]:
            if k in result and result[k] is not None:
                df.at[idx, k] = result[k]

        if "广西外国语" in school_name:
            print("  >>> 当前广西外国语学院 LLM 文本：")
            print(df.at[idx, "llm_intl_summary"])
            print(df.at[idx, "llm_intl_keywords"])
            print(df.at[idx, "llm_asean_keywords"])

        time.sleep(SLEEP_BASE + random.random() * SLEEP_JITTER)

        if (idx + 1) % 20 == 0:
            print("  保存中间结果...")
            df.to_excel(OUTPUT_FILE, index=False)

    print("\n>>> 所有学校 LLM 文本补全完成，开始计算 TLI")

    df["raw_tli_score_llm"] = df.apply(
        lambda r: compute_tli_from_text(
            r.get("llm_intl_summary", ""),
            r.get("llm_intl_keywords", ""),
            r.get("llm_asean_keywords", ""),
        ),
        axis=1,
    )

    df["TLI"] = minmax(df["raw_tli_score_llm"])

    df_sorted = df.sort_values("TLI", ascending=False)
    print("\n=== LLM-TLI Top 20 学校 ===")
    print(df_sorted[["school_name", "raw_tli_score_llm", "TLI"]].head(20).to_string(index=False))

    gx = df_sorted[df_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)]
    print("\n=== 广西外国语学院的 LLM-TLI 情况 ===")
    print(gx[["school_name", "raw_tli_score_llm", "TLI"]].to_string(index=False))

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存带有 LLM 文本与 TLI 的表到：{OUTPUT_FILE}")
    print(">>> Step5-A10-LLM 完成")


if __name__ == "__main__":
    main()
