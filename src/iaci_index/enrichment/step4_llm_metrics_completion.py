"""
Step4-C：使用 Kimi (Moonshot) 联网，按“学校”为单位补全多个缺失指标
本版本：直接在脚本中写死 API Key（你要求的版本）
"""

import os
import json
import time
import random
from typing import Dict, Any, List

import pandas as pd
from openai import OpenAI

# ===================== 1. 配置区 =====================

# 你要求：直接把 API Key 写在脚本中，不依赖环境变量
MOONSHOT_API_KEY = "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy"

# Kimi OpenAI 兼容接口
MOONSHOT_API_BASE = "https://api.moonshot.cn/v1"
KIMI_MODEL_NAME = "kimi-k2-turbo-preview"

# 输入 / 输出文件名
INPUT_EXCEL = "step4_merged_full_metrics_2025.xlsx"
OUTPUT_EXCEL = "step4C_kimi_completed_metrics_2025.xlsx"

# 想要联网补全的指标
TARGET_METRICS: Dict[str, Dict[str, str]] = {
    "employment_rate_2024": {
        "desc": "最近一届本科毕业生就业率（2024 届），例如 0.945 表示 94.5%。如果只看到 95%、94%以上之类，请尽量给出小数形式。",
    },
    "employment_rate_2025": {
        "desc": "最近一届本科毕业生就业率（2025 届），没有就返回 null。",
    },
    "further_study_rate_2025": {
        "desc": "最近一届本科毕业生升学率（含考研、出国深造），用 0-1 之间的小数表示。",
    },
    "intl_partner_universities_count": {
        "desc": "学校公开的海外合作高校数量（中外高校、合作院校数量），用整数表示。",
    },
    "intl_partner_countries_count": {
        "desc": "学校公开的合作国家或地区数量，用整数表示。",
    },
    "studyabroad_students_annual": {
        "desc": "每年出国（境）交流、访学、双学位、短期项目等人数。整数，模糊描述则给估计值。",
    },
}

# LLM 信息列后缀
EVIDENCE_SUFFIX = "_llm_evidence"
SOURCE_SUFFIX = "_llm_source"
CONFID_SUFFIX = "_llm_confidence"

# 防止调用太频繁
SLEEP_BASE = 2.0
SLEEP_JITTER = 1.0

# 初始化 Kimi 客户端
client = OpenAI(
    api_key=MOONSHOT_API_KEY,
    base_url=MOONSHOT_API_BASE,
)

# ===================== 2. Prompt =====================

def build_row_prompt(school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]) -> str:
    ctx = json.dumps(row_context, ensure_ascii=False)

    metric_desc = "\n".join(
        [f"- {m}: {TARGET_METRICS[m]['desc']}" for m in missing_metrics]
    )

    return f"""
你是一名严谨的教育数据抽取助手，现在要通过联网搜索来补全一所中国民办本科高校的若干指标。

【已知信息】
{ctx}

【学校】
{school_name}

【需要补全的指标】
{metric_desc}

请你联网搜索（enable_search=true），从官网与权威来源获取数据。
每个字段输出：
{{
  "value": 数值或 null,
  "evidence": "证据中文原文",
  "source_url": "来源链接",
  "confidence": "high/medium/low"
}}
以 JSON 格式返回，不要额外解释。
""".strip()


def parse_json(text: str):
    text = text.strip()
    try:
        return json.loads(text)
    except:
        pass

    # 尝试截取大括号
    s = text.find("{")
    e = text.rfind("}")
    if s != -1 and e != -1:
        try:
            return json.loads(text[s:e+1])
        except:
            pass

    return {}

# ===================== Kimi 调用 =====================

def kimi_complete_row(school_name, row_context, missing_metrics):
    prompt = build_row_prompt(school_name, row_context, missing_metrics)

    resp = client.chat.completions.create(
        model=KIMI_MODEL_NAME,
        messages=[
            {"role": "system", "content": "你是一个严谨的数据提取助手，只输出 JSON"},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        max_tokens=1200,
        extra_body={"enable_search": True},
    )

    content = resp.choices[0].message.content
    if isinstance(content, list):
        content = "".join(getattr(p, "text", "") for p in content)

    return parse_json(content)


# ===================== 3. 主流程 =====================

def main():
    print(f"读取：{INPUT_EXCEL}")
    df = pd.read_excel(INPUT_EXCEL)

    if "school_name" not in df.columns:
        raise ValueError("缺少 school_name 列")

    # 增加证据信息列
    for m in TARGET_METRICS:
        for suf in [EVIDENCE_SUFFIX, SOURCE_SUFFIX, CONFID_SUFFIX]:
            col = m + suf
            if col not in df.columns:
                df[col] = None

    context_cols = [
        "school_name",
        "students_total_final",
        "teachers_total_final",
        "major_count_final",
        "campus_area_m2_final",
    ]
    context_cols = [c for c in context_cols if c in df.columns]

    total = len(df)

    for idx, row in df.iterrows():
        school = str(row["school_name"])
        missing = [m for m in TARGET_METRICS if pd.isna(row.get(m))]

        if not missing:
            continue

        print(f"\n[{idx+1}/{total}] 学校：{school}")
        print("  缺失：", missing)

        ctx = row[context_cols].to_dict()

        try:
            result = kimi_complete_row(school, ctx, missing)
        except Exception as e:
            print(" 调用失败：", e)
            continue

        for m in missing:
            r = result.get(m, {}) or {}
            df.loc[idx, m] = r.get("value")
            df.loc[idx, m + EVIDENCE_SUFFIX] = r.get("evidence")
            df.loc[idx, m + SOURCE_SUFFIX] = r.get("source_url")
            df.loc[idx, m + CONFID_SUFFIX] = r.get("confidence")

            print(f"  {m} -> {r.get('value')} ({r.get('confidence')})")

        time.sleep(SLEEP_BASE + random.random() * SLEEP_JITTER)

        if (idx + 1) % 20 == 0:
            print("  保存中间结果")
            df.to_excel(OUTPUT_EXCEL, index=False)

    print("\n全部完成！保存最终结果。")
    df.to_excel(OUTPUT_EXCEL, index=False)


def run_step4() -> None:
    """Perform LLM-based completion of missing metrics."""
    main()


if __name__ == "__main__":
    run_step4()
