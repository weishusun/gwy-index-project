"""
Step5-A2：使用 Kimi 联网补全“语种数量 + 语种列表”
"""

import os
import json
import time
import random
from typing import Dict, Any, List

import pandas as pd
from openai import OpenAI

# ===================== 1. 配置区 =====================

# 建议：从环境变量读取，或者你自己手动改成与你 step4C 相同的 key
MOONSHOT_API_KEY = "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy"

MOONSHOT_API_BASE = "https://api.moonshot.cn/v1"
KIMI_MODEL_NAME = "kimi-k2-turbo-preview"

INPUT_EXCEL = "step5_intl_features_A1_language_and_majors.xlsx"
OUTPUT_EXCEL = "step5_intl_features_A2_kimi_languages.xlsx"

SLEEP_BASE = 3
SLEEP_JITTER = 2

# 这次只关心两个指标
TARGET_METRICS: Dict[str, Dict[str, str]] = {
    "languages_offered_count": {
        "desc": "学校目前开设的外国语言语种数量（不要算汉语相关，只算其他自然语言，如英语、日语、泰语等），请给出整数。",
    },
    "languages_list": {
        "desc": "学校目前开设的外国语言语种列表，用中文全称表示，用顿号（、）分隔，例如：英语、日语、韩语、泰语、越南语。",
    },
}

client = OpenAI(
    api_key=MOONSHOT_API_KEY,
    base_url=MOONSHOT_API_BASE,
)

# ===================== 2. Prompt & 解析 =====================

def build_row_prompt(school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]) -> str:
    """
    根据一所学校的已知信息，构造补全语种的 Prompt
    """
    ctx = json.dumps(row_context, ensure_ascii=False)

    metric_desc = "\n".join(
        [f"- {m}: {TARGET_METRICS[m]['desc']}" for m in missing_metrics]
    )

    return f"""
你是一名严谨的教育数据抽取助手，将通过联网搜索来补全一所中国民办本科高校的“外国语言语种相关”指标。

【已知信息（JSON）】
{ctx}

【任务】
学校名称：{school_name}

请你：
1. 访问该校的官网（可以结合招生简章、专业设置、国际学院介绍等页面）。
2. 识别该校开设了多少种“外国语言”相关的专业或课程（例如英语、日语、韩语、泰语、越南语、柬埔寨语、印尼语等），汉语相关方向不要计入。
3. 提取出这些语种的列表（中文全称）。

【需要补全的字段】
{metric_desc}

【输出要求】
只输出一个 JSON，对应字段是：languages_offered_count, languages_list。
不要输出解释性文字。
"""

def parse_json(text: str):
    text = text.strip()
    try:
        return json.loads(text)
    except:
        pass

    s = text.find("{")
    e = text.rfind("}")
    if s != -1 and e != -1:
        try:
            return json.loads(text[s:e+1])
        except:
            pass

    return {}

def kimi_complete_row(school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]) -> Dict[str, Any]:
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

    data = parse_json(content)
    if not isinstance(data, dict):
        return {}
    return data

# ===================== 3. 主流程 =====================

def main():
    print(f"读取：{INPUT_EXCEL}")
    df = pd.read_excel(INPUT_EXCEL)

    if "school_name" not in df.columns:
        raise ValueError("缺少 school_name 列")

    # 为结果增加列（如果不存在）
    for col in ["languages_offered_count", "languages_list"]:
        if col not in df.columns:
            df[col] = None

    # 这里提供给 Kimi 的上下文字段，可以根据需要增减
    context_cols = [
        "school_name",
        "location",
        "official_site",
        "major_program_url_1",
        "major_program_url_2",
        "major_count_final",
        "major_language_related",
        "foreign_major_count_final",
    ]
    context_cols = [c for c in context_cols if c in df.columns]

    for idx, row in df.iterrows():
        school_name = str(row["school_name"])
        print(f"\n[{idx+1}/{len(df)}] 学校：{school_name}")

        # 判断还有哪些指标需要补
        missing = []
        for m in TARGET_METRICS.keys():
            if pd.isna(row.get(m, None)) or row.get(m, None) in [None, ""]:
                missing.append(m)

        if not missing:
            print("  -> 所有指标已存在，跳过。")
            continue

        # 构造上下文
        ctx = row[context_cols].to_dict()
        print(f"  缺失字段：{missing}")

        try:
            result = kimi_complete_row(school_name, ctx, missing)
        except Exception as e:
            print(f"  ⚠ 调用 Kimi 失败：{e}")
            continue

        print(f"  返回 JSON：{result}")

        for m in TARGET_METRICS.keys():
            if m in result and result[m] not in [None, ""]:
                df.at[idx, m] = result[m]

        # 间隔一下，避免请求过快
        time.sleep(SLEEP_BASE + random.random() * SLEEP_JITTER)

        # 每 20 行保存一次中间结果
        if (idx + 1) % 20 == 0:
            print("  保存中间结果...")
            df.to_excel(OUTPUT_EXCEL, index=False)

    print("\n全部完成，保存最终结果。")
    df.to_excel(OUTPUT_EXCEL, index=False)

if __name__ == "__main__":
    main()
