import json
import os
import random
import re
import time
from typing import Any, Dict, List

import numpy as np
import pandas as pd
from openai import OpenAI

INPUT_FILE_A1 = "step4C_kimi_completed_metrics_2025.xlsx"
OUTPUT_FILE_A1 = "step5_intl_features_A1_language_and_majors.xlsx"

MOONSHOT_API_KEY = "sk-PK4uOTEqNwfrVmKsP6ew5lMFmPtfqIN9uJMEKHED52VnsiPy"
MOONSHOT_API_BASE = "https://api.moonshot.cn/v1"
KIMI_MODEL_NAME = "kimi-k2-turbo-preview"

INPUT_EXCEL_A2 = "step5_intl_features_A1_language_and_majors.xlsx"
OUTPUT_EXCEL_A2 = "step5_intl_features_A2_kimi_languages.xlsx"
SLEEP_BASE = 3
SLEEP_JITTER = 2

TARGET_METRICS: Dict[str, Dict[str, str]] = {
    "languages_offered_count": {
        "desc": "学校目前开设的外国语言语种数量（不要算汉语相关，只算其他自然语言，如英语、日语、泰语等），请给出整数。",
    },
    "languages_list": {
        "desc": "学校目前开设的外国语言语种列表，用中文全称表示，用顿号（、）分隔，例如：英语、日语、韩语、泰语、越南语。",
    },
}

INPUT_FILE_A3 = "step5_intl_features_A2_kimi_languages.xlsx"

def split_and_count_languages(text: str) -> int:
    """
    尝试从 languages_list 文本中拆分出语种数量。
    支持分隔符：、 / ， / , / ; / 、空格 等。
    """
    if not isinstance(text, str):
        return np.nan

    # 替换常见分隔符为统一逗号
    seps = ['、', '，', ';', '；', '/', '\\', '|']
    for s in seps:
        text = text.replace(s, ',')
    # 再按逗号拆分
    parts = [p.strip() for p in text.split(',') if p.strip()]

    if len(parts) == 0:
        return np.nan
    return len(parts)


def build_language_feature_a1() -> None:
    print(">>> Step5-A1 - 语种数量 & 外语类专业数量 特征构建 开始")

    # 1. 读取原始表
    df = pd.read_excel(INPUT_FILE_A1)
    print(f"读取原始数据：{df.shape[0]} 行 × {df.shape[1]} 列")

    # 检查需要的列是否存在
    needed_cols = ["languages_offered_count", "languages_list", "major_language_related"]
    for col in needed_cols:
        if col not in df.columns:
            print(f"⚠ 警告：未在表中找到列 {col}，后续该字段相关的特征会为空。")

    # 2. 构建 language_count_final
    if "languages_offered_count" in df.columns:
        lang_count = pd.to_numeric(df["languages_offered_count"], errors="coerce")
    else:
        lang_count = pd.Series([np.nan] * len(df))

    # 对没有 languages_offered_count 的，尝试从 languages_list 推断
    if "languages_list" in df.columns:
        inferred = df["languages_list"].apply(split_and_count_languages)
        # 仅在原来 lang_count 为空时用 inferred 填补
        lang_count = lang_count.fillna(inferred)

    df["language_count_final"] = lang_count

    # 3. 构建 foreign_major_count_final（外语类专业数量）
    if "major_language_related" in df.columns:
        foreign_major = pd.to_numeric(df["major_language_related"], errors="coerce")
    else:
        foreign_major = pd.Series([np.nan] * len(df))

    df["foreign_major_count_final"] = foreign_major

    # 4. 简单统计信息
    print("\n=== 语种数量（language_count_final）统计 ===")
    print(df["language_count_final"].describe())
    print(f"非空学校数量：{df['language_count_final'].notna().sum()} / {len(df)}")

    print("\n=== 外语类专业数量（foreign_major_count_final）统计 ===")
    print(df["foreign_major_count_final"].describe())
    print(f"非空学校数量：{df['foreign_major_count_final'].notna().sum()} / {len(df)}")

    # 5. 保存结果
    df.to_excel(OUTPUT_FILE_A1, index=False)
    print(f"\n已保存带有新特征的表到：{OUTPUT_FILE_A1}")
    print(">>> Step5-A1 - 完成")


def build_row_prompt_a2(
    school_name: str, row_context: Dict[str, Any], missing_metrics: List[str]
) -> str:
    ctx = json.dumps(row_context, ensure_ascii=False)
    metric_desc = "\n".join(
        [f"- {m}: {TARGET_METRICS[m]['desc']}" for m in missing_metrics]
    )
    return f"""
你是一名严谨的教育数据抽取助手，将通过联网搜索来补全一所中国民办本科高校的“外国语言语种相关”指标。

【已知信息（JSON）】
{ctx}

【需要补全的指标】
{metric_desc}

【注意】
- 请通过“多渠道搜索 + 核实后”再输出。
- 若网上信息缺失或模糊，请给出“最有可能的估计”和1句话解释。
- 数据请以 JSON 格式返回，例如：{{"languages_offered_count": 5, "languages_list": "英语、日语、韩语、泰语、越南语"}}。

现在请补全：{school_name}
"""


def parse_llm_response_a2(resp_text: str, missing_metrics: List[str]) -> Dict[str, Any]:
    try:
        data = json.loads(resp_text)
        if isinstance(data, dict):
            return {k: data.get(k) for k in missing_metrics}
    except json.JSONDecodeError:
        pass

    result = {m: None for m in missing_metrics}
    for m in missing_metrics:
        pattern = rf"{m}\s*[:：]\s*([^\n，,。；;]+)"
        m_obj = re.search(pattern, resp_text)
        if m_obj:
            result[m] = m_obj.group(1).strip()
    return result


def _build_kimi_client() -> OpenAI:
    return OpenAI(
        api_key=os.getenv("MOONSHOT_API_KEY", MOONSHOT_API_KEY),
        base_url=MOONSHOT_API_BASE,
    )


def build_language_feature_a2() -> None:
    print(">>> Step5-A2 - 使用 Kimi 联网补全语种指标")
    df = pd.read_excel(INPUT_EXCEL_A2)
    print(f"读取 A2 输入：{df.shape[0]} 行 × {df.shape[1]} 列")

    messages_template = [
        {"role": "system", "content": "你是教育领域的严谨数据抽取助手。"},
    ]
    client = _build_kimi_client()

    for idx, row in df.iterrows():
        missing = [m for m in TARGET_METRICS if pd.isna(row.get(m))]
        if not missing:
            continue

        prompt = build_row_prompt_a2(row["school_name"], row.to_dict(), missing)
        messages = messages_template + [{"role": "user", "content": prompt}]

        try:
            resp = client.chat.completions.create(
                model=KIMI_MODEL_NAME,
                messages=messages,
                temperature=0.3,
            )
            reply = resp.choices[0].message.content.strip()
        except Exception as e:
            print(f"[WARN] 调用 Kimi 失败，跳过 {row['school_name']}：{e}")
            continue

        parsed = parse_llm_response_a2(reply, missing)
        for k, v in parsed.items():
            if pd.isna(row.get(k)):
                df.at[idx, k] = v

        sleep_time = random.uniform(SLEEP_BASE, SLEEP_BASE + SLEEP_JITTER)
        print(f"  已处理 {row['school_name']}，等待 {sleep_time:.1f}s")
        time.sleep(sleep_time)

        if (idx + 1) % 10 == 0:
            df.to_excel(OUTPUT_EXCEL_A2, index=False)
            print("  已写入中间结果。")

    df.to_excel(OUTPUT_EXCEL_A2, index=False)
    print(f"✅ 已保存到：{OUTPUT_EXCEL_A2}")
    print(">>> Step5-A2 - 完成")


def inspect_language_features() -> None:
    print(">>> Step5-A3 - 检查语种数量分布 & 异常值")
    df = pd.read_excel(INPUT_FILE_A3)

    if "languages_offered_count" not in df.columns:
        raise ValueError("表中没有 languages_offered_count 列")

    print("\n=== 语种数量（languages_offered_count）描述统计 ===")
    print(df["languages_offered_count"].describe())

    top = df.sort_values(by="languages_offered_count", ascending=False)
    print("\n=== 语种数量 Top 20 学校 ===")
    print(top[["school_name", "languages_offered_count"]].head(20).to_string(index=False))

    mask_gx = df["school_name"].astype(str).str.contains("广西外国语", na=False)
    gx = df[mask_gx]
    print("\n=== 广西外国语相关学校的语种情况 ===")
    if gx.empty:
        print("未找到包含“广西外国语”的 school_name。")
    else:
        print(
            gx[
                [
                    "school_name",
                    "languages_offered_count",
                    "languages_list",
                    "foreign_major_count_final",
                ]
            ].to_string(index=False)
        )

    outliers = df[df["languages_offered_count"] >= 20]
    print("\n=== 疑似异常的高语种数量学校（>=20） ===")
    if outliers.empty:
        print("没有 >=20 的语种数量记录。")
    else:
        print(outliers[["school_name", "languages_offered_count", "languages_list"]].to_string(index=False))

    print("\n>>> Step5-A3 - 完成")


def build_lri_features() -> None:
    """Run the language-related feature engineering steps (A1–A3)."""
    build_language_feature_a1()
    build_language_feature_a2()
    inspect_language_features()


if __name__ == "__main__":
    build_lri_features()
