# -*- coding: utf-8 -*-
"""
Step4-A: 从本地 html_cache_step3 中离线抽取指标，
         补全 step3_private_undergrad_metrics_2025.xlsx

核心思路：
- 不再只依赖 profile_page_url（你目前是空的）
- 而是和 Step3 一样：
  - 读取该校所有 *_url* 类字段（official_site, info_url_*, disclosure_url_* 等）
  - 用 normalize_special_url() 规范化
  - 用 md5(规范化后的 url) 计算缓存文件名
  - 在 html_cache_step3/ 中读取对应 HTML
  - 拼成一大段文本交给大模型抽取结构化指标

使用方式：
1. 安装依赖：
   pip install pandas openpyxl beautifulsoup4 lxml openai

2. 配置环境变量（只需做一次）：
   Windows PowerShell:
       setx OPENAI_API_KEY "你的API_KEY"
   然后重新打开终端。

3. 运行：
   python step4_offline_metrics_from_cache.py
"""

import os
import json
import hashlib
from typing import Dict, Any, Optional, List, Tuple

import pandas as pd
from bs4 import BeautifulSoup
from openai import OpenAI

# ============= 配置区域 =============

EXCEL_PATH = "step3_private_undergrad_metrics_2025.xlsx"
HTML_CACHE_DIR = "html_cache_step3"

# 一次试跑先少一点，确认没问题后改成 None 跑全量
MAX_ROWS_PER_RUN: Optional[int] = None

# 每校最多拼接多少个页面的文本
MAX_PAGES_PER_SCHOOL = 5

# 单页 & 总文本长度限制（防止 token 爆炸）
MAX_CHARS_PER_PAGE = 5000
MAX_CHARS_TOTAL = 20000

METRIC_FIELDS: Dict[str, str] = {
    "founded_year": "建校年份（例如 2004）",
    "students_total": "在校生总人数（大概数量即可，例如 18000）",
    "junior_students_total": "在校专科（高职高专）学生总数（没有就返回 null）",
    "international_students": "在校国际学生人数（没有就返回 null）",
    "teachers_total": "教师总数（含专任、兼职，如未明确可估算或返回 null）",
    "fulltime_teachers": "专任教师人数（没有就返回 null）",
    "phd_teachers_count": "具有博士学位教师人数（没有就返回 null）",
    "master_teachers_count": "具有硕士学位教师人数（没有就返回 null）",
    "student_teacher_ratio": "生师比（格式如 18:1 或数值 18）",
    "major_count": "本科专业总数",
    "major_language_related": "外语类相关本科专业数量",
    "major_business_related": "商科/管理类相关本科专业数量",
    "national_first_class_majors": "国家级一流本科专业建设点数量",
    "provincial_first_class_majors": "省级一流本科专业建设点数量",
    "campus_count": "校区数量（含分校区）",
    "campus_area_mu": "校园占地面积（单位：亩，如网页只有平方米，可换算）",
    "campus_area_m2": "校园占地面积（单位：平方米）",
    "library_books": "纸质藏书册数",
    "library_ebooks": "电子图书册数（或种数）",
    "off_campus_bases_count": "校外实习实训基地数量",
    "studyabroad_students_annual": "每年出国（境）交流/留学学生人数（大概即可）",
}

EVIDENCE_SUFFIX = "_evidence"
CONFIDENCE_SUFFIX = "_confidence"

OPENAI_MODEL = "gpt-4o-mini"
client = OpenAI()

# ============= 与 Step3 保持一致的工具函数 =============

def cache_key_for_url(url: str) -> str:
    """根据 URL 生成简单文件名（与 Step3 完全一致）"""
    h = hashlib.md5(url.encode("utf-8")).hexdigest()
    return h + ".html"


def normalize_special_url(url: str) -> Optional[str]:
    """
    从 Step3 复制过来的 URL 规范化逻辑（简化版）：
    - 丢掉 javascript:/mailto:/tel:
    - 三亚学院统一认证特殊修正
    """
    if not isinstance(url, str):
        return None
    u = url.strip()
    if not u:
        return None

    lower = u.lower()
    if lower.startswith("javascript:") or lower.startswith("mailto:") or lower.startswith("tel:"):
        return None

    # Step3 里的特殊规则：id.sanyau.edu.cn:9092 -> https://id.sanyau.edu.cn/
    if "id.sanyau.edu.cn:9092" in u:
        return "https://id.sanyau.edu.cn/"

    return u


def load_html_text(path: str, max_chars: int = MAX_CHARS_PER_PAGE) -> str:
    """读取 HTML 文件并提取正文文本，限制最大长度。"""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            html = f.read()
    except UnicodeDecodeError:
        with open(path, "r", encoding="gb18030", errors="ignore") as f:
            html = f.read()

    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    lines = [line.strip() for line in text.splitlines()]
    text = "\n".join(line for line in lines if line)

    if len(text) > max_chars:
        text = text[:max_chars]

    return text


def collect_candidate_urls_from_row(row: pd.Series) -> List[Tuple[str, str, str]]:
    """
    收集该行可用的 URL 列：
    - official_site
    - 所有列名中满足：以 '_url' 结尾 或 包含 '_url_'
    返回列表[(列名, 原始值, 规范化后 url)]
    """
    results: List[Tuple[str, str, str]] = []

    # official_site
    official = row.get("official_site", None)
    if isinstance(official, str) and official.strip():
        nv = normalize_special_url(official.strip())
        if nv:
            results.append(("official_site", official.strip(), nv))

    # 其它 *_url* 列
    for col in row.index:
        if not isinstance(col, str):
            continue
        if col == "official_site":
            continue
        if col.endswith("_url") or "_url_" in col:
            val = row.get(col, None)
            if isinstance(val, str) and val.strip():
                nv = normalize_special_url(val.strip())
                if nv:
                    results.append((col, val.strip(), nv))

    # 按规范化 URL 去重
    seen = set()
    deduped: List[Tuple[str, str, str]] = []
    for col, raw, norm in results:
        if norm in seen:
            continue
        seen.add(norm)
        deduped.append((col, raw, norm))

    return deduped


def build_merged_text_for_row(row: pd.Series) -> str:
    """
    根据该行的所有 *_url* 字段，从缓存中读取若干页面，
    拼成一个大文本给 LLM 使用。
    """
    parts: List[str] = []
    candidates = collect_candidate_urls_from_row(row)

    for col, raw_url, norm_url in candidates:
        fname = cache_key_for_url(norm_url)
        path = os.path.join(HTML_CACHE_DIR, fname)
        if not os.path.exists(path):
            continue

        try:
            text = load_html_text(path)
        except Exception:
            continue

        if not text.strip():
            continue

        header = f"\n\n===== 来源列: {col} | URL: {norm_url} =====\n"
        parts.append(header + text)

        if len(parts) >= MAX_PAGES_PER_SCHOOL:
            break

    if not parts:
        return ""

    merged = "\n".join(parts)
    if len(merged) > MAX_CHARS_TOTAL:
        merged = merged[:MAX_CHARS_TOTAL]

    return merged


def build_prompt(text: str, school_name: str) -> str:
    fields_desc = "\n".join(
        [f"- {col}: {desc}" for col, desc in METRIC_FIELDS.items()]
    )

    prompt = f"""
你是一名严谨的数据抽取助手，正在阅读一所高校的若干官网页面文本（已经转为纯文本）。
学校名称：{school_name}

这些文本来自学校概况、信息公开、就业质量报告、国际合作等多个子页面。

请从下面的文本中，尽可能抽取以下指标（如果文本中完全找不到，就返回 null）：

{fields_desc}

要求：
1. 不要凭空猜测，只能根据文本里的数字和语句来判断。
2. 如果有多个地方提到不同数字，请优先选择最新或最权威的表述（例如“截至2024年，学校有……人”）。
3. 对每个字段，你需要返回：
   - value：数值或字符串（无法确定就 null）
   - evidence：支持这个数值的原文片段（尽量简短）
   - confidence：high / medium / low（表示你对该字段的把握程度）

最后，请严格输出一个 JSON 对象，不要加入任何多余说明，结构如下：

{{
  "metrics": {{
    "字段名1": {{
      "value": ...,
      "evidence": "...",
      "confidence": "high/medium/low"
    }},
    "字段名2": {{
      ...
    }},
    ...
  }}
}}

下面是网页文本（可能较长）：
----------------- 文本开始 -----------------
{text}
----------------- 文本结束 -----------------
"""
    return prompt


def call_llm_extract_metrics(text: str, school_name: str) -> Optional[Dict[str, Any]]:
    prompt = build_prompt(text, school_name)

    resp = client.chat.completions.create(
        model=OPENAI_MODEL,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": "你是一个负责高校官网数据抽取的助手，只能返回 JSON。",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
    )

    content = resp.choices[0].message.content
    try:
        data = json.loads(content)
        return data
    except json.JSONDecodeError:
        print("⚠️ JSON 解析失败，原始返回：")
        print(content[:4000])
        return None


def is_empty_value(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


# ============= 主流程 =============

def main():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel 文件不存在：{EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH)

    # 确保所有需要的列存在
    for col in METRIC_FIELDS.keys():
        if col not in df.columns:
            df[col] = None
        ev_col = col + EVIDENCE_SUFFIX
        if ev_col not in df.columns:
            df[ev_col] = None
        conf_col = col + CONFIDENCE_SUFFIX
        if conf_col not in df.columns:
            df[conf_col] = None

    if "metrics_status" not in df.columns:
        df["metrics_status"] = "missing"

    rows_to_process: List[Tuple[int, str]] = []

    for idx, row in df.iterrows():
        # 1. 至少有一个指标是空的
        has_empty_metric = any(
            is_empty_value(row.get(col, None)) for col in METRIC_FIELDS.keys()
        )
        if not has_empty_metric:
            continue

        # 2. 找到该校可用的缓存文本
        merged_text = build_merged_text_for_row(row)
        if not merged_text.strip():
            continue

        rows_to_process.append((idx, merged_text))

    if MAX_ROWS_PER_RUN is not None:
        rows_to_process = rows_to_process[:MAX_ROWS_PER_RUN]

    print(f"共发现 {len(rows_to_process)} 行需要尝试离线抽取。")

    for idx, merged_text in rows_to_process:
        row = df.loc[idx]
        school_name = str(row.get("school_name", "")).strip()
        print(f"\n====== 处理第 {idx} 行：{school_name} ======")

        llm_result = call_llm_extract_metrics(merged_text, school_name)
        if not llm_result or "metrics" not in llm_result:
            print("大模型未返回 metrics 字段，跳过。")
            continue

        metrics: Dict[str, Any] = llm_result["metrics"]

        filled_any = False
        for col in METRIC_FIELDS.keys():
            m = metrics.get(col)
            if not isinstance(m, dict):
                continue

            value = m.get("value", None)
            evidence = m.get("evidence", None)
            confidence = m.get("confidence", None)

            if is_empty_value(row.get(col, None)) and not is_empty_value(value):
                df.at[idx, col] = value
                filled_any = True

            ev_col = col + EVIDENCE_SUFFIX
            conf_col = col + CONFIDENCE_SUFFIX
            if not is_empty_value(evidence):
                df.at[idx, ev_col] = evidence
            if not is_empty_value(confidence):
                df.at[idx, conf_col] = confidence

        if filled_any:
            old_status = str(df.at[idx, "metrics_status"]).strip().lower()
            if old_status in ("missing", "", "nan"):
                df.at[idx, "metrics_status"] = "partial"

        print("✔ 已尝试写回该行。")

    backup_path = EXCEL_PATH.replace(".xlsx", "_backup_before_step4A.xlsx")
    if not os.path.exists(backup_path):
        df.to_excel(backup_path, index=False)
        print(f"\n已生成备份文件：{backup_path}")

    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ 已写回 Excel：{EXCEL_PATH}")


def run_offline_cache_fill() -> None:
    """Re-run metric extraction from cached HTML."""
    main()


if __name__ == "__main__":
    run_offline_cache_fill()
