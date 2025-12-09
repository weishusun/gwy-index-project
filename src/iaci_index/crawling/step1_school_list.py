# -*- coding: utf-8 -*-
"""
Step3 初始化脚本（强化版）：
根据 Step1 + Step2，创建 2025 年民办本科院校指标总表 step3_private_undergrad_metrics_2025.xlsx

特点：
1）以 Step2（带官网及各类 *_url_*）为主体行索引，确保后续爬虫完全对齐；
2）如果存在 Step1，则自动 merge 省份 / 城市 / 学校类型等基础字段；
3）预先创建尽可能完整的指标列（基础办学 + 国际化 + 就业 + 东盟 + 语种结构 + NLP 文本 + 技术列），
   方便后续所有 Step3 / Step3b / Step4 直接在同一张表上填数据。

依赖：
    pip install pandas openpyxl
"""

import os
import pandas as pd

# ---------------- 配置区：按需修改 ----------------

# Step1：基础名单（可选，如果不存在会跳过）
STEP1_FILE = "step1_private_undergrad.xlsx"

# Step2：已完成官网 + 各类功能链接采集的结果
STEP2_FILE = "step2_private_undergrad_with_urls_selenium.xlsx"

# Step3 输出：2025 指标空表
STEP3_FILE = "step3_private_undergrad_metrics_2025.xlsx"

TARGET_YEAR = 2025


def load_step2():
    if not os.path.exists(STEP2_FILE):
        raise FileNotFoundError(f"找不到 Step2 文件：{STEP2_FILE}")
    df2 = pd.read_excel(STEP2_FILE)

    if "school_name" not in df2.columns:
        raise ValueError(f"{STEP2_FILE} 中必须包含列 'school_name'。")

    # 统一去掉首尾空格
    df2["school_name"] = df2["school_name"].astype(str).str.strip()
    return df2


def maybe_load_step1():
    """
    尝试读取 Step1，如果不存在则返回 None。
    Step1 里常见字段示例：
        - school_name
        - province
        - city
        - school_type       (民办 / 公办 / 独立学院 / 转设等)
        - level             (本科 / 高职等)
        - supervising_dept  (主管部门)
    你实际文件里的字段名如果不同，可以在下面做一次映射。
    """
    if not os.path.exists(STEP1_FILE):
        print(f"[INFO] 未发现 Step1 文件 {STEP1_FILE}，跳过基础字段合并。")
        return None

    df1 = pd.read_excel(STEP1_FILE)
    if "school_name" not in df1.columns:
        print(f"[WARN] {STEP1_FILE} 中不含 'school_name' 列，跳过合并。")
        return None

    df1["school_name"] = df1["school_name"].astype(str).str.strip()
    return df1


def init_step3_table():
    # 1. 读取 Step2（主体）
    df2 = load_step2()

    # 2. 尝试合并 Step1 的基础信息
    df1 = maybe_load_step1()
    if df1 is not None:
        # 为了兼容不同命名，这里做一层“建议字段名映射”，不存在就忽略
        rename_map = {}
        # 如果 Step1 中有这些常见列，就统一改名成我们内部使用的字段
        for src, dst in [
            ("省份", "province"),
            ("省", "province"),
            ("城市", "city"),
            ("学校类型", "school_type"),
            ("办学性质", "school_type"),
            ("层次", "school_level"),
            ("类别", "school_category"),
            ("主管部门", "supervising_dept"),
        ]:
            if src in df1.columns and dst not in df1.columns:
                rename_map[src] = dst
        if rename_map:
            df1 = df1.rename(columns=rename_map)

        base_cols = [
            "province",
            "city",
            "school_type",
            "school_level",
            "school_category",
            "supervising_dept",
        ]
        keep_cols = ["school_name"] + [c for c in base_cols if c in df1.columns]
        df1_sub = df1[keep_cols].drop_duplicates(subset=["school_name"])
        df = df2.merge(df1_sub, on="school_name", how="left")
        print(f"[INFO] 已从 Step1 中合并基础字段：{[c for c in base_cols if c in df1.columns]}")
    else:
        df = df2.copy()

    # 3. 添加 year 列
    df["year"] = TARGET_YEAR

    # 4. 预先定义各类指标列
    # 4.1 基础信息类（部分可能来自 Step1/手动补全）
    basic_info_cols = [
        "school_full_name",      # 学校全称（以后可以从官网/民教网等修正）
        "school_short_name",     # 学校简称
        "english_name",          # 英文名
        "school_type",           # 办学性质：民办 / 转设 / 独立学院…
        "school_level",          # 办学层次：普通本科 / 高水平应用型学院 等
        "school_category",       # 学科类别：外语类 / 财经类 / 理工类 / 综合等
        "supervising_dept",      # 主管部门（教育厅 / 集团 / 无）
        "province",              # 省份
        "city",                  # 城市
        "founded_year",          # 建校年份
        "campus_count",          # 校区数量
    ]

    # 4.2 规模类
    scale_cols = [
        "students_total",         # 在校生总人数
        "undergrads_total",       # 本科生人数
        "junior_students_total",  # 专科生人数（如有）
        "postgrads_total",        # 研究生人数（民办里一般较少，但保留列）
        "international_students", # 留学生人数
        "annual_new_enrollment",  # 年招生规模（本专科合计）
    ]

    # 4.3 师资与办学实力
    faculty_cols = [
        "teachers_total",          # 教职工总人数
        "fulltime_teachers",       # 专任教师人数
        "professors_count",        # 正高人数
        "associate_professors_count",  # 副高人数
        "phd_teachers_count",      # 博士学位教师人数
        "master_teachers_count",   # 硕士学位教师人数
        "student_teacher_ratio",   # 师生比（数值，例如 18.5 表示 18.5:1）
        "college_count",           # 二级学院数量
        "department_count",        # 系 / 部数量
        "major_count",             # 本科专业数量
        "major_language_related",  # 语言类专业数量（外语、翻译等）
        "major_business_related",  # 商科经管类专业数量
        "national_first_class_majors",   # 国家一流本科专业数
        "provincial_first_class_majors", # 省级一流本科专业数
    ]

    # 4.4 办学资源
    resource_cols = [
        "campus_area_mu",        # 占地面积（亩）
        "campus_area_m2",        # 占地面积（平方米）
        "building_area_m2",      # 校舍建筑面积
        "library_books",         # 馆藏纸质图书
        "library_ebooks",        # 电子图书数量
        "labs_count",            # 实验室数量
        "training_bases_count",  # 校内实验实训基地数量
        "off_campus_bases_count" # 校外实习实践基地数量
    ]

    # 4.5 国际化 / 东盟 / 语种结构（后续 Step3b 用）
    intl_cols = [
        "intl_partner_universities_count",    # 海外合作高校数量
        "intl_partner_countries_count",       # 合作国家数量
        "intl_exchange_programs_count",       # 交换/访学项目数量
        "intl_double_degree_programs_count",  # 联合培养 / 双学位项目数量
        "studyabroad_students_annual",        # 每年出国交换/留学学生人数
        "overseas_internship_programs_count", # 海外实习项目数量
        "asean_partner_count",                # 东盟方向合作高校数量
        "asean_programs_count",               # 东盟 / RCEP 相关项目数量
        "languages_offered_count",            # 开设语种数量
        "languages_list",                     # 开设语种列表（字符串）
    ]

    # 4.6 就业 / 产教融合
    employment_cols = [
        "employment_rate_2024",        # 2024 年就业率（如果能拿到）
        "employment_rate_2025",        # 2025 年就业率（目标年份）
        "further_study_rate_2025",     # 2025 年升学率 / 深造率
        "top_employment_regions_2025", # 就业地区分布（文字）
        "top_employment_industries_2025", # 就业行业分布（文字）
        "top_employment_employers_2025",  # 主要雇主 / 企业（文字）
        "employment_quality_index",   # 自定义就业质量综合指数（后续 Step4 用）
    ]

    # 4.7 NLP 文本相关（概况 / 国际化 / 就业 / 东盟专用）
    nlp_cols = [
        "profile_page_url",            # 学校简介 / 概况页 URL
        "profile_text_snippet",        # 学校简介文本片段
        "intl_page_url",               # 国际合作 / 留学栏目 URL
        "intl_text_snippet",           # 国际化相关文本片段
        "employment_page_url",         # 就业质量报告 / 就业栏目 URL
        "employment_text_snippet",     # 就业相关文本片段
        "asean_page_url",              # 东盟 / 区域合作栏目 URL
        "asean_text_snippet",          # 东盟/区域合作相关文本片段
        "positioning_keywords",        # 办学定位标签（由简介 NLP 抽取）
        "intl_keywords",               # 国际化标签（Step3b NLP 抽取）
        "employment_keywords",         # 就业标签（Step3b NLP 抽取）
        "asean_keywords",              # 东盟/区域价值标签
    ]

    # 4.8 技术 / 状态控制（断点续跑 & 质量控制）
    tech_cols = [
        "metrics_status",        # 基础指标爬取状态：ok / partial / missing
        "intl_status",           # 国际化爬取状态
        "employment_status",     # 就业爬取状态
        "asean_status",          # 东盟爬取状态
        "last_crawled_at",       # 最近一次任意指标爬取时间
        "last_intl_crawled_at",  # 最近一次国际化爬取时间
        "last_employment_crawled_at", # 最近一次就业爬取时间
        "last_asean_crawled_at", # 最近一次东盟爬取时间
        "manual_notes",          # 人工备注（例如特殊办学情况、数据不确定等）
    ]

    # 把所有需要新增的列合并去重
    all_new_cols = (
        basic_info_cols
        + scale_cols
        + faculty_cols
        + resource_cols
        + intl_cols
        + employment_cols
        + nlp_cols
        + tech_cols
    )

    for col in all_new_cols:
        if col not in df.columns:
            df[col] = None  # 统一先填空值，方便后续脚本直接使用

    # 建议学校名作为第一列，方便人工查看
    # （保留 step2 中已有列顺序，只把 school_name / year 放前面）
    cols = list(df.columns)
    # 确保 school_name 在第一列，year 紧随其后
    if "school_name" in cols:
        cols.remove("school_name")
        cols.insert(0, "school_name")
    if "year" in cols:
        cols.remove("year")
        cols.insert(1, "year")

    df = df[cols]

    # 5. 写出
    df.to_excel(STEP3_FILE, index=False)
    print(f"[DONE] 已生成 Step3 指标空表：{STEP3_FILE}")
    print(f"  - 共 {len(df)} 所学校")
    print(f"  - 新增指标列数量：{len(all_new_cols)}（含基础 / 国际化 / 就业 / 东盟 / NLP / 技术列）")


def run_step1() -> None:
    """Entry point for preparing the unified Step3 metrics table."""
    init_step3_table()


if __name__ == "__main__":
    run_step1()
