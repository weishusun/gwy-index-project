import pandas as pd
import re
import numpy as np

INPUT_FILE = "step4C_kimi_completed_metrics_2025.xlsx"
OUTPUT_FILE = "step5_intl_features_A1_language_and_majors.xlsx"

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


def main():
    print(">>> Step5-A1 - 语种数量 & 外语类专业数量 特征构建 开始")

    # 1. 读取原始表
    df = pd.read_excel(INPUT_FILE)
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
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存带有新特征的表到：{OUTPUT_FILE}")
    print(">>> Step5-A1 - 完成")

if __name__ == "__main__":
    main()
