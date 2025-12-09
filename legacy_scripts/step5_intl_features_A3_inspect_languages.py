import pandas as pd

INPUT_FILE = "step5_intl_features_A2_kimi_languages.xlsx"

def main():
    print(">>> Step5-A3 - 检查语种数量分布 & 异常值")

    df = pd.read_excel(INPUT_FILE)

    if "languages_offered_count" not in df.columns:
        raise ValueError("表中没有 languages_offered_count 列")

    # 1. 基本统计
    print("\n=== 语种数量（languages_offered_count）描述统计 ===")
    print(df["languages_offered_count"].describe())

    # 2. 按语种数量从高到低列出前 20 所学校
    top = df.sort_values(by="languages_offered_count", ascending=False)
    print("\n=== 语种数量 Top 20 学校 ===")
    print(top[["school_name", "languages_offered_count"]].head(20).to_string(index=False))

    # 3. 单独查看包含“广西外国语”的学校
    mask_gx = df["school_name"].astype(str).str.contains("广西外国语", na=False)
    gx = df[mask_gx]

    print("\n=== 广西外国语相关学校的语种情况 ===")
    if gx.empty:
        print("未找到包含“广西外国语”的 school_name。")
    else:
        print(gx[["school_name", "languages_offered_count", "languages_list", "foreign_major_count_final"]].to_string(index=False))

    # 4. 找出疑似异常的大值（例如 ≥ 20）
    outliers = df[df["languages_offered_count"] >= 20]
    print("\n=== 疑似异常的高语种数量学校（>=20） ===")
    if outliers.empty:
        print("没有 >=20 的语种数量记录。")
    else:
        print(outliers[["school_name", "languages_offered_count", "languages_list"]].to_string(index=False))

    print("\n>>> Step5-A3 - 完成")

if __name__ == "__main__":
    main()
