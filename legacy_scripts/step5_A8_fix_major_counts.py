import pandas as pd
import numpy as np

INPUT_FILE = "step5_intl_features_A2_kimi_languages.xlsx"
PCA_NAMES_FILE = "step5_pca_with_names_and_ranks.xlsx"
OUTPUT_FILE = "step5_A8_major_fixed_IACI.xlsx"

GXUFL_MAJOR_COUNT = 12  # 你提供的真实值


def minmax(series):
    s = series.astype(float)
    if s.max() == s.min():
        return s * 0
    return (s - s.min()) / (s.max() - s.min())


def main():
    print(">>> Step5-A8 - 修正外语类专业数量 + 重算 IACI（两维版）")

    df_feat = pd.read_excel(INPUT_FILE)
    df_pca = pd.read_excel(PCA_NAMES_FILE)

    # 仅保留需要的 PC4
    df_pca = df_pca[["school_name", "intl_score_raw"]]

    # 合并
    df = pd.merge(df_feat, df_pca, on="school_name", how="inner")

    # 语种数
    df["languages_offered_count"] = pd.to_numeric(df["languages_offered_count"], errors="coerce").fillna(0)

    # 外语专业数（修正）
    df["foreign_major_count_fixed"] = df["foreign_major_count_final"].copy()

    # NaN → 0 → 再单独修正 GXUFL
    df["foreign_major_count_fixed"] = pd.to_numeric(df["foreign_major_count_fixed"], errors="coerce").fillna(0)

    gx_mask = df["school_name"].str.contains("广西外国语", na=False)
    df.loc[gx_mask, "foreign_major_count_fixed"] = GXUFL_MAJOR_COUNT

    # ===== 重算 LRI =====
    df["lang_log"] = np.log1p(df["languages_offered_count"])
    df["lang_log_norm"] = minmax(df["lang_log"])

    df["major_log"] = np.log1p(df["foreign_major_count_fixed"])
    df["major_log_norm"] = minmax(df["major_log"])

    df["LRI"] = 0.6 * df["lang_log_norm"] + 0.4 * df["major_log_norm"]

    # ===== 重算国际合作指数 ICI =====
    df["ICI"] = minmax(pd.to_numeric(df["intl_score_raw"], errors="coerce").fillna(0))

    # ===== 两维综合指数 IACI =====
    df["IACI"] = 0.6 * df["LRI"] + 0.4 * df["ICI"]

    df["IACI_rank"] = df["IACI"].rank(ascending=False).astype(int)

    # 排序
    df_sorted = df.sort_values(by="IACI", ascending=False)

    print("\n=== 修正后的 IACI Top 15 ===")
    print(df_sorted[["IACI_rank", "school_name", "IACI", "LRI", "ICI",
                     "languages_offered_count", "foreign_major_count_fixed"]].head(15).to_string(index=False))

    print("\n=== 广西外国语学院（修正后） ===")
    print(df_sorted[gx_mask][["IACI_rank", "IACI", "LRI", "ICI",
                              "languages_offered_count", "foreign_major_count_fixed"]])

    df_sorted.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存：{OUTPUT_FILE}")
    print(">>> Step5-A8 完成")


if __name__ == "__main__":
    main()
