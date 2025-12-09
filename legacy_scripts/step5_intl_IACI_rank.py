import pandas as pd
import numpy as np

# 输入文件
FEATURE_FILE = "step5_intl_features_A2_kimi_languages.xlsx"   # 有 languages_offered_count, foreign_major_count_final
PCA_WITH_NAMES_FILE = "step5_pca_with_names_and_ranks.xlsx"   # 有 school_name, intl_score_raw(PC4), 等
OUTPUT_FILE = "step5_IACI_internationalization_ranking.xlsx"

def minmax_normalize(series: pd.Series) -> pd.Series:
    """0-1 归一化，若常数列则返回 0"""
    s = series.astype(float)
    min_v = s.min()
    max_v = s.max()
    if pd.isna(min_v) or pd.isna(max_v) or max_v == min_v:
        return pd.Series(0.0, index=s.index)
    return (s - min_v) / (max_v - min_v)

def main():
    print(">>> Step5-A7 - 计算国际化办学能力综合指数 IACI")

    # 1. 读入特征表（语言类）
    feat_df = pd.read_excel(FEATURE_FILE)
    print(f"特征表形状：{feat_df.shape}")

    # 2. 读入 PCA + intl_score_raw（PC4）
    pca_df = pd.read_excel(PCA_WITH_NAMES_FILE)
    print(f"PCA+名称表形状：{pca_df.shape}")

    if "school_name" not in feat_df.columns or "school_name" not in pca_df.columns:
        raise ValueError("两个表都必须包含 school_name 列。")

    # 3. 按 school_name 合并（内连接，确保名称匹配）
    merged = pd.merge(
        feat_df,
        pca_df[["school_name", "intl_score_raw"]],  # 只取需要的列
        on="school_name",
        how="inner",
    )
    print(f"合并后形状：{merged.shape}")

    # 4. 处理语言资源相关字段
    # 4.1 语种数量 log 变换 + 标准化
    lang_raw = pd.to_numeric(merged["languages_offered_count"], errors="coerce").fillna(0)
    lang_log = np.log1p(lang_raw)  # log(1+x)
    merged["lang_log"] = lang_log
    merged["lang_log_norm"] = minmax_normalize(lang_log)

    # 4.2 外语类专业数 log 变换 + 标准化
    if "foreign_major_count_final" in merged.columns:
        foreign_raw = pd.to_numeric(merged["foreign_major_count_final"], errors="coerce").fillna(0)
    else:
        foreign_raw = pd.Series(0, index=merged.index)

    foreign_log = np.log1p(foreign_raw)
    merged["foreign_major_log"] = foreign_log
    merged["foreign_major_log_norm"] = minmax_normalize(foreign_log)

    # 4.3 语言资源子指数 LRI
    # 权重：语言数 0.6，外语专业数 0.4
    merged["LRI"] = (
        0.6 * merged["lang_log_norm"] +
        0.4 * merged["foreign_major_log_norm"]
    )

    # 5. 国际合作子指数 ICI（基于 intl_score_raw = PC4）
    intl_raw = pd.to_numeric(merged["intl_score_raw"], errors="coerce").fillna(0)
    merged["ICI"] = minmax_normalize(intl_raw)

    # 6. 综合国际化办学能力指数 IACI
    # 当前版本：IACI = 0.6 * LRI + 0.4 * ICI
    merged["IACI"] = 0.6 * merged["LRI"] + 0.4 * merged["ICI"]

    # 7. 计算排名（分数越高，名次越小）
    merged["IACI_rank"] = merged["IACI"].rank(ascending=False, method="min").astype(int)
    merged["LRI_rank"] = merged["LRI"].rank(ascending=False, method="min").astype(int)
    merged["ICI_rank"] = merged["ICI"].rank(ascending=False, method="min").astype(int)

    # 8. 按 IACI 从高到低排序
    merged_sorted = merged.sort_values(by="IACI", ascending=False)

    # 打印 Top 20
    print("\n=== 国际化办学能力综合指数 IACI Top 20 学校 ===")
    print(
        merged_sorted[
            ["IACI_rank", "school_name", "IACI", "LRI", "ICI", "LRI_rank", "ICI_rank", "languages_offered_count", "foreign_major_count_final", "intl_score_raw"]
        ].head(20).to_string(index=False)
    )

    # 单独看看广西外国语学院
    mask_gx = merged_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)
    gx = merged_sorted[mask_gx]
    print("\n=== 广西外国语学院在 IACI 体系中的位置 ===")
    if gx.empty:
        print("未找到包含“广西外国语”的 school_name。")
    else:
        print(
            gx[
                ["IACI_rank", "IACI", "LRI", "ICI", "LRI_rank", "ICI_rank", "languages_offered_count", "foreign_major_count_final", "intl_score_raw", "school_name"]
            ].to_string(index=False)
        )

    # 9. 保存完整结果
    merged_sorted.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存完整国际化指数与排名到：{OUTPUT_FILE}")
    print(">>> Step5-A7 - 完成")

if __name__ == "__main__":
    main()
