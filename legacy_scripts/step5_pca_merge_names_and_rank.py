import pandas as pd

# 配置区
ORIGINAL_FILE = "step4C_kimi_completed_metrics_2025.xlsx"
PCA_COMPONENTS_FILE = "step5_pca_components.xlsx"
OUTPUT_FILE = "step5_pca_with_names_and_ranks.xlsx"

def main():
    print(">>> Step5 - 合并学校名称 & 生成国际化维度排名 开始")

    # 1. 读取原始完整数据（含 school_name）
    df_all = pd.read_excel(ORIGINAL_FILE)
    print(f"原始表形状：{df_all.shape[0]} 行 × {df_all.shape[1]} 列")
    if "school_name" not in df_all.columns:
        raise ValueError("原始表中找不到列 'school_name'，请确认列名。")

    # 2. 读取 PCA 主成分得分
    pca_df = pd.read_excel(PCA_COMPONENTS_FILE)
    print(f"PCA 得分表形状：{pca_df.shape[0]} 行 × {pca_df.shape[1]} 列")

    # 安全检查：行数要一致
    if len(df_all) != len(pca_df):
        raise ValueError("原始表与 PCA 表行数不一致，请检查前面步骤。")

    # 3. 合并：按行顺序一一对应，加上学校名称、省份等字段
    merged = pca_df.copy()
    merged["school_name"] = df_all["school_name"]

    # 如果有省份/城市列，也一并带上（可选）
    for extra_col in ["province", "city", "location"]:
        if extra_col in df_all.columns and extra_col not in merged.columns:
            merged[extra_col] = df_all[extra_col]

    # 4. 选定哪个 PC 作为“国际化维度”
    # 根据前面载荷分析，PC4 由 intl_partner_universities_count / countries / studyabroad_students_annual 主导
    intl_pc = "PC4"
    if intl_pc not in merged.columns:
        raise ValueError(f"{intl_pc} 不在 PCA 得分表中，请检查列名。")

    merged["intl_score_raw"] = merged[intl_pc]

    # 5. 计算国际化维度排名：分数越高，名次越靠前
    merged["intl_rank"] = merged["intl_score_raw"].rank(ascending=False, method="min").astype(int)

    # 6. 也顺便算一个“规模维度”的排名（PC1），以后可以用来做对比（可选）
    if "PC1" in merged.columns:
        merged["scale_score_raw"] = merged["PC1"]
        merged["scale_rank"] = merged["scale_score_raw"].rank(ascending=False, method="min").astype(int)

    # 7. 按国际化得分从高到低排序，方便查看 Top 20
    merged_sorted = merged.sort_values(by="intl_score_raw", ascending=False)

    print("\n=== 国际化维度（PC4）得分 Top 20 学校 ===")
    cols_to_show = ["intl_rank", "school_name", "intl_score_raw"]
    extra_cols = [c for c in ["province", "city", "location"] if c in merged_sorted.columns]
    cols_to_show.extend(extra_cols)
    print(merged_sorted[cols_to_show].head(20).to_string(index=False))

    # 8. 查找“广西外国语”相关学校
    mask_gx = merged_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)
    gx_rows = merged_sorted[mask_gx]

    print("\n=== 含“广西外国语”的学校在国际化维度上的位置 ===")
    if gx_rows.empty:
        print("未找到包含“广西外国语”的 school_name，请检查原始表中学校名称的具体写法。")
    else:
        show_cols = ["intl_rank", "intl_score_raw"]
        if "scale_rank" in merged_sorted.columns:
            show_cols.append("scale_rank")
        show_cols.append("school_name")
        show_cols.extend(extra_cols)

        print(gx_rows[show_cols].to_string(index=False))

    # 9. 保存完整结果到 Excel
    merged_sorted.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存包含学校名称与排名的完整表到：{OUTPUT_FILE}")

    print("\n>>> Step5 - 合并 & 排名 完成")

if __name__ == "__main__":
    main()
