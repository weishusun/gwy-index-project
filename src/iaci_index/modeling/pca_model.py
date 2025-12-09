import pandas as pd
import numpy as np
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

# ========= 配置区：如有需要可以改 =========
PREP_INPUT_FILE = "step4C_kimi_completed_metrics_2025.xlsx"
PREP_OUTPUT_FILE = "step5_pca_numeric_cleaned.xlsx"
MISSING_COL_THRESHOLD = 0.5  # 某一列缺失比例超过 50% 就丢弃

PCA_INPUT_FILE = "step5_pca_numeric_cleaned.xlsx"
PCA_COMPONENTS_FILE = "step5_pca_components.xlsx"
PCA_LOADINGS_FILE = "step5_pca_loadings.xlsx"

LOADINGS_TOP_N = 5

MERGE_ORIGINAL_FILE = "step4C_kimi_completed_metrics_2025.xlsx"
MERGED_OUTPUT_FILE = "step5_pca_with_names_and_ranks.xlsx"
# ======================================


def prepare_pca_data() -> None:
    print(">>> Step5 - PCA 数据预处理 开始")

    # 1. 读取原始总表
    print(f"读取文件：{PREP_INPUT_FILE}")
    df = pd.read_excel(PREP_INPUT_FILE)
    print(f"原始数据形状：{df.shape[0]} 行 × {df.shape[1]} 列")

    # 2. 只保留数值型列
    numeric_df = df.select_dtypes(include=["number"]).copy()
    print(f"数值型列数量：{numeric_df.shape[1]}")

    # 3. 删除不适合作为评价指标的基础列（例如 school_code / year）
    drop_cols = ["school_code", "year"]
    drop_cols = [c for c in drop_cols if c in numeric_df.columns]
    if drop_cols:
        print(f"删除基础列：{drop_cols}")
        numeric_df = numeric_df.drop(columns=drop_cols)

    # 4. 删除缺失超过阈值的列（例如缺失 > 50%）
    min_non_na = int(len(numeric_df) * (1 - MISSING_COL_THRESHOLD))
    before_cols = numeric_df.shape[1]
    numeric_df = numeric_df.dropna(axis=1, thresh=min_non_na)
    after_cols = numeric_df.shape[1]
    print(f"按缺失比例筛列：从 {before_cols} 列 -> {after_cols} 列")

    # 5. 用列中位数填补其余缺失值（每一列单独计算）
    medians = numeric_df.median(numeric_only=True)
    numeric_df = numeric_df.fillna(medians)

    # 6. 保存预处理后的数据
    numeric_df.to_excel(PREP_OUTPUT_FILE, index=False)
    print(f"已保存预处理后的数值数据到：{PREP_OUTPUT_FILE}")
    print(f"清洗后数据形状：{numeric_df.shape[0]} 行 × {numeric_df.shape[1]} 列")

    # 7. 打印前 10 列列名，方便你确认
    print("前 10 个数值列：")
    print(list(numeric_df.columns[:10]))

    print(">>> Step5 - PCA 数据预处理 完成")


def run_pca_components() -> None:
    print(">>> Step5 - PCA 主成分分析开始")

    df = pd.read_excel(PCA_INPUT_FILE)
    print(f"读取数据形状：{df.shape[0]} 行 × {df.shape[1]} 列")

    col_names = df.columns.tolist()

    scaler = StandardScaler()
    scaled = scaler.fit_transform(df)

    pca = PCA()
    pca.fit(scaled)

    explained = pca.explained_variance_ratio_
    print("\n前 10 个主成分的解释率（方差贡献）：")
    for i, var in enumerate(explained[:10]):
        print(f"PC{i+1}: {var:.4f}")

    components = pca.transform(scaled)
    comp_df = pd.DataFrame(
        components,
        columns=[f"PC{i+1}" for i in range(components.shape[1])]
    )
    comp_df.to_excel(PCA_COMPONENTS_FILE, index=False)
    print(f"\n主成分得分已保存：{PCA_COMPONENTS_FILE}")

    loadings = pd.DataFrame(
        pca.components_.T,
        index=col_names,
        columns=[f"PC{i+1}" for i in range(pca.components_.shape[0])]
    )
    loadings.to_excel(PCA_LOADINGS_FILE)
    print(f"指标载荷矩阵已保存：{PCA_LOADINGS_FILE}")

    print(">>> Step5 - PCA 主成分分析完成")


def interpret_pca_loadings() -> None:
    print(">>> PCA 指标载荷解释开始")
    loadings = pd.read_excel(PCA_LOADINGS_FILE, index_col=0)

    for pc in loadings.columns:
        print(f"\n====== {pc} Top {LOADINGS_TOP_N} 指标 ======")
        top_loadings = loadings[pc].abs().sort_values(ascending=False).head(LOADINGS_TOP_N)
        for idx in top_loadings.index:
            weight = loadings.loc[idx, pc]
            print(f"{idx:40s} 载荷 = {weight:.4f}")

    print("\n>>> PCA 指标载荷解释完成")


def merge_pca_with_names() -> None:
    print(">>> Step5 - 合并学校名称 & 生成国际化维度排名 开始")

    df_all = pd.read_excel(MERGE_ORIGINAL_FILE)
    print(f"原始表形状：{df_all.shape[0]} 行 × {df_all.shape[1]} 列")
    if "school_name" not in df_all.columns:
        raise ValueError("原始表中找不到列 'school_name'，请确认列名。")

    pca_df = pd.read_excel(PCA_COMPONENTS_FILE)
    print(f"PCA 得分表形状：{pca_df.shape[0]} 行 × {pca_df.shape[1]} 列")

    if len(df_all) != len(pca_df):
        raise ValueError("原始表与 PCA 表行数不一致，请检查前面步骤。")

    merged = pca_df.copy()
    merged["school_name"] = df_all["school_name"]

    for extra_col in ["province", "city", "location"]:
        if extra_col in df_all.columns and extra_col not in merged.columns:
            merged[extra_col] = df_all[extra_col]

    intl_pc = "PC4"
    if intl_pc not in merged.columns:
        raise ValueError(f"{intl_pc} 不在 PCA 得分表中，请检查列名。")

    merged["intl_score_raw"] = merged[intl_pc]
    merged["intl_rank_desc"] = merged["intl_score_raw"].rank(ascending=False, method="dense")
    merged_sorted = merged.sort_values(by="intl_score_raw", ascending=False)

    merged_sorted.to_excel(MERGED_OUTPUT_FILE, index=False)
    print(f"\n已保存合并后的排名表：{MERGED_OUTPUT_FILE}")
    print(">>> Step5 - 合并学校名称 & 排名 完成")


def run_pca_for_intl_index() -> None:
    prepare_pca_data()
    run_pca_components()
    interpret_pca_loadings()
    merge_pca_with_names()


if __name__ == "__main__":
    run_pca_for_intl_index()
