import pandas as pd
import numpy as np

# ========= 配置区：如有需要可以改 =========
INPUT_FILE = "step4C_kimi_completed_metrics_2025.xlsx"
OUTPUT_FILE = "step5_pca_numeric_cleaned.xlsx"
MISSING_COL_THRESHOLD = 0.5  # 某一列缺失比例超过 50% 就丢弃
# ======================================


def main():
    print(">>> Step5 - PCA 数据预处理 开始")

    # 1. 读取原始总表
    print(f"读取文件：{INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
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
    numeric_df.to_excel(OUTPUT_FILE, index=False)
    print(f"已保存预处理后的数值数据到：{OUTPUT_FILE}")
    print(f"清洗后数据形状：{numeric_df.shape[0]} 行 × {numeric_df.shape[1]} 列")

    # 7. 打印前 10 列列名，方便你确认
    print("前 10 个数值列：")
    print(list(numeric_df.columns[:10]))

    print(">>> Step5 - PCA 数据预处理 完成")


if __name__ == "__main__":
    main()
