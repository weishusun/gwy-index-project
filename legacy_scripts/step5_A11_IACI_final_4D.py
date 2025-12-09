"""
Step5-A11：国际化办学能力综合指数 IACI-4D（四维终版）

输入：step5_A10_tli_llm_features.xlsx
  - 已包含：school_name, LRI, ICI, ARII, TLI 等字段

输出：step5_A11_IACI_final_4D.xlsx
  - 新增：LRI_norm, ICI_norm, ARII_norm, TLI_norm, IACI_final_4D, IACI_rank
"""

import pandas as pd

# ===== 路径配置 =====
INPUT_FILE = "step5_A10_tli_llm_features.xlsx"
OUTPUT_FILE = "step5_A11_IACI_final_4D.xlsx"

# ===== 四个维度的权重（可以按需要微调）=====
W_LRI = 0.25   # 语言资源指数（语种 + 外语类专业）
W_ICI = 0.25   # 国际合作指数（PCA 主成分）
W_ARII = 0.25  # 东盟区域国际化指数
W_TLI = 0.25   # 文本国际化指数（LLM 补全）


def minmax(s: pd.Series) -> pd.Series:
    """0-1 归一化"""
    s = pd.to_numeric(s, errors="coerce").fillna(0).astype(float)
    min_v, max_v = s.min(), s.max()
    if max_v == min_v:
        return pd.Series(0.0, index=s.index)
    return (s - min_v) / (max_v - min_v)


def main():
    print(">>> Step5-A11-4D - 四维 IACI_final 计算开始")

    df = pd.read_excel(INPUT_FILE)
    print(f"读取：{df.shape[0]} 行 × {df.shape[1]} 列")

    # ===== 1. 检查必要列 =====
    needed_cols = ["school_name", "LRI", "ICI", "ARII", "TLI"]
    for col in needed_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要列：{col}，请检查 {INPUT_FILE} 是否来自前一步 A10-LLM")

    # ===== 2. 四个维度 0-1 归一化 =====
    df["LRI_norm"] = minmax(df["LRI"])
    df["ICI_norm"] = minmax(df["ICI"])
    df["ARII_norm"] = minmax(df["ARII"])
    df["TLI_norm"] = minmax(df["TLI"])

    # ===== 3. 计算四维综合指数 IACI_final_4D =====
    df["IACI_final_4D"] = (
        W_LRI * df["LRI_norm"]
        + W_ICI * df["ICI_norm"]
        + W_ARII * df["ARII_norm"]
        + W_TLI * df["TLI_norm"]
    )

    # 排名：数值越大，排名越靠前
    df["IACI_rank"] = df["IACI_final_4D"].rank(ascending=False, method="min").astype(int)

    # ===== 4. 按综合指数排序，打印 Top 20 =====
    df_sorted = df.sort_values(by="IACI_final_4D", ascending=False)

    print("\n=== IACI_final_4D Top 20 学校（四维国际化综合指数） ===")
    print(
        df_sorted[
            ["IACI_rank", "school_name", "IACI_final_4D",
             "LRI_norm", "ICI_norm", "ARII_norm", "TLI_norm"]
        ]
        .head(20)
        .to_string(index=False)
    )

    # ===== 5. 查看广西外国语学院的位置 =====
    gx_mask = df_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)
    gx = df_sorted[gx_mask]

    print("\n=== 广西外国语学院在 IACI-4D 体系中的位置 ===")
    if gx.empty:
        print("未找到包含“广西外国语”的记录，请检查 school_name 列。")
    else:
        print(
            gx[
                ["IACI_rank", "IACI_final_4D",
                 "LRI_norm", "ICI_norm", "ARII_norm", "TLI_norm"]
            ].to_string(index=False)
        )

    # ===== 6. 保存结果 =====
    df_sorted.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存最终综合指数与排名到：{OUTPUT_FILE}")
    print(">>> Step5-A11-4D 完成")


if __name__ == "__main__":
    main()
