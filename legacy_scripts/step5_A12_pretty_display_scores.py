"""
Step5-A12：指数展示美化（保持排名不变，仅调整 0/1 观感）

输入：step5_A11_IACI_final_4D.xlsx
输出：step5_A12_IACI_final_4D_pretty.xlsx

- 保留原始：
    LRI_norm, ICI_norm, ARII_norm, TLI_norm, IACI_final_4D, IACI_rank
- 新增展示用列（仅用于报告展示）：
    LRI_norm_disp, ICI_norm_disp, ARII_norm_disp, TLI_norm_disp, IACI_final_4D_disp
"""

import pandas as pd

INPUT_FILE = "step5_A11_IACI_final_4D.xlsx"
OUTPUT_FILE = "step5_A12_IACI_final_4D_pretty.xlsx"


def pretty_scale(s: pd.Series) -> pd.Series:
    """
    将 0-1 区间线性映射到 0.05 - 0.95 区间：
        y = 0.05 + 0.90 * x
    保证单调递增，因此不会改变排序。
    """
    s = pd.to_numeric(s, errors="coerce").fillna(0).astype(float)
    return 0.05 + 0.90 * s


def main():
    print(">>> Step5-A12 - 指数展示美化 开始")

    df = pd.read_excel(INPUT_FILE)
    print(f"读取：{df.shape[0]} 行 × {df.shape[1]} 列")

    needed = ["LRI_norm", "ICI_norm", "ARII_norm", "TLI_norm", "IACI_final_4D", "IACI_rank", "school_name"]
    for col in needed:
        if col not in df.columns:
            raise ValueError(f"缺少必要列：{col}")

    # 为四个维度和综合指数生成展示用版本
    df["LRI_norm_disp"] = pretty_scale(df["LRI_norm"])
    df["ICI_norm_disp"] = pretty_scale(df["ICI_norm"])
    df["ARII_norm_disp"] = pretty_scale(df["ARII_norm"])
    df["TLI_norm_disp"]  = pretty_scale(df["TLI_norm"])
    df["IACI_final_4D_disp"] = pretty_scale(df["IACI_final_4D"])

    # 排名不变，这里只是检查一下（不重新算）
    df_sorted = df.sort_values("IACI_rank", ascending=True)

    print("\n=== IACI_final_4D（展示版）Top 20 学校 ===")
    print(
        df_sorted[
            ["IACI_rank", "school_name", "IACI_final_4D", "IACI_final_4D_disp",
             "LRI_norm_disp", "ICI_norm_disp", "ARII_norm_disp", "TLI_norm_disp"]
        ]
        .head(20)
        .to_string(index=False)
    )

    gx = df_sorted[df_sorted["school_name"].astype(str).str.contains("广西外国语", na=False)]
    print("\n=== 广西外国语学院（展示版） ===")
    if gx.empty:
        print("未找到包含“广西外国语”的记录。")
    else:
        print(
            gx[
                ["IACI_rank", "IACI_final_4D", "IACI_final_4D_disp",
                 "LRI_norm_disp", "ICI_norm_disp", "ARII_norm_disp", "TLI_norm_disp"]
            ].to_string(index=False)
        )

    df_sorted.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存美化后的指数到：{OUTPUT_FILE}")
    print(">>> Step5-A12 完成")


if __name__ == "__main__":
    main()
