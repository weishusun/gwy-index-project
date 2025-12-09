import pandas as pd

LOADINGS_FILE = "step5_pca_loadings.xlsx"
TOP_N = 5   # 每个主成分查看前 5 个高载荷指标

def main():
    print(">>> PCA 指标载荷解释开始")

    # 1. 读取载荷矩阵
    loadings = pd.read_excel(LOADINGS_FILE, index_col=0)

    # 2. 遍历每一个主成分 PC1, PC2 ...
    for pc in loadings.columns:
        print(f"\n====== {pc} Top {TOP_N} 指标 ======")
        # 选出载荷绝对值最大的前 TOP_N 个指标
        top_loadings = loadings[pc].abs().sort_values(ascending=False).head(TOP_N)
        for idx in top_loadings.index:
            weight = loadings.loc[idx, pc]
            print(f"{idx:40s} 载荷 = {weight:.4f}")

    print("\n>>> PCA 指标载荷解释完成")

if __name__ == "__main__":
    main()
