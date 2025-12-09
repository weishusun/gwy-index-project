import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

INPUT_FILE = "step5_pca_numeric_cleaned.xlsx"
OUTPUT_COMPONENTS = "step5_pca_components.xlsx"
OUTPUT_LOADINGS = "step5_pca_loadings.xlsx"

def main():
    print(">>> Step5 - PCA 主成分分析开始")

    # 1. 读取清洗后的数据
    df = pd.read_excel(INPUT_FILE)
    print(f"读取数据形状：{df.shape[0]} 行 × {df.shape[1]} 列")

    col_names = df.columns.tolist()

    # 2. 标准化
    scaler = StandardScaler()
    scaled = scaler.fit_transform(df)

    # 3. 运行 PCA
    pca = PCA()
    pca.fit(scaled)

    explained = pca.explained_variance_ratio_
    print("\n前 10 个主成分的解释率（方差贡献）：")
    for i, var in enumerate(explained[:10]):
        print(f"PC{i+1}: {var:.4f}")

    # 4. 保存 PCA 转换后的主成分数据（每个学校的 PC1, PC2, ...）
    components = pca.transform(scaled)
    comp_df = pd.DataFrame(
        components,
        columns=[f"PC{i+1}" for i in range(components.shape[1])]
    )
    comp_df.to_excel(OUTPUT_COMPONENTS, index=False)
    print(f"\n主成分得分已保存：{OUTPUT_COMPONENTS}")

    # 5. 保存各指标在各主成分上的载荷（指标权重）
    loadings = pd.DataFrame(
        pca.components_.T,
        columns=[f"PC{i+1}" for i in range(pca.components_.shape[0])],
        index=col_names
    )
    loadings.to_excel(OUTPUT_LOADINGS)
    print(f"指标载荷矩阵已保存：{OUTPUT_LOADINGS}")

    print("\n>>> Step5 - PCA 主成分分析完成")

if __name__ == "__main__":
    main()
