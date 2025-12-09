import pandas as pd

INPUT_FILE = "step5_A9_asean_features.xlsx"
OUTPUT_FILE = "step5_A10_tli_features.xlsx"

# === 使用你数据中真实存在的文本列 ===
TEXT_COLS = [
    "profile_text_snippet",
    "intl_text_snippet",
    "asean_text_snippet",
    "positioning_keywords",
    "intl_keywords",
    "asean_keywords",
]

INTL_KEYWORDS = [
    "国际化", "国际", "全球", "全球化", "国际视野",
    "外国语", "外语", "多语种", "多语言", "跨文化", "跨国",
    "国际合作", "海外交流", "境外交流", "访学", "交换生",
    "留学生", "国际学生",
    "东盟", "东南亚", "RCEP", "一带一路",
]

def minmax(s: pd.Series) -> pd.Series:
    s = s.astype(float)
    if s.max() == s.min():
        return s * 0
    return (s - s.min()) / (s.max() - s.min())

def combine_text(row):
    texts = []
    for col in TEXT_COLS:
        if col in row and isinstance(row[col], str):
            texts.append(row[col])
    return "\n".join(texts)

def intl_score(text: str) -> float:
    if not isinstance(text, str):
        return 0.0
    total = 0
    hits = 0
    for kw in INTL_KEYWORDS:
        c = text.count(kw)
        total += c
        if c > 0:
            hits += 1
    return total + 0.5 * hits

def main():
    print(">>> Step5-A10 – 文本国际化指数 TLI 计算开始")

    df = pd.read_excel(INPUT_FILE)

    df["text_all"] = df.apply(combine_text, axis=1)
    df["raw_tli_score"] = df["text_all"].apply(intl_score)
    df["TLI"] = minmax(df["raw_tli_score"])

    df_sorted = df.sort_values("TLI", ascending=False)

    print("\n=== TLI Top 20 ===")
    print(df_sorted[["school_name", "raw_tli_score", "TLI"]].head(20).to_string(index=False))

    gx = df[df["school_name"].astype(str).str.contains("广西外国语")]
    print("\n=== 广西外国语学院的 TLI 情况 ===")
    print(gx[["school_name", "raw_tli_score", "TLI"]].to_string(index=False))

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存：{OUTPUT_FILE}")
    print(">>> Step5-A10 完成")

if __name__ == "__main__":
    main()
