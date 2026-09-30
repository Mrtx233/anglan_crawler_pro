# -*- coding: utf-8 -*-
from pathlib import Path
import pandas as pd

# ========== 只需要改这里：换成你的真实 CSV 路径 ==========
SRC = Path(r"D:\C_code\anglan_crawler_pro\DrissionPage_Json_Shopify\02OUTPUT_XLSX\A0916\陈港\sylvelin.com女装\csv\wp-sylvelin.com女装_合并\wp-sylvelin.com女装_合并_Shopify_价格匹配_价格修正.csv")
SAMPLE_N = 10
# ======================================================

def main():
    if not SRC.exists():
        raise FileNotFoundError(f"文件不存在：{SRC}")

    # 读取：兼容带 BOM 的 UTF-8
    df = pd.read_csv(
        SRC,
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig"
    )

    required = ["Type", "SKU", "Parent"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"CSV 缺少必要列：{missing}，实际列：{list(df.columns)}")

    type_lower = df["Type"].fillna("").astype(str).str.strip().str.lower()
    sku_clean = df["SKU"].fillna("").astype(str).str.strip()
    parent_clean = df["Parent"].fillna("").astype(str).str.strip()

    # 所有主体商品 SKU
    parent_skus = (
        sku_clean[type_lower.eq("variable")]
        .drop_duplicates()
        .tolist()
    )

    if not parent_skus:
        raise ValueError("没有找到 Type=variable 的主体商品")

    n = min(SAMPLE_N, len(parent_skus))
    if n < SAMPLE_N:
        print(f"警告：主体商品只有 {len(parent_skus)} 个，将全部导出。")

    # 随机抽 10 个主体商品
    picked_parents = pd.Series(parent_skus).sample(n=n).tolist()

    # 保留：抽中的主体商品 + 对应变体
    mask = sku_clean.isin(picked_parents) | parent_clean.isin(picked_parents)
    sample_df = df.loc[mask].copy()

    # 输出：原文件名 + _抽测数据.csv
    out_path = SRC.with_name(SRC.stem + "_抽测数据.csv")

    # 保存为纯 UTF-8
    sample_df.to_csv(out_path, index=False, encoding="utf-8")

    print("抽中的主体商品：")
    for sku in picked_parents:
        name_rows = df.loc[sku_clean.eq(sku), "Name"]
        name = name_rows.iloc[0] if len(name_rows) else ""
        print(f"  - {sku}  {name}")

    print(f"\n已导出：{out_path}")
    print(f"编码：UTF-8（无 BOM）")
    print(f"总行数：{len(sample_df)}（主体 {len(picked_parents)} 个 + 对应变体）")


if __name__ == "__main__":
    main()