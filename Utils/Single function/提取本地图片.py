import pandas as pd
import shutil

# ====================== 填写你的文件路径 ======================
INPUT_FILE = r"D:\C_code\anglan_crawler_pro\DrissionPage_Json_Shopify\02OUTPUT_XLSX\A0910\陈港\vilterras.com男装 减量\fitizenjeans.com.xlsx"
# ======================================================================

def extract_images_from_styles1(style_str):
    """解析styles1字段，提取@后面的图片链接，返回set集合"""
    if pd.isna(style_str) or not str(style_str).strip():
        return set()
    style_str = str(style_str)
    parts = style_str.split("#")
    img_set = set()
    for p in parts:
        if "@" in p:
            _, img_url = p.split("@", maxsplit=1)
            img_url = img_url.strip()
            if img_url:
                img_set.add(img_url)
    return img_set


def parse_src_links(src_str):
    """解析src_links，#分割返回图片list(保留原始顺序) 和 set集合"""
    if pd.isna(src_str) or not str(src_str).strip():
        return [], set()
    src_str = str(src_str)
    img_list = [u.strip() for u in src_str.split("#") if u.strip()]
    return img_list, set(img_list)


def main():
    # 1. 先创建 .bak 备份
    bak_file = INPUT_FILE + ".bak"
    shutil.copy2(INPUT_FILE, bak_file)
    print(f"✅ 已创建备份文件：{bak_file}")

    # 2. 读取表格
    df = pd.read_excel(INPUT_FILE)

    new_list = []
    for idx, row in df.iterrows():
        style1_val = row["styles1"]
        src_links_val = row["src_links"]

        set_styles1 = extract_images_from_styles1(style1_val)
        src_list, set_src = parse_src_links(src_links_val)

        both_set = set_styles1 & set_src          # 两边都有
        src_only_set = set_src - set_styles1      # src_links独有

        # 保留src原始顺序，筛选出src独有的列表
        src_only_list = [x for x in src_list if x in src_only_set]
        src_only_top4 = src_only_list[:4]         # 截取src独有前4张
        both_list = [x for x in src_list if x in both_set]

        # 组合新src_links：src独有前四张 + 两边都有的图片
        new_src_links_list = src_only_top4 + both_list
        new_src_links = "#".join(new_src_links_list)
        new_list.append(new_src_links)

        print(f"行号:{idx+1} 处理完成 -> {len(new_src_links_list)} 张图")

    # 3. 写回 src_links 列并保存到原文件
    df["src_links"] = new_list
    df.to_excel(INPUT_FILE, index=False)
    print(f"✅ 已写回原文件：{INPUT_FILE}")
    print(f"共处理 {len(df)} 行")


if __name__ == "__main__":
    main()
