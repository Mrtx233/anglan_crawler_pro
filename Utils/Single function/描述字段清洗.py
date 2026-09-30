import pandas as pd
import re
import shutil
import os

# ====================== 配置 修改这里文件路径 ======================
INPUT_CSV = r"C:\Users\Mrtx0\Desktop\ba.csv"   #替换为你的csv路径
# =================================================================

# 1.匹配 Bxx / Cxx / BCxx 两位数字编号
pattern_code = re.compile(r"\b(?:B|C|BC)\d{2}\b")
# 2.匹配以Celebrity Favorites such as开头的完整p标签块，整段删除
pattern_celebrity = re.compile(r"<p>Celebrity Favorites such as.*?</p>", re.DOTALL)


def clean_desc(text):
    if pd.isna(text):
        return text
    s = str(text)
    # 删除名人推荐的完整p段落
    s = pattern_celebrity.sub("", s)
    # 删除 B/C/BC+两位数字编号
    s = pattern_code.sub("", s)
    # 清理连续空白、换行、多余空格
    s = re.sub(r"\s+", " ", s).strip()
    return s


def main():
    # 备份源文件
    bak_path = INPUT_CSV + ".bak"
    if os.path.exists(INPUT_CSV):
        shutil.copy2(INPUT_CSV, bak_path)
        print(f"✅已备份源文件到：{bak_path}")
    else:
        print(f"❌文件不存在 {INPUT_CSV}")
        return

    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)

    if "Description" not in df.columns:
        print("❌csv中找不到 Description 列！")
        return

    df["Description"] = df["Description"].apply(clean_desc)

    df.to_csv(INPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"✅处理完成，已写回：{INPUT_CSV}")


if __name__ == "__main__":
    main()
