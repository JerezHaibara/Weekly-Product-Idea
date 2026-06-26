
import streamlit as st
from pptx import Presentation
import re

# =========================================================
# ✅ 页面标题（产品化）
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("自动解析结构化产品 · 按类别浏览")

# =========================================================
# ✅ STEP 1：上传 PPT（不展示解析细节）
# =========================================================
uploaded_file = st.file_uploader("Upload PPT", type=["pptx"])

slides_data = []

if uploaded_file:

    prs = Presentation(uploaded_file)

    for i, slide in enumerate(prs.slides):

        text_content = ""

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_content += shape.text + " "

        slides_data.append({
            "page": i + 1,
            "text": text_content
        })

# =========================================================
# ✅ 文本清洗（关键）
# =========================================================
def clean_text(text):
    text = text.lower()

    # 去控制字符
    text = re.sub(r'[\x00-\x1F\x7F]', ' ', text)

    # 替换换行
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

    # 全角转半角
    text = text.replace("（", "(").replace("）", ")")

    # 去多余空格
    text = re.sub(r'\s+', ' ', text)

    return text

# =========================================================
# ✅ 分类逻辑（保持你当前版本）
# =========================================================
def classify(text):
    t = text

    if "fcn" in t:
        return "Yield", "FCN"

    elif "sharkfin" in t:
        return "Option", "Sharkfin"

    elif "aq" in t:
        return "Option", "AQ"

    elif "dq" in t:
        return "Option", "DQ"

    elif "twinwin" in t:
        return "Option", "Twinwin"

    elif "dual digital" in t or "warrant" in t:
        return "Option", "Options"

    elif "range accrual" in t or "dci" in t:
        return "Yield", "Range Accrual / DCI"

    elif "fund" in t:
        return "Others", "Fund"

    # ✅ 扩展产品（你已有但还未启用）
    elif "ben" in t:
        return "Others", "BEN"

    elif "tarf" in t:
        return "Others", "TARF"

    elif "inverse floater" in t:
        return "Others", "Inverse Floater"

    elif "stable note" in t:
        return "Others", "Stable Note"

    else:
        return "Others", "Unknown Product"

# =========================================================
# ✅ STEP 2：产品库 UI（唯一展示）
# =========================================================
if slides_data:

    # ===== 构建分组 =====
    grouped = {}

    for slide in slides_data:

        text = slide["text"]
        page = slide["page"]

        if text.strip() == "":
            continue

        cleaned = clean_text(text)
        main, sub = classify(cleaned)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # ===== ✅ 关键：排序 Others 到最后 =====
    ordered_main = ["Yield", "Option", "Others"]

    # ===== ✅ 展示 =====
    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        st.subheader(f"📂 {main_category}")

        for sub_category in grouped[main_category]:

            slides_list = grouped[main_category][sub_category]

            with st.expander(f"📁 {sub_category} ({len(slides_list)})"):

                for slide in slides_list:

                    st.markdown(f"""
**📄 Page {slide['page']}**
👉 分类：{main_category} / {sub_category}
""")

                    with st.expander("📖 查看详情"):
                        st.write(slide["text"][:1200])
