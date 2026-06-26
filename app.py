
import streamlit as st
from pptx import Presentation
import re

# =========================================================
# 🟩 STEP 0：页面标题
# =========================================================
st.title("Investment Product Hub")
st.write("PPT 自动解析 + 产品分类系统")

# =========================================================
# 🟦 STEP 1：上传 PPT + 读取文本
# =========================================================
uploaded_file = st.file_uploader("Upload PPT", type=["pptx"])

slides_data = []

if uploaded_file:

    st.success("Upload successful ✅")

    prs = Presentation(uploaded_file)

    st.header("STEP 1：原始读取结果（每一页）")

    for i, slide in enumerate(prs.slides):

        text_content = ""

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_content += shape.text + " "

        if text_content.strip() == "":
            st.write(f"Page {i+1} → ⚠️ 空白页")
        else:
            st.write(f"Page {i+1}")
            st.write(text_content[:300])

        # ✅ 存入数据池
        slides_data.append({
            "page": i + 1,
            "text": text_content
        })

# =========================================================
# 🟨 STEP 2A：文本清洗（关键升级）
# =========================================================
def clean_text(text):
    text = text.lower()

    # 去控制字符
    text = re.sub(r'[\x00-\x1F\x7F]', ' ', text)

    # 替换换行等
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

    # 全角转半角（常见）
    text = text.replace("（", "(").replace("）", ")")

    # 去多余空格
    text = re.sub(r'\s+', ' ', text)

    return text

# =========================================================
# 🟨 STEP 2B：分类逻辑（产品优先识别）
# =========================================================
def classify(text):
    t = text

    # ===== 产品优先识别 =====

    # FCN
    if "fcn" in t:
        return "Yield", "FCN"

    # Sharkfin
    elif "sharkfin" in t:
        return "Option", "Sharkfin"

    # AQ
    elif "aq" in t:
        return "Option", "AQ"

    # DQ
    elif "dq" in t:
        return "Option", "DQ"

    # Twinwin
    elif "twinwin" in t:
        return "Option", "Twinwin"

    # Dual Digital / Warrant
    elif "dual digital" in t or "warrant" in t:
        return "Option", "Options"

    # Range Accrual / DCI
    elif "range accrual" in t or "dci" in t:
        return "Yield", "Range Accrual / DCI"

    # Fund
    elif "fund" in t or "对冲基金" in t:
        return "Others", "Fund"

    # fallback（理论很少触发）
    else:
        return "Others", "Unknown Product"

# =========================================================
# 🟥 STEP 3：分类展示
# =========================================================
if slides_data:

    st.header("STEP 2：分类结果")

    for slide in slides_data:

        page = slide["page"]
        text = slide["text"]

        if text.strip() == "":
            st.write(f"Page {page} → ⚠️ 空白页")
            continue

        # ✅ 先清洗
        cleaned = clean_text(text)

        # ✅ 再分类（用完整文本，不截断）
        main, sub = classify(cleaned)

        st.write(f"Page {page} → {main} / {sub}")

