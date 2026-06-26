
import streamlit as st
from pptx import Presentation

# =========================================================
# 🟩 STEP 0：页面标题
# =========================================================
st.title("Investment Product Hub")
st.write("PPT 自动解析与分类系统")

# =========================================================
# 🟦 STEP 1：上传 PPT + 读取每一页文本（原功能）
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

        # ✅ 存下来给后面用
        slides_data.append({
            "page": i + 1,
            "text": text_content
        })

# =========================================================
# 🟨 STEP 2：定义分类规则（新增）
# =========================================================
def classify(text):
    t = text.lower()

    # ===== Yield Products =====
    if "fcn" in t:
        return "Yield", "FCN"

    elif "range accrual" in t or "dci" in t:
        return "Yield", "Range Accrual"

    # ===== Option Strategies =====
    elif "sharkfin" in t:
        return "Option", "Sharkfin"

    elif "aq" in t:
        return "Option", "AQ"

    elif "dual digital" in t or "warrant" in t:
        return "Option", "Options"

    elif "twinwin" in t:
        return "Option", "Other Structures"

    # ===== Others =====
    else:
        return "Others", "Others"

# =========================================================
# 🟥 STEP 3：展示分类结果（新增）
# =========================================================
if slides_data:

    st.header("STEP 2：分类结果")

    for slide in slides_data:

        text_content = slide["text"]

        # 空页跳过
        if text_content.strip() == "":
            st.write(f"Page {slide['page']} → ⚠️ 空白页")
            continue

        main, sub = classify(text_content[:200])

        st.write(f"Page {slide['page']} → {main} / {sub}")
