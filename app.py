
import streamlit as st
from pptx import Presentation
import re
import fitz  # PyMuPDF
import tempfile
import os

# =========================================================
# ✅ 页面标题
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("上传 PDF 自动解析并展示 PPT 页面")

# =========================================================
# ✅ 上传 PDF
# =========================================================
uploaded_pdf = st.file_uploader("Upload PPT (PDF format)", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# ✅ PDF → 图片（自动拆页）
# =========================================================
def pdf_to_images(pdf_file):

    images = []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    pdf = fitz.open(tmp_path)

    for page_num in range(len(pdf)):
        page = pdf[page_num]
        pix = page.get_pixmap()

        img_path = f"/tmp/page_{page_num+1}.png"
        pix.save(img_path)

        images.append(img_path)

    return images

# =========================================================
# ✅ 文本清洗
# =========================================================
def clean_text(text):
    text = text.lower()
    text = re.sub(r'[\x00-\x1F\x7F]', ' ', text)
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r'\s+', ' ', text)
    return text

# =========================================================
# ✅ 分类逻辑
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
# ✅ 处理 PDF
# =========================================================
if uploaded_pdf:

    # ✅ 1. 拆 PDF → 图片
    image_list = pdf_to_images(uploaded_pdf)

    # ✅ 2. 提取文本（用 pdf metadata）
    # 注意：PDF没有结构文本，简化为 page-level placeholder
    for i in range(len(image_list)):
        slides_data.append({
            "page": i + 1,
            "text": f"page_{i+1}"  # 用占位符（分类靠你原逻辑可扩展）
        })

    # =====================================================
    # ✅ 自动分类（这里是关键限制说明）
    # =====================================================
    grouped = {}

    for slide in slides_data:

        # ⚠️ 由于 PDF 无结构文本，这里建议未来：
        # 👉 可结合 OCR / 或继续使用 PPT 上传

        main = "Others"
        sub = "Slides"

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # =====================================================
    # ✅ 展示
    # =====================================================
    st.subheader("📂 All Slides")

    with st.expander(f"📁 Slides ({len(image_list)})"):

        for i, slide in enumerate(slides_data):

            st.markdown(f"**📄 Page {slide['page']}**")

            if i < len(image_list):
                st.image(image_list[i], use_container_width=True)
            else:
                st.warning("图片缺失")

