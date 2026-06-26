
import streamlit as st
from pptx import Presentation
import fitz  # PyMuPDF
import tempfile
import re

# =========================================================
# ✅ 页面标题
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("PPT分类 + PDF原页面展示（完全自动）")

# =========================================================
# ✅ 上传（必须两个）
# =========================================================
ppt_file = st.file_uploader("Upload PPTX (for classification)", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF (for display)", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# ✅ PDF → 图片
# =========================================================
def pdf_to_images(pdf_file):

    images = []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    pdf = fitz.open(tmp_path)

    for i in range(len(pdf)):
        page = pdf[i]
        pix = page.get_pixmap()
        img_path = f"/tmp/page_{i+1}.png"
        pix.save(img_path)
        images.append(img_path)

    return images

# =========================================================
# ✅ 清洗文本
# =========================================================
def clean_text(text):
    text = text.lower()
    text = re.sub(r'\s+', ' ', text)
    return text

# =========================================================
# ✅ 分类逻辑（你的核心）
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
# ✅ 主逻辑
# =========================================================
if ppt_file and pdf_file:

    # -------------------------------------------
    # ✅ 1️⃣ 解析 PPT（获取分类）
    # -------------------------------------------
    prs = Presentation(ppt_file)

    for i, slide in enumerate(prs.slides):

        text_content = ""

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_content += shape.text + " "

        slides_data.append({
            "page": i + 1,
            "text": text_content
        })

    # -------------------------------------------
    # ✅ 2️⃣ PDF 转图片
    # -------------------------------------------
    image_list = pdf_to_images(pdf_file)

    # ✅ 页数检查
    if len(slides_data) != len(image_list):
        st.warning(f"⚠️ 页数不一致：PPT({len(slides_data)}) vs PDF({len(image_list)})")

    # -------------------------------------------
    # ✅ 3️⃣ 分类构建（✅已修复你之前的bug）
    # -------------------------------------------
    grouped = {}

    for slide in slides_data:

        cleaned = clean_text(slide["text"])
        main, sub = classify(cleaned)

        # ✅ 一级分类
        if main not in grouped:
            grouped[main] = {}

        # ✅ 二级分类（关键修复点）
        if sub not in grouped[main]:
            grouped[main][sub] = []

        # ✅ 添加 slide
        grouped[main][sub].append(slide)

    # -------------------------------------------
    # ✅ 4️⃣ 展示（核心：page 对齐）
    # -------------------------------------------
    ordered_main = ["Yield", "Option", "Others"]

    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        st.subheader(f"📂 {main_category}")

        for sub_category, slides_list in grouped[main_category].items():

            with st.expander(f"📁 {sub_category} ({len(slides_list)})"):

                for slide in slides_list:

                    page_num = slide["page"]

                    st.markdown(f"**📄 Page {page_num}**")

                    # ✅ 关键对齐逻辑
                    if page_num - 1 < len(image_list):
                        _ = st.image(image_list[page_num - 1], use_container_width=True)
                    else:
                        st.error(f"❌ 缺少 PDF 第 {page_num} 页")

