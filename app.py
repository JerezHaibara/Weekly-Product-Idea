
import streamlit as st
from pptx import Presentation
import fitz  # PyMuPDF
import tempfile
import re
from datetime import datetime

# =========================================================
# 页面标题 + 日期
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("PPT分类 + PDF页面展示（V1）")

report_date = st.date_input("📅 选择报告日期", value=datetime.today())

# =========================================================
# 使用说明
# =========================================================
st.info("""
📌 使用说明：
请上传【同一份文件】的两个版本：

1️⃣ PPTX（用于分类）  
2️⃣ PDF（用于页面展示）

⚠️ 页数必须完全一致
""")

# =========================================================
# 上传文件
# =========================================================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# PDF转图片
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
# 文本清洗
# =========================================================
def clean_text(text):
    text = text.lower()
    text = re.sub(r'\s+', ' ', text)
    return text

# =========================================================
# 分类逻辑
# =========================================================
def classify(text):

    if "fcn" in text:
        return "Yield", "FCN"
    elif "sharkfin" in text:
        return "Option", "Sharkfin"
    elif "aq" in text:
        return "Option", "AQ"
    elif "dq" in text:
        return "Option", "DQ"
    elif "twinwin" in text:
        return "Option", "Twinwin"
    elif "dual digital" in text or "warrant" in text:
        return "Option", "Options"
    elif "range accrual" in text or "dci" in text:
        return "Yield", "Range Accrual / DCI"
    elif "fund" in text:
        return "Others", "Fund"
    elif "ben" in text:
        return "Others", "BEN"
    elif "tarf" in text:
        return "Others", "TARF"
    elif "inverse floater" in text:
        return "Others", "Inverse Floater"
    elif "stable note" in text:
        return "Others", "Stable Note"
    else:
        return "Unknown Information", "Unknown"

# =========================================================
# 主逻辑
# =========================================================
if ppt_file and pdf_file:

    # ✅ 解析PPT
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

    # ✅ PDF转图片
    image_list = pdf_to_images(pdf_file)

    # ✅ 页数校验
    if len(slides_data) != len(image_list):

        st.error(f"""
❌ 页数不一致！

PPT 页数：{len(slides_data)}  
PDF 页数：{len(image_list)}
""")
        st.stop()

    # ✅ 分类构建（✅ 已修复bug）
    grouped = {}

    for slide in slides_data:

        cleaned = clean_text(slide["text"])
        main, sub = classify(cleaned)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # ✅ 展示顺序
    ordered_main = ["Yield", "Option", "Others", "Unknown Information"]

    # =====================================================
    # ✅ 展示结果
    # =====================================================
    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        # ✅ Unknown Information（单层折叠）
        if main_category == "Unknown Information":

            unknown_list = grouped[main_category]["Unknown"]

            with st.expander(f"📂 Unknown Information ({len(unknown_list)})"):

                for slide in unknown_list:
                    page_num = slide["page"]

                    st.markdown(f"**📄 Page {page_num}**")
                    st.image(image_list[page_num - 1], use_container_width=True)

            continue

        # ✅ 正常分类
        st.subheader(f"📂 {main_category}")

        for sub_category, slides_list in grouped[main_category].items():

            with st.expander(f"📁 {sub_category} ({len(slides_list)})"):

                for slide in slides_list:

                    page_num = slide["page"]

                    st.markdown(f"**📄 Page {page_num}**")
                    st.image(image_list[page_num - 1], use_container_width=True)
