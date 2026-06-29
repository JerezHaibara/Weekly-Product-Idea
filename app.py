
import streamlit as st
from pptx import Presentation
import fitz
import tempfile
import re
from datetime import datetime

# =========================================================
# 页面标题
# =========================================================
st.title("📊 Investment Product Explorer V2")
st.caption("Enhanced Classification + Professional Structure")

report_date = st.date_input("📅 报告日期", value=datetime.today())

# =========================================================
# 使用说明
# =========================================================
st.info("""
📌 使用说明：
请上传【同一份文件】的：

1️⃣ PPTX（用于分类）  
2️⃣ PDF（用于展示）

⚠️ 页数必须完全一致
""")

# =========================================================
# 上传
# =========================================================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# PDF → 图片
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
# 清洗文本
# =========================================================
def clean_text(text):
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text

# =========================================================
# ✅ V2 分类逻辑（核心升级）
# =========================================================
def classify(text):

    # ===== 主类 =====
    if "fcn" in text:
        return "FCN", "FCN"

    elif "range accrual" in text:
        return "Range Accrual", "Range Accrual"

    elif "dci" in text:
        return "DCI", "DCI"

    elif "sharkfin" in text:
        return "Sharkfin", "Sharkfin"

    elif "aq" in text:
        return "AQ", "Accumulator"

    elif "fund" in text or "bluebay" in text or "singularity" in text:
        return "Fund", "Fund"

    # ===== Others =====
    elif "twinwin" in text:
        return "Others", "Twinwin"

    elif "dual digital" in text or "warrant" in text:
        return "Others", "Dual Digital"

    elif "ben" in text:
        return "Others", "BEN"

    elif "dq" in text:
        return "Others", "DQ"

    elif "tarf" in text:
        return "Others", "TARF"

    elif "inverse floater" in text:
        return "Others", "Inverse Floater"

    elif "stable note" in text:
        return "Others", "Stable Note"

    else:
        return "Others", "Unknown"

# =========================================================
# 主逻辑
# =========================================================
if ppt_file and pdf_file:

    # ✅ 读取 PPT
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

    # ✅ PDF 转图
    image_list = pdf_to_images(pdf_file)

    # ✅ 页数校验
    if len(slides_data) != len(image_list):
        st.error(f"""
❌ 页数不一致！

PPT：{len(slides_data)} 页  
PDF：{len(image_list)} 页
""")
        st.stop()

    # =====================================================
    # 分类构建（稳定版 ✅）
    # =====================================================
    grouped = {}

    for slide in slides_data:

        cleaned = clean_text(slide["text"])
        main, sub = classify(cleaned)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # =====================================================
    # 展示结构（V2优化 ✅）
    # =====================================================
    ordered_main = [
        "FCN",
        "Range Accrual",
        "DCI",
        "Sharkfin",
        "AQ",
        "Fund",
        "Others"
    ]

    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        st.subheader(f"📂 {main_category}")

        for sub_category, slides_list in grouped[main_category].items():

            with st.expander(f"{sub_category} ({len(slides_list)})"):

                for slide in slides_list:

                    page_num = slide["page"]

                    st.markdown(f"**Page {page_num}**")
                    st.image(image_list[page_num - 1], use_container_width=True)
