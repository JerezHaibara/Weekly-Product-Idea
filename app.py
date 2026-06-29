
import streamlit as st
from pptx import Presentation
import fitz
import tempfile
import re
from datetime import datetime

# =========================================================
# ✅ 页面标题
# =========================================================
st.title("📊 Investment Product Explorer V2")
st.caption("Clean Classification Structure")

report_date = st.date_input("📅 报告日期", value=datetime.today())

# =========================================================
# ✅ 使用说明
# =========================================================
st.info("""
📌 使用说明：
上传同一份文件的：

1️⃣ PPTX（分类）  
2️⃣ PDF（展示）

⚠️ 页数必须一致
""")

# =========================================================
# ✅ 上传文件
# =========================================================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

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
# ✅ 文本清洗
# =========================================================
def clean_text(text):
    return re.sub(r"\s+", " ", text.lower())

# =========================================================
# ✅ 分类逻辑（最终版）
# =========================================================
def classify(text):

    # ✅ DCI 优先（关键）
    if "dci" in text:

        if "dual" in text and "range accrual" in text:
            return "DCI", "Dual Range Accrual DCI"

        elif "range accrual" in text:
            return "DCI", "Range Accrual DCI"

        else:
            return "DCI", "Vanilla DCI"

    # ✅ Accrual Note（原 Range Accrual）
    elif "range accrual" in text:

        if "dual" in text:
            return "Accrual Note", "Dual Accrual Note"

        else:
            return "Accrual Note", "Accrual Note"

    # ✅ 主分类
    elif "fcn" in text:
        return "FCN", "FCN"

    elif "sharkfin" in text:
        return "Sharkfin", "Sharkfin"

    elif "aq" in text:
        return "AQ", "Accumulator"

    elif "fund" in text or "bluebay" in text or "singularity" in text:
        return "Fund", "Fund"

    # ✅ Others
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
# ✅ 主逻辑
# =========================================================
if ppt_file and pdf_file:

    prs = Presentation(ppt_file)

    # ✅ 读取 PPT
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

PPT：{len(slides_data)}  
PDF：{len(image_list)}
""")
        st.stop()

    # =====================================================
    # ✅ 分类构建（已修复所有bug）
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
    # ✅ 展示顺序
    # =====================================================
    ordered_main = [
        "FCN",
        "Accrual Note",
        "DCI",
        "Sharkfin",
        "AQ",
        "Fund",
        "Others"
    ]

    # =====================================================
    # ✅ UI 展示（最终版本 ✅）
    # =====================================================
    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        total_count = sum(len(v) for v in grouped[main_category].values())

        with st.expander(f"📂 {main_category} ({total_count})", expanded=False):

            # ✅ Others（唯一二级分类）
            if main_category == "Others":

                sorted_subcats = sorted(
                    grouped[main_category].keys(),
                    key=lambda x: (x == "Unknown", x)
                )

                for sub_category in sorted_subcats:

                    slides_list = grouped[main_category][sub_category]

                    with st.expander(f"{sub_category} ({len(slides_list)})"):

                        for slide in slides_list:

                            page_num = slide["page"]

                            st.markdown(f"**Page {page_num}**")
                            st.image(image_list[page_num - 1], use_container_width=True)

            # ✅ 其他分类（不分子类）
            else:

                slides_list = []

                for sub_list in grouped[main_category].values():
                    slides_list.extend(sub_list)

                for slide in slides_list:

                    page_num = slide["page"]

                    st.markdown(f"**Page {page_num}**")
                    st.image(image_list[page_num - 1], use_container_width=True)

