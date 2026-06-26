
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
# ✅ 使用说明
# =========================================================
st.info("""
📌 使用说明：
请上传【同一份文件】的两个版本：

1️⃣ PPTX（用于分类）  
2️⃣ PDF（用于页面展示）

⚠️ 必须保证两者页数完全一致，否则无法匹配
""")

# =========================================================
# ✅ 上传文件
# =========================================================
ppt_file = st.file_uploader("Upload PPTX (for classification)", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF (for display)", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# ✅ PDF 转图片
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
    text = text.lower()
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
        return "Unknown Information", "Unknown"

# =========================================================
# ✅ 主逻辑
# =========================================================
if ppt_file and pdf_file:

    # -------------------------------------------
    # ✅ 1️⃣ 解析 PPT
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

    # -------------------------------------------
    # ✅ 页数校验
    # -------------------------------------------
    if len(slides_data) != len(image_list):

        st.error(f"""
❌ 页数不一致！

PPT 页数：{len(slides_data)}  
PDF 页数：{len(image_list)}

👉 请使用【同一文件】导出的 PDF
""")

        st.stop()

    # -------------------------------------------
    # ✅ 3️⃣ 分类构建
    # -------------------------------------------
    grouped = {}

    for slide in slides_data:

        cleaned = clean_text(slide["text"])
        main, sub = classify(cleaned)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # -------------------------------------------
    # ✅ 4️⃣ 展示
    # -------------------------------------------
    ordered_main = ["Yield", "Option", "Others", "Unknown Information"]

    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        # ✅ Unknown Information（✅一层折叠）
        if main_category == "Unknown Information":

            unknown_list = grouped[main_category]["Unknown"]

            with st.expander(f"📂 Unknown Information ({len(unknown_list)})"):

                for slide in unknown_list:

                    page_num = slide["page"]

                    st.markdown(f"**📄 Page {page_num}**")
                    _ = st.image(image_list[page_num - 1], use_container_width=True)

            continue

        # ✅ 普通分类
        st.subheader(f"📂 {main_category}")

        for sub_category, slides_list in grouped[main_category].items():

            with st.expander(f"📁 {sub_category} ({len(slides_list)})"):

                for slide in slides_list:

                    page_num = slide["page"]

                    st.markdown(f"**📄 Page {page_num}**")
                    _ = st.image(image_list[page_num - 1], use_container_width=True)
