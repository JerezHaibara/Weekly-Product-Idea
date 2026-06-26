
import streamlit as st
from pptx import Presentation
import fitz  # PyMuPDF
import tempfile
import re

# ✅ PDF生成
from reportlab.platypus import SimpleDocTemplate, Image as RLImage, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

# =========================================================
# ✅ 页面标题
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("PPT分类 + PDF原页面展示 + 自动生成报告")

# =========================================================
# ✅ 使用说明
# =========================================================
st.info("""
📌 使用说明：
请上传【同一份文件】的两个版本：

1️⃣ PPTX（用于分类）  
2️⃣ PDF（用于展示）

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
# ✅ 生成最终 PDF 报告
# =========================================================
def generate_report_pdf(grouped, image_list):

    output_path = "/tmp/final_report.pdf"

    doc = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()
    content = []

    content.append(Paragraph("Weekly Product Report", styles["Title"]))
    content.append(Spacer(1, 20))

    ordered_main = ["Yield", "Option", "Others", "Unknown Information"]

    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        content.append(Paragraph(main_category, styles["Heading1"]))
        content.append(Spacer(1, 10))

        # ✅ Unknown 单层
        if main_category == "Unknown Information":

            slides_list = grouped[main_category]["Unknown"]

            for slide in slides_list:
                page_num = slide["page"]

                content.append(Paragraph(f"Page {page_num}", styles["Heading3"]))

                img_path = image_list[page_num - 1]
                content.append(RLImage(img_path, width=500, height=350))
                content.append(Spacer(1, 20))

            continue

        # ✅ 普通分类
        for sub_category, slides_list in grouped[main_category].items():

            content.append(Paragraph(sub_category, styles["Heading2"]))
            content.append(Spacer(1, 10))

            for slide in slides_list:
                page_num = slide["page"]

                content.append(Paragraph(f"Page {page_num}", styles["Heading3"]))

                img_path = image_list[page_num - 1]
                content.append(RLImage(img_path, width=500, height=350))
                content.append(Spacer(1, 20))

    doc.build(content)

    return output_path

# =========================================================
# ✅ 主逻辑
# =========================================================
if ppt_file and pdf_file:

    # ✅ PPT解析
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

    # ✅ PDF拆页
    image_list = pdf_to_images(pdf_file)

    # ✅ 页数校验
    if len(slides_data) != len(image_list):

        st.error(f"""
❌ 页数不一致！

PPT 页数：{len(slides_data)}  
PDF 页数：{len(image_list)}

👉 请使用同一文件导出
""")

        st.stop()

    # ✅ 分类
    grouped = {}

    for slide in slides_data:

        cleaned = clean_text(slide["text"])
        main, sub = classify(cleaned)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # ✅ 展示
    ordered_main = ["Yield", "Option", "Others", "Unknown Information"]

    for main_category in ordered_main:

        if main_category not in grouped:
            continue

        if main_category == "Unknown Information":

            unknown_list = grouped[main_category]["Unknown"]

            with st.expander(f"📂 Unknown Information ({len(unknown_list)})"):

                for slide in unknown_list:
                    page_num = slide["page"]
                    st.markdown(f"**📄 Page {page_num}**")
                    _ = st.image(image_list[page_num - 1], use_container_width=True)

            continue

        st.subheader(f"📂 {main_category}")

        for sub_category, slides_list in grouped[main_category].items():

            with st.expander(f"📁 {sub_category} ({len(slides_list)})"):

                for slide in slides_list:
                    page_num = slide["page"]
                    st.markdown(f"**📄 Page {page_num}**")
                    _ = st.image(image_list[page_num - 1], use_container_width=True)

    # ✅ 生成 PDF 报告
    report_path = generate_report_pdf(grouped, image_list)

    st.success("✅ 分类报告 PDF 已生成")

    with open(report_path, "rb") as f:
        st.download_button(
            label="📥 下载完整分类报告 PDF",
            data=f,
            file_name="weekly_product_report.pdf"
        )
``
