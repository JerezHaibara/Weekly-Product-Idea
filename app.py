
import streamlit as st
from pptx import Presentation
import fitz
import tempfile
import re
from datetime import datetime

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Image as RLImage, PageBreak, Table
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

PAGE_WIDTH, PAGE_HEIGHT = A4

# =========================================================
# UI
# =========================================================
st.title("📊 Investment Product Explorer V3")
st.caption("Final Stable Version + Professional PDF")

report_date = st.date_input("📅 报告日期", value=datetime.today())

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
        path = f"/tmp/page_{i+1}.png"
        pix.save(path)
        images.append(path)

    return images

# =========================================================
# 清洗
# =========================================================
def clean_text(text):
    return re.sub(r"\s+", " ", text.lower())

# =========================================================
# 分类（V3稳定版）
# =========================================================
def classify(text):

    if "dci" in text:
        if "dual" in text and "range accrual" in text:
            return "DCI", "Dual Range DCI"
        elif "range accrual" in text:
            return "DCI", "Range DCI"
        else:
            return "DCI", "Vanilla DCI"

    elif "range accrual" in text:
        if "dual" in text:
            return "Accrual Note", "Dual Accrual"
        else:
            return "Accrual Note", "Accrual"

    elif "fcn" in text:
        return "FCN", "FCN"

    elif "sharkfin" in text:
        return "Sharkfin", "Sharkfin"

    elif "aq" in text:
        return "AQ", "AQ"

    elif "fund" in text:
        return "Fund", "Fund"

    elif "twinwin" in text:
        return "Others", "Twinwin"

    elif "digital" in text:
        return "Others", "Digital"

    elif "ben" in text:
        return "Others", "BEN"

    elif "dq" in text:
        return "Others", "DQ"

    elif "tarf" in text:
        return "Others", "TARF"

    elif "inverse" in text:
        return "Others", "Inverse Floater"

    else:
        return "Others", "Unknown"

# =========================================================
# ✅ PDF（专业版）
# =========================================================
def generate_pdf(grouped, image_list):

    output = "/tmp/final_report.pdf"
    doc = SimpleDocTemplate(output, pagesize=A4)
    styles = getSampleStyleSheet()

    content = []

    ordered_main = [
        "FCN",
        "Accrual Note",
        "DCI",
        "Sharkfin",
        "AQ",
        "Fund",
        "Others"
    ]

    # 👉 页眉（分类名）
    def draw_header(canvas, doc):
        canvas.setFont("Helvetica", 10)
        if hasattr(doc, "cat"):
            canvas.drawRightString(
                PAGE_WIDTH - 40,
                PAGE_HEIGHT - 30,
                doc.cat
            )

    # 👉 封面
    content.append(Paragraph("Investment Product Report", styles["Title"]))
    content.append(Spacer(1, 40))
    content.append(Paragraph(str(report_date), styles["Normal"]))
    content.append(PageBreak())

    # =====================================================
    # 主体
    # =====================================================
    for main in ordered_main:

        if main not in grouped:
            continue

        # 👉 当前分类
        slides = []

        if main == "Others":
            for sub in grouped[main].values():
                slides.extend(sub)
        else:
            for sub in grouped[main].values():
                slides.extend(sub)

        # ✅ 每类单独开始
        first_page = True

        for i in range(0, len(slides), 6):

            batch = slides[i:i+6]

            # ✅ 设置 header 分类
            doc.cat = main

            table_data = []
            row = []

            for slide in batch:

                page_num = slide["page"]
                img_path = image_list[page_num - 1]

                # ✅ 等比缩放
                img = RLImage(img_path)
                img._restrictSize(230, 160)

                cell = [
                    Paragraph(f"<b>Page {page_num}</b>", styles["Normal"]),
                    img
                ]

                row.append(cell)

                if len(row) == 2:
                    table_data.append(row)
                    row = []

            if row:
                table_data.append(row)

            table = Table(table_data, colWidths=[260, 260])
            content.append(table)

            # ✅ 分页（但最后一个不加）
            if i + 6 < len(slides):
                content.append(PageBreak())

        # ✅ 分类结束后强制分页（下一类）
        content.append(PageBreak())

    doc.build(content, onFirstPage=draw_header, onLaterPages=draw_header)

    return output

# =========================================================
# 主流程
# =========================================================
if ppt_file and pdf_file:

    prs = Presentation(ppt_file)

    for i, slide in enumerate(prs.slides):
        text = ""
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + " "

        slides_data.append({
            "page": i + 1,
            "text": text
        })

    image_list = pdf_to_images(pdf_file)

    if len(slides_data) != len(image_list):
        st.error("❌ 页数不一致")
        st.stop()

    # ✅ 👇 最关键修复（分类恢复）
    grouped = {}

    for slide in slides_data:

        text = clean_text(slide["text"])
        main, sub = classify(text)

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # =====================================================
    # UI展示
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

    for main in ordered_main:

        if main not in grouped:
            continue

        count = sum(len(v) for v in grouped[main].values())

        with st.expander(f"📂 {main} ({count})"):

            if main == "Others":

                for sub, slides in grouped[main].items():

                    with st.expander(f"{sub} ({len(slides)})"):

                        for s in slides:
                            st.markdown(f"**Page {s['page']}**")
                            st.image(image_list[s["page"] - 1])

            else:
                flat = []
                for v in grouped[main].values():
                    flat.extend(v)

                for s in flat:
                    st.markdown(f"**Page {s['page']}**")
                    st.image(image_list[s["page"] - 1])

    # =====================================================
    # PDF下载
    # =====================================================
    pdf_path = generate_pdf(grouped, image_list)

    with open(pdf_path, "rb") as f:
        st.download_button("📄 下载PDF", f, "report.pdf")
``
