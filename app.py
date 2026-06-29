
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


# =========================================================
# UI
# =========================================================
st.title("📊 Investment Product Explorer V4 FINAL")

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
# 分类（增强版）
# =========================================================
def classify(raw_text):

    text = clean_text(raw_text)

    # ✅ Unclassified（空页 / 切割页）
    if len(text.strip()) < 15:
        return "Unclassified", "Empty"

    if "dci" in text:
        return "DCI", "DCI"

    elif "range accrual" in text:
        return "Accrual Note", "Accrual Note"

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

    elif "dq" in text:
        return "Others", "DQ"

    elif "ben" in text:
        return "Others", "BEN"

    else:
        return "Others", "Unknown"


# =========================================================
# ✅ PDF生成
# =========================================================
def generate_pdf(grouped, image_list):

    styles = getSampleStyleSheet()

    filename = f"/tmp/Weekly Product Idea {report_date}.pdf"

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=15,
        rightMargin=15,
        topMargin=30,
        bottomMargin=15
    )

    story = []

    # ✅ 优先级排序（关键✅）
    priority_map = {
        "FCN": 1,
        "Accrual Note": 2,
        "DCI": 3,
        "Sharkfin": 4,
        "AQ": 5,
        "Fund": 6,
        "Others": 99,
        "Unclassified": 100
    }

    ordered_main = sorted(
        grouped.keys(),
        key=lambda x: priority_map.get(x, 999)
    )

    # =========================
    # ✅ 目录
    # =========================
    story.append(Paragraph("Table of Contents", styles["Title"]))
    story.append(Spacer(1, 20))

    for main in ordered_main:

        total = sum(len(v) for v in grouped[main].values())
        story.append(Paragraph(f"{main} ({total})", styles["Normal"]))

    story.append(PageBreak())

    # =========================
    # ✅ 分类内容
    # =========================
    for main in ordered_main:

        sub_dict = grouped[main]
        slides = []

        for sub in sub_dict.values():
            slides.extend(sub)

        count = len(slides)

        # ✅ 分类标题页
        story.append(Spacer(1, 200))
        story.append(Paragraph(f"<b>{main} ({count})</b>", styles["Title"]))

        # ✅ Others子分类排序
        if main == "Others":

            story.append(Spacer(1, 20))

            sorted_sub = sorted(
                sub_dict.keys(),
                key=lambda x: (x in ["Unknown", "Empty"], x)
            )

            for sub in sorted_sub:
                items = sub_dict[sub]
                story.append(
                    Paragraph(f"{sub} ({len(items)})", styles["Normal"])
                )

        story.append(PageBreak())

        # ✅ 图片页（6 per page）
        for i in range(0, len(slides), 6):

            batch = slides[i:i+6]
            table_data = []
            row = []

            for slide in batch:

                page_num = slide["page"]

                img = RLImage(image_list[page_num - 1])
                img._restrictSize(260, 180)

                cell = [
                    Paragraph(f"<font size=7>Page {page_num}</font>", styles["Normal"]),
                    img
                ]

                row.append(cell)

                if len(row) == 2:
                    table_data.append(row)
                    row = []

            if row:
                table_data.append(row)

            story.append(Table(table_data, colWidths=[260, 260]))

            if i + 6 < len(slides):
                story.append(PageBreak())

        story.append(PageBreak())

    doc.build(story)
    return filename


# =========================================================
# 主逻辑
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

    # ✅ 构建分类
    grouped = {}

    for slide in slides_data:

        main, sub = classify(slide["text"])

        if main not in grouped:
            grouped[main] = {}

        if sub not in grouped[main]:
            grouped[main][sub] = []

        grouped[main][sub].append(slide)

    # ✅ 同样排序 UI
    priority_map = {
        "FCN": 1,
        "Accrual Note": 2,
        "DCI": 3,
        "Sharkfin": 4,
        "AQ": 5,
        "Fund": 6,
        "Others": 99,
        "Unclassified": 100
    }

    ordered_main = sorted(
        grouped.keys(),
        key=lambda x: priority_map.get(x, 999)
    )

    # =========================
    # ✅ UI
    # =========================
    for main in ordered_main:

        sub_dict = grouped[main]
        total = sum(len(v) for v in sub_dict.values())

        with st.expander(f"{main} ({total})"):

            sorted_sub = sorted(
                sub_dict.keys(),
                key=lambda x: (x in ["Unknown", "Empty"], x)
            )

            for sub in sorted_sub:

                slides = sub_dict[sub]

                with st.expander(f"{sub} ({len(slides)})"):

                    for s in slides:
                        st.markdown(f"Page {s['page']}")
                        st.image(image_list[s["page"] - 1])

    # =========================
    # ✅ PDF下载
    # =========================
    pdf_path = generate_pdf(grouped, image_list)

    with open(pdf_path, "rb") as f:
        st.download_button(
            "📄 下载PDF",
            f,
            file_name=f"Weekly Product Idea {report_date}.pdf"
        )
