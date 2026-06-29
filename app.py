
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

st.markdown("""
### ⚠️ IMPORTANT 使用说明

- PPTX 和 PDF **必须来自同一份文件**
- PPTX 用于：**文本识别 & 分类**
- PDF 用于：**图像展示 & 报告生成**
""")

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
# 分类（不改 ✅）
# =========================================================
def classify(raw_text):

    text = clean_text(raw_text)

    if len(text.strip()) < 15:
        return "Unclassified", "Empty"

    if "dci" in text:
        return "DCI", None

    elif "range accrual" in text:
        return "Accrual Note", None

    elif "fcn" in text:
        return "FCN", None

    elif "sharkfin" in text:
        return "Sharkfin", None

    elif "aq" in text:
        return "AQ", None

    elif "fund" in text:
        return "Fund", None

    elif "twinwin" in text:
        return "Others", "Twinwin"

    elif "dq" in text:
        return "Others", "DQ"

    elif "ben" in text:
        return "Others", "BEN"

    else:
        return "Others", "Unknown"


# =========================================================
# PDF（不改 ✅）
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

    ordered_main = sorted(grouped.keys(), key=lambda x: priority_map.get(x, 999))

    # TOC
    story.append(Paragraph("Table of Contents", styles["Title"]))
    story.append(Spacer(1, 20))

    for main in ordered_main:
        total = sum(len(v) for v in grouped[main].values())
        story.append(Paragraph(f"{main} ({total})", styles["Normal"]))

    story.append(PageBreak())

    # Content
    for main in ordered_main:

        sub_dict = grouped[main]

        if main in ["FCN", "Accrual Note", "DCI", "Sharkfin", "AQ", "Fund"]:
            slides = sub_dict["all"]
        else:
            slides = []
            for v in sub_dict.values():
                slides.extend(v)

        count = len(slides)

        story.append(Spacer(1, 200))
        story.append(Paragraph(f"<b>{main} ({count})</b>", styles["Title"]))

        if main == "Others":

            story.append(Spacer(1, 20))

            sorted_sub = sorted(
                sub_dict.keys(),
                key=lambda x: (x in ["Unknown", "Empty"], x)
            )

            for sub in sorted_sub:
                story.append(
                    Paragraph(f"{sub} ({len(sub_dict[sub])})", styles["Normal"])
                )

        story.append(PageBreak())

        for i in range(0, len(slides), 6):

            batch = slides[i:i+6]
            table_data = []
            row = []

            for slide in batch:

                page_num = slide["page"]

                img = RLImage(image_list[page_num - 1])
                img._restrictSize(260, 180)

                row.append([
                    Paragraph(f"<font size=7>Page {page_num}</font>", styles["Normal"]),
                    img
                ])

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
# 主流程（仅修 bug ✅）
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

    grouped = {}

    for slide in slides_data:

        main, sub = classify(slide["text"])

        if main not in grouped:
            grouped[main] = {}

        # ✅ 修复点 1（all）
        if sub is None:
            if "all" not in grouped[main]:
                grouped[main]["all"] = []
            grouped[main]["all"].append(slide)

        # ✅ 修复点 2（sub）
        else:
            if sub not in grouped[main]:
                grouped[main][sub] = []
            grouped[main][sub].append(slide)

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

    ordered_main = sorted(grouped.keys(), key=lambda x: priority_map.get(x, 999))

    # UI
    for main in ordered_main:

        sub_dict = grouped[main]

        if main in ["FCN", "Accrual Note", "DCI", "Sharkfin", "AQ", "Fund"]:

            slides = sub_dict["all"]

            with st.expander(f"{main} ({len(slides)})"):

                for s in slides:
                    st.markdown(f"Page {s['page']}")
                    st.image(image_list[s["page"] - 1])

        else:

            total = sum(len(v) for v in sub_dict.values())

            with st.expander(f"{main} ({total})"):

                for sub, slides in sub_dict.items():

                    with st.expander(f"{sub} ({len(slides)})"):

                        for s in slides:
                            st.markdown(f"Page {s['page']}")
                            st.image(image_list[s["page"] - 1])

    # PDF
    pdf_path = generate_pdf(grouped, image_list)

    with open(pdf_path, "rb") as f:
        st.download_button(
            "📄 下载PDF",
            f,
            file_name=f"Weekly Product Idea {report_date}.pdf"
        )

