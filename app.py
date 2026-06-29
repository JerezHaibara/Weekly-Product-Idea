
import streamlit as st
from pptx import Presentation
import fitz  # PyMuPDF
import tempfile
import re
from datetime import datetime

# PDF生成
from reportlab.platypus import SimpleDocTemplate, Image as RLImage, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

# =========================================================
# ✅ 页面标题 + 日期
# =========================================================
st.title("📊 Investment Product Explorer")
st.caption("PPT分类 + PDF展示 + 自动报告生成")

report_date = st.date_input("📅 选择报告日期", value=datetime.today())
formatted_date = report_date.strftime("%Y_%m_%d")

# =========================================================
# ✅ 使用说明
# =========================================================
st.info("""
📌 使用说明：
请上传【同一份文件】的两个版本：

1️⃣ PPTX（用于分类）  
2️⃣ PDF（用于页面展示）

⚠️ 页数必须完全一致
""")

# =========================================================
# ✅ 上传文件
# =========================================================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

slides_data = []
image_list = []

# =========================================================
# ✅ PDF转图片
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
# ✅ 生成最终PDF报告
# =========================================================
def generate_report_pdf(grouped, image_list, formatted_date):

    output_path = f"/tmp/weekly_product_report_{formatted_date}.pdf"

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

