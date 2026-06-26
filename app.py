
import streamlit as st
from pptx import Presentation

# ===== 页面标题 =====
st.title("Investment Product Hub")

st.write("Step 1: 上传 PPT 并读取内容")

# ===== 上传文件 =====
uploaded_file = st.file_uploader("Upload PPT", type=["pptx"])

if uploaded_file:

    st.success("Upload successful ✅")

    # 读取 PPT
    prs = Presentation(uploaded_file)

    st.header("读取结果（每一页文本）")

    # 遍历每一页
    for i, slide in enumerate(prs.slides):

        text_content = ""

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_content += shape.text + " "

        st.subheader(f"Page {i+1}")

        if text_content.strip() == "":
            st.write("⚠️ 没有识别到文字")
        else:
            st.write(text_content[:300])  # 显示前300字符
