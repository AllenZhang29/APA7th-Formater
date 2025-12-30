import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io

# === 简单的密码保护 ===
def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        if st.session_state["password"] == "20251112": 
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "请输入启动密码 (Password):", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect, show input + error.
        st.text_input(
            "密码错误，请重试:", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if check_password():
    # === 这里是主程序 ===
    st.title("📄 APA 7th Format Helper")
    st.write("Designed specially for Dr. [Her Name]")

    uploaded_file = st.file_uploader("Upload your Word Document (.docx)", type="docx")

    if uploaded_file is not None:
        # 1. 读取文件
        doc = Document(uploaded_file)
        
        # 2. 处理逻辑 (这里只是示例，你需要完善你的 python-docx 逻辑)
        # 你的核心代码就在这里发挥作用
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        
        for paragraph in doc.paragraphs:
            paragraph_format = paragraph.paragraph_format
            paragraph_format.line_spacing = 2.0 # 双倍行距
            # 其他处理逻辑...

        # 3. 保存到内存流 (不存硬盘)
        bio = io.BytesIO()
        doc.save(bio)
        
        # 4. 提供下载按钮
        st.download_button(
            label="Download Formatted Doc",
            data=bio.getvalue(),
            file_name="formatted_paper.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )