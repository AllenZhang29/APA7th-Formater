import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.text import WD_BREAK
import re
import io

# ==============================================================================
# 1. 核心逻辑与算法模块 (Backend Logic)
# ==============================================================================

def set_global_document_settings(doc):
    """
    全局设置：页边距 (1英寸)
    注意：APA 7th 要求所有页边距均为 1 英寸 (2.54cm)
    """
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

def apply_basic_font_style(paragraph):
    """
    基础样式应用：Times New Roman, 12pt, 双倍行距
    """
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE # 双倍行距
    
    # 有些文档可能混杂了复杂的 Style，这里强制覆盖 Run 级别的字体
    for run in paragraph.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        
    # 为了保险，也尝试设置 Style 级别（如果有 Normal 样式）
    try:
        style = paragraph.style
        if style and hasattr(style, 'font'):
            style.font.name = 'Times New Roman'
            style.font.size = Pt(12)
    except:
        pass

def locate_structural_indices(doc, has_title_page):
    """
    智能定位算法：
    1. 寻找 body_start_index (正文开始的段落索引)
    2. 寻找 ref_start_index (参考文献开始的段落索引)
    """
    paragraphs = doc.paragraphs
    total_pars = len(paragraphs)
    
    body_start_index = 0
    ref_start_index = total_pars # 默认为末尾，即没找到

    # --- A. 定位参考文献 (Reference) ---
    # 策略：倒序或正序查找 "Reference" 独占一行的段落
    # 优先找 References，这样可以确定正文的边界
    for i, p in enumerate(paragraphs):
        # 清洗文本：去空格，转小写
        text = p.text.strip().lower()
        # 匹配 "reference" 或 "references"，且字数不能太多（防止匹配到正文里的句子）
        if text in ['reference', 'references'] or text == 'reference list':
            ref_start_index = i
            break
    
    # --- B. 定位正文起始 (Body Start) ---
    # 这是一个高难度动作，涉及“安全视窗”和“非空穿透”逻辑
    
    if has_title_page:
        found_page_break = False
        SAFE_SEARCH_LIMIT = 50 # 安全视窗：只在前50段寻找标题页逻辑
        search_limit = min(SAFE_SEARCH_LIMIT, ref_start_index)
        
        # 策略 1: 寻找物理分页符 (Hard Page Break)
        for i in range(search_limit):
            # 深入 XML 检查是否有 <w:br w:type="page"/>
            if '<w:br w:type="page"/>' in p._element.xml:
                body_start_index = i + 1 # 分页符所在段落的下一段是正文
                found_page_break = True
                break
        
        # 策略 2: 软换行穿透 (Rule of 6)
        if not found_page_break:
            non_empty_count = 0
            target_index = 0
            
            # 1. 计数：找到第6个有文字的段落 (通常是 Date)
            for i in range(search_limit):
                if paragraphs[i].text.strip():
                    non_empty_count += 1
                if non_empty_count == 6:
                    target_index = i
                    break
            
            # 2. 穿透：从第6个非空段落往后，跳过所有空行，直到遇到文字
            for j in range(target_index + 1, search_limit):
                if paragraphs[j].text.strip():
                    body_start_index = j
                    break
                    
    return body_start_index, ref_start_index

def process_formatting(doc, config):
    """
    主处理逻辑 (V2 Updated)
    """
    # 1. 全局设置
    set_global_document_settings(doc)
    
    # 2. 定位结构
    body_start, ref_start = locate_structural_indices(doc, config['has_title_page'])
    
    paragraphs = doc.paragraphs
    
    # ==========================
    # 阶段 0: 标题页特殊处理 (Title Page Formatting)
    # ==========================
    if config['has_title_page'] and body_start > 0:
        # 需求：标题页的前6行居中，第一行加粗
        title_lines_count = 0
        for i in range(body_start):
            p = paragraphs[i]
            if p.text.strip(): # 只处理有字的行
                title_lines_count += 1
                if title_lines_count <= 6:
                    p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    # 第一行加粗 (文章主标题)
                    if title_lines_count == 1:
                        for run in p.runs:
                            run.bold = True
                else:
                    # 超过6行的其他内容（如日期后的附加信息），暂维持居中或默认
                    p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # 需求：软换行变硬分页
        # 检查正文前一段 (body_start - 1)，如果不是分页符，强制插入一个分页符
        if body_start > 0:
            prev_p = paragraphs[body_start - 1]
            # 这是一个简单的判断：直接在上一段末尾加分页符
            # 这样无论原本是空行还是文字，都会强制换页，并且原来的空行虽然保留但会在上一页
            # 为了更干净，可以尝试清空中间的空段落，但比较复杂，直接加 Break 最稳妥
            
            # 避免重复：如果上一段已经是 Page Break (XML check)，就不加了
            if '<w:br w:type="page"/>' not in prev_p._element.xml:
                # 在上一段的最后一个 run 后面加 break，或者新加一个 run
                prev_p.add_run().add_break(WD_BREAK.PAGE)

    # ==========================
    # 阶段 I: 处理正文 (Body)
    # ==========================
    for i in range(body_start, ref_start):
        p = paragraphs[i]
        text = p.text.strip()
        
        # 跳过空行
        if not text:
            continue
            
        apply_basic_font_style(p)
        pf = p.paragraph_format
        
        # --- 标题与缩进逻辑 ---
        
        # Case 1: 文章主标题 (Body 的第一段)
        if i == body_start and config['has_article_title']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            pf.first_line_indent = Inches(0)
            for run in p.runs:
                run.bold = True
                
        # Case 2: 潜在的二级标题 (Level 2 Heading)
        elif len(text.split()) < 15 and text[-1] not in ['.', ':', '?', '!']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            pf.first_line_indent = Inches(0)
            pf.left_indent = Inches(0)
            for run in p.runs:
                run.bold = True
                
        # Case 3: 普通正文段落
        else:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            pf.first_line_indent = Inches(0.5)
            pf.left_indent = Inches(0) 

    # ==========================
    # 阶段 II: 处理参考文献 (Refs)
    # ==========================
    if ref_start < len(paragraphs):
        # 需求：References 前强制分页
        # 检查 ref_start 的前一段
        if ref_start > 0:
            prev_p_ref = paragraphs[ref_start - 1]
            if '<w:br w:type="page"/>' not in prev_p_ref._element.xml:
                 prev_p_ref.add_run().add_break(WD_BREAK.PAGE)

        # 1. 处理 "References" 标题
        ref_title_p = paragraphs[ref_start]
        ref_title_p.text = "References" 
        apply_basic_font_style(ref_title_p)
        ref_title_p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        ref_title_p.paragraph_format.first_line_indent = Inches(0)
        for run in ref_title_p.runs:
            run.bold = True
            
        # 2. 获取并处理条目 (保持原有逻辑，此处省略重复代码，请保留你原文件中 Sort/Hanging 的部分)
        # ... (这里保留你原来 app.py 中处理 Reference list 和 Sorting 的代码) ...
        # (为方便你复制，下面我把 Reference List 的处理逻辑简写补全，确保你直接覆盖不出错)
        
        ref_entries = []
        entries_indices = []
        for i in range(ref_start + 1, len(paragraphs)):
            p = paragraphs[i]
            if p.text.strip():
                ref_entries.append(p.text.strip())
                entries_indices.append(p)

        if config['sort_references']:
            ref_entries.sort()
            current_idx = ref_start + 1
            for text_content in ref_entries:
                if current_idx < len(paragraphs):
                    p = paragraphs[current_idx]
                    p.text = text_content
                    apply_basic_font_style(p)
                    p.paragraph_format.first_line_indent = Inches(-0.5)
                    p.paragraph_format.left_indent = Inches(0.5)
                    current_idx += 1
                else:
                    new_p = doc.add_paragraph(text_content)
                    apply_basic_font_style(new_p)
                    new_p.paragraph_format.first_line_indent = Inches(-0.5)
                    new_p.paragraph_format.left_indent = Inches(0.5)
            
            while current_idx < len(paragraphs):
                paragraphs[current_idx].text = ""
                paragraphs[current_idx].clear()
                current_idx += 1     
        else:
            for i in range(ref_start + 1, len(paragraphs)):
                p = paragraphs[i]
                if not p.text.strip(): continue
                apply_basic_font_style(p)
                p.paragraph_format.left_indent = Inches(0.5)
                p.paragraph_format.first_line_indent = Inches(-0.5)

    return doc

def check_missing_citations(doc):
    """
    引用查漏报告 (只读逻辑)
    """
    text_full = "\n".join([p.text for p in doc.paragraphs])
    
    # 1. 提取参考文献列表的首作者 (假设 Ref 标题后都是条目)
    # 简易逻辑：找 "References" 后的段落
    refs_authors = []
    found_ref = False
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.lower() == 'references':
            found_ref = True
            continue
        if found_ref and txt:
            # 提取第一个单词作为姓氏 (比如 "Zhang, S. (2020)") -> "Zhang"
            first_word = txt.split(',')[0].split(' ')[0]
            if len(first_word) > 1: # 排除杂讯
                refs_authors.append(first_word)

    # 2. 提取正文引用
    # 正则策略：匹配 (Name, Year) 或 (Name & Name, Year)
    # 这是一个非常粗略的正则，用于 MVP
    potential_citations = re.findall(r'\(([^)]+?,\s?\d{4})\)', text_full)
    
    missing_report = []
    
    # 3. 对比：正文引用的名字是否出现在 Reference 作者列表中
    if found_ref:
        for cite in potential_citations:
            # cite 可能是 "Wang & Li, 2020"
            # 只要 cite 包含 refs_authors 中的任何一个，就算匹配成功
            is_found = False
            for auth in refs_authors:
                if auth in cite:
                    is_found = True
                    break
            
            if not is_found:
                # 再次过滤：有时候引用里包含 'see Table 1' 这种误报
                if not re.search(r'Table|Figure|See|e\.g\.', cite, re.IGNORECASE):
                     missing_report.append(cite)
    
    return list(set(missing_report)) # 去重

# ==============================================================================
# 2. 前端交互模块 (Frontend UI)
# ==============================================================================

def main():
    st.set_page_config(page_title="APA 7th Format Helper", page_icon="🎓")

# --- CSS 注入：美化 & 隐藏水印 & 底部签名 ---
    # 修复：移除了 header 的隐藏，找回侧边栏按钮
    hide_streamlit_style = """
                <style>
                #MainMenu {visibility: visible;} 
                footer {visibility: hidden;}
                
                /* 自定义底部签名 */
                .custom-footer {
                    position: fixed;
                    left: 0;
                    bottom: 0;
                    width: 100%;
                    background-color: #f0f2f6;
                    color: #555;
                    text-align: center;
                    padding: 10px;
                    font-size: 14px;
                    font-family: 'Arial', sans-serif;
                    border-top: 1px solid #e6e6e6;
                    z-index: 999;
                }
                </style>
                """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

    # --- 标题区 ---
    st.title("📄 APA 7th Format Assistant")
    # st.markdown("Designed specifically for **Dr. Jin**'s academic workflow.")
    st.markdown("---")

    # --- 侧边栏配置 ---
    st.sidebar.header("⚙️ Configuration")
    
    has_title_page = st.sidebar.checkbox(
        "Has Title Page? (Skip Page 1)", 
        value=False,
        help="勾选后，工具将智能跳过封面页（识别分页符或前6行内容），从第二页开始格式化。"
    )
    
    has_article_title = st.sidebar.checkbox(
        "Has Article Title?", 
        value=True,
        help="勾选后，正文的第一段将被格式化为居中加粗的主标题。"
    )
    
    sort_references = st.sidebar.checkbox(
        "Auto-sort References (A-Z)", 
        value=False,
    )
    
    # 动态警告
    if sort_references:
        st.sidebar.warning(
            "⚠️ Warning: Auto-sorting will verify strict alphabetical order but "
            "**MAY REMOVE ITALICS** (e.g., journal names). Uncheck if you want to keep existing italics."
        )

    check_citations_opt = st.sidebar.checkbox(
        "Check Missing Citations", 
        value=True,
        help="生成一份报告，检查正文中引用的文献是否在 Reference 列表中缺失。"
    )

    # --- 文件上传区 ---
    uploaded_file = st.file_uploader("Drop your dissertation/paper here (.docx)", type="docx")

    if uploaded_file is not None:
        try:
            # 读取文件
            doc = Document(uploaded_file)
            
            # --- 运行处理逻辑 ---
            processed_doc = process_formatting(doc, {
                'has_title_page': has_title_page,
                'has_article_title': has_article_title,
                'sort_references': sort_references
            })
            
            st.success("✅ Formatting complete! Ready for download.")
            
            # --- 引用检查报告 (V2 Updated) ---
            if check_citations_opt:
                missing = check_missing_citations(doc)
                
                # 构建报告文本字符串
                report_content = ""
                
                # 根据是否排序，添加头部提示
                if sort_references:
                    report_content += "⚠️ [ACTION REQUIRED] References have been auto-sorted. ITALICS ARE REMOVED. Please re-apply italics to journal/book titles manually.\n\n"
                else:
                    report_content += "ℹ️ [INFO] References order kept as original. Please ensure they are alphabetical.\n\n"
                
                if missing:
                    report_content += "🧐 Potential Missing Citations (In-text vs Reference List):\n"
                    for m in missing:
                        report_content += f"- {m}\n"
                else:
                    report_content += "✅ No obvious missing citations found.\n"
                
                report_content += "\n*Report generated by APA 7th Format Assistant*"

                # UI 展示
                st.warning("🧐 **Citation Check Report:**")
                
                # 使用 st.code 展示报告，这样会自动带有一个 "Copy" 按钮
                st.code(report_content, language="markdown")
                
                st.caption("*Click the copy button in the top-right corner of the box above to send this report.*")

            # --- 导出 ---
            bio = io.BytesIO()
            processed_doc.save(bio)
            
            # 构建新文件名
            original_name = uploaded_file.name.rsplit('.', 1)[0]
            new_name = f"{original_name}_APA_Formatted.docx"
            
            st.download_button(
                label="📥 Download Formatted Document",
                data=bio.getvalue(),
                file_name=new_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error("Oops! Something went wrong processing the file.")
            st.error(f"Error details: {e}")

    # --- 底部签名 (Inject Footer) ---
    st.markdown(
        """
        <div class="custom-footer">
            Designed specially for Dr. Jin
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
