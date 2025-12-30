import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
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
    主处理逻辑
    """
    # 1. 全局设置
    set_global_document_settings(doc)
    
    # 2. 定位结构
    body_start, ref_start = locate_structural_indices(doc, config['has_title_page'])
    
    paragraphs = doc.paragraphs
    
    # ==========================
    # 阶段 I: 处理正文 (Body)
    # ==========================
    for i in range(body_start, ref_start):
        p = paragraphs[i]
        text = p.text.strip()
        
        # 跳过空行，不处理（避免产生带缩进的空行垃圾）
        if not text:
            continue
            
        # 应用基础字体和行距
        apply_basic_font_style(p)
        
        pf = p.paragraph_format
        
        # --- 标题与缩进逻辑 ---
        
        # Case 1: 文章主标题 (Body 的第一段)
        if i == body_start and config['has_article_title']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            pf.first_line_indent = Inches(0) # 标题不缩进
            # 加粗
            for run in p.runs:
                run.bold = True
                
        # Case 2: 潜在的二级标题 (Level 2 Heading)
        # 判据：字数少于15 且 结尾无标点 且 不是主标题
        elif len(text.split()) < 15 and text[-1] not in ['.', ':', '?', '!']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            pf.first_line_indent = Inches(0) # 标题不缩进
            pf.left_indent = Inches(0)
            # 加粗
            for run in p.runs:
                run.bold = True
                
        # Case 3: 普通正文段落
        else:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            # APA 7th 首行缩进 0.5 英寸
            pf.first_line_indent = Inches(0.5)
            # 确保没有奇怪的悬挂缩进
            pf.left_indent = Inches(0) 

    # ==========================
    # 阶段 II: 处理参考文献 (Refs)
    # ==========================
    if ref_start < len(paragraphs):
        # 1. 处理 "References" 标题
        ref_title_p = paragraphs[ref_start]
        ref_title_p.text = "References" # 强制修正单复数
        apply_basic_font_style(ref_title_p)
        ref_title_p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        ref_title_p.paragraph_format.first_line_indent = Inches(0)
        for run in ref_title_p.runs:
            run.bold = True
            
        # 2. 获取参考文献条目列表
        ref_entries = []
        # 收集 ref_start 之后的所有非空段落
        entries_indices = [] # 记录索引方便后续删除
        
        for i in range(ref_start + 1, len(paragraphs)):
            p = paragraphs[i]
            if p.text.strip():
                ref_entries.append(p.text.strip())
                entries_indices.append(p)

        # 3. 排序逻辑 (如果启用)
        if config['sort_references']:
            # 警告：这会丢失斜体
            ref_entries.sort()
            
            # 删除旧段落 (反向删除以保持索引稳定，虽然 python-docx 删除段落比较hacky)
            # 这里的简单做法是：清空原段落内容，填入新内容。
            # 如果数量不一致（比如删了空行），则清空后在末尾追加。
            
            # 为了简单稳健：我们只保留标题，清除后面所有段落，然后重新添加
            # 注意：python-docx 删除段落需要操作 XML，这里用一个更安全的方法：
            # 将排序后的文本回写。如果原位置不够，就 add_paragraph。
            
            # --- 简易回写策略 ---
            current_idx = ref_start + 1
            # 覆盖现有的
            for text_content in ref_entries:
                if current_idx < len(paragraphs):
                    p = paragraphs[current_idx]
                    p.text = text_content
                    apply_basic_font_style(p)
                    # 悬挂缩进
                    p.paragraph_format.first_line_indent = Inches(-0.5)
                    p.paragraph_format.left_indent = Inches(0.5)
                    current_idx += 1
                else:
                    # 新增
                    new_p = doc.add_paragraph(text_content)
                    apply_basic_font_style(new_p)
                    new_p.paragraph_format.first_line_indent = Inches(-0.5)
                    new_p.paragraph_format.left_indent = Inches(0.5)
            
            # 如果原文档段落比新条目多（比如原文档有很多空行），清空剩余的
            while current_idx < len(paragraphs):
                paragraphs[current_idx].text = ""
                paragraphs[current_idx].clear() # 尽力清除
                current_idx += 1
                
        else:
            # 不排序，仅格式化 (保留斜体)
            for i in range(ref_start + 1, len(paragraphs)):
                p = paragraphs[i]
                if not p.text.strip(): continue
                
                apply_basic_font_style(p)
                # 悬挂缩进逻辑: Left Indent 0.5, First Line -0.5
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
    hide_streamlit_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                
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
                }
                </style>
                """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

    # --- 标题区 ---
    st.title("📄 APA 7th Format Assistant")
    st.markdown("Designed specifically for **Dr. Jin**'s academic workflow.")
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
            
            # --- 引用检查报告 ---
            if check_citations_opt:
                missing = check_missing_citations(doc)
                if missing:
                    st.warning("🧐 **Citation Check Report:**")
                    st.write("The following in-text citations might be missing from the Reference list:")
                    for m in missing:
                        st.markdown(f"- `{m}`")
                    st.caption("*Note: This is an automated check. Please verify manually.*")
                else:
                    st.info("👏 No obvious missing citations found.")

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
