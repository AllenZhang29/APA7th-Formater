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

def delete_paragraph(paragraph):
    """
    辅助函数：彻底删除一个段落对象
    """
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def process_formatting(doc, config):
    """
    主处理逻辑 (V3 Updated: Title Page Cleaning & Spacing)
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
        title_lines_count = 0
        last_title_paragraph = None
        
        # A. 格式化标题页的内容 (前 body_start 段)
        for i in range(body_start):
            p = paragraphs[i]
            
            # 如果是空行，暂不处理，后面统一清洗
            if not p.text.strip():
                continue
                
            title_lines_count += 1
            last_title_paragraph = p # 记录最后一行有字的标题页段落
            
            # 1. 应用基础样式 (包括双倍行距 Times New Roman 12pt)
            apply_basic_font_style(p)
            
            # 2. 居中对齐
            p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # 3. 第一行加粗 (文章主标题)
            if title_lines_count == 1:
                for run in p.runs:
                    run.bold = True

        # B. 清洗标题页与正文之间的“垃圾空行”并插入分页符
        # 策略：从 body_start - 1 倒序遍历回到 last_title_paragraph
        # 为什么要倒序？因为删除 list 元素时倒序最安全
        if last_title_paragraph:
            # 这里的逻辑是：我们已经知道 body_start 是正文第一段
            # 那么 body_start 之前，且在 last_title_paragraph 之后的所有段落，都是多余的空行
            
            # 获取 last_title_paragraph 的索引
            # 注意：由于 paragraphs 是动态对象，直接用索引可能因为之前的删除操作而变化
            # 但在这里我们还没开始删，所以是安全的。
            
            # 我们需要找到 last_title_paragraph 在 paragraphs 中的 index
            # 为了简单，我们再次遍历一下前 body_start 段
            last_title_idx = -1
            for idx in range(body_start):
                if paragraphs[idx] == last_title_paragraph:
                    last_title_idx = idx
                    break
            
            # 开始清理：从 body_start-1 倒着删到 last_title_idx+1
            if last_title_idx != -1:
                for idx in range(body_start - 1, last_title_idx, -1):
                    # 再次确认是空行才删 (双重保险)
                    if not paragraphs[idx].text.strip():
                        delete_paragraph(paragraphs[idx])
            
            # C. 在标题页最后一行内容后，强制插入分页符
            # 这样无论后面有没有内容，正文都会乖乖去下一页
            # 检查是否已经有了 page break
            if '<w:br w:type="page"/>' not in last_title_paragraph._element.xml:
                last_title_paragraph.add_run().add_break(WD_BREAK.PAGE)

    # ==========================
    # 阶段 I: 处理正文 (Body)
    # ==========================
    # 注意：由于我们在上面删除了段落，paragraphs 的长度和索引其实已经变了！
    # 如果继续用原来的 body_start 索引会导致错位。
    # 最稳妥的方法：重新获取一次 paragraphs 列表，并重新定位 body_start
    # 但由于我们删的是 body_start 之前的，body_start 之后的相对顺序没变，
    # 只是 body_start 的值应该减去删除的行数。
    
    # 为了代码的鲁棒性（防止索引越界），建议这里重新读取一下 doc.paragraphs
    # 并且简单的重新定位正文开始（正文开始就是 Title Page 后的第一个非空段）
    
    paragraphs = doc.paragraphs # 刷新列表
    
    # 重新寻找新的 body_start (因为前面删了空行，现在的 body_start 可能变小了)
    new_body_start = 0
    if config['has_title_page']:
        # 略过标题页那种居中的段落，找到第一个左对齐或者首行缩进的？
        # 不，还是用之前的逻辑：找到 Page Break 后的第一段
        for i, p in enumerate(paragraphs):
            if '<w:br w:type="page"/>' in p._element.xml:
                new_body_start = i + 1
                break
            # 如果是上面刚刚插入的 run break，xml 结构可能不同，需注意
            # 上面的 add_break(WD_BREAK.PAGE) 会在 xml 里产生 <w:br w:type="page"/>
            # 但它是在 last_title_paragraph 内部。
            
            # 简化逻辑：我们直接找 last_title_paragraph 的下一段
            if config['has_title_page'] and last_title_paragraph:
                 if p == last_title_paragraph:
                     new_body_start = i + 1
                     break
    
    # 开始处理正文
    for i in range(new_body_start, ref_start): # 注意 ref_start 可能也因为删除行而需要前移，但通常 ref 在最后，影响较小，除非 doc 很大。
        # 为保险起见，我们重新定位一下 ref_start
        pass 
    
    # --- 修正 Ref Start ---
    # 既然删除了行，索引肯定乱了。最安全的做法是：不要依赖索引数字，而是依赖对象。
    # 但为了不把代码写得太复杂，我们重新跑一次定位 ref 的逻辑是最高效的。
    _, new_ref_start = locate_structural_indices(doc, False) # has_title_page传False是为了只找Ref
    
    for i in range(new_body_start, new_ref_start):
        p = paragraphs[i]
        text = p.text.strip()
        
        if not text: continue
            
        apply_basic_font_style(p)
        pf = p.paragraph_format
        
        # Case 1: 文章主标题 (Body 的第一段)
        if i == new_body_start and config['has_article_title']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            pf.first_line_indent = Inches(0)
            for run in p.runs:
                run.bold = True
                
        # Case 2: 潜在的二级标题
        elif len(text.split()) < 15 and text[-1] not in ['.', ':', '?', '!']:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            pf.first_line_indent = Inches(0)
            pf.left_indent = Inches(0)
            for run in p.runs:
                run.bold = True
                
        # Case 3: 普通正文
        else:
            pf.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            pf.first_line_indent = Inches(0.5)
            pf.left_indent = Inches(0) 

    # ==========================
    # 阶段 II: 处理参考文献 (Refs)
    # ==========================
    if new_ref_start < len(paragraphs):
        # 强制分页 (和之前逻辑一样)
        if new_ref_start > 0:
            prev_p_ref = paragraphs[new_ref_start - 1]
            if '<w:br w:type="page"/>' not in prev_p_ref._element.xml:
                 prev_p_ref.add_run().add_break(WD_BREAK.PAGE)

        ref_title_p = paragraphs[new_ref_start]
        ref_title_p.text = "References" 
        apply_basic_font_style(ref_title_p)
        ref_title_p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        ref_title_p.paragraph_format.first_line_indent = Inches(0)
        for run in ref_title_p.runs:
            run.bold = True
            
        ref_entries = []
        for i in range(new_ref_start + 1, len(paragraphs)):
            p = paragraphs[i]
            if p.text.strip():
                ref_entries.append(p.text.strip())

        # 稍微重构一下写入逻辑，避免删除段落带来的索引困扰
        # 策略：直接清空 ref_title 之后的所有段落，然后重写
        # 1. 删除所有旧条目段落
        for i in range(len(paragraphs) - 1, new_ref_start, -1):
            delete_paragraph(paragraphs[i])
            
        # 2. 排序 (如果需要)
        if config['sort_references']:
            ref_entries.sort()
            
        # 3. 追加新段落
        for entry in ref_entries:
            new_p = doc.add_paragraph(entry)
            apply_basic_font_style(new_p)
            new_p.paragraph_format.first_line_indent = Inches(-0.5)
            new_p.paragraph_format.left_indent = Inches(0.5)

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

# --- CSS 注入：美化 & 修复侧边栏 & 高亮复制按钮 ---
    hide_streamlit_style = """
                <style>
                /* 1. 找回侧边栏和菜单 */
                #MainMenu {visibility: visible;} 
                
                /* 2. 隐藏页脚 */
                footer {visibility: hidden;}
                
                /* 3. 底部签名 */
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
                
                /* 4. 高亮检查报告的复制按钮 */
                /* 针对 Streamlit 的代码块复制按钮进行样式覆盖 */
                [data-testid="stCopyButton"] {
                    background-color: #FF4B4B !important; /* 显眼的红色背景，或者换成你喜欢的蓝色 #4B9EFF */
                    color: white !important;
                    opacity: 1 !important; /* 强制不透明 */
                    border: 1px solid white !important;
                    border-radius: 4px !important;
                    transform: scale(1.1); /* 稍微放大一点 */
                }
                
                /* 鼠标悬停时的效果 */
                [data-testid="stCopyButton"]:hover {
                    background-color: #FF2B2B !important;
                    transform: scale(1.2);
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
        "Has Title Page? ", 
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
        help="勾选后，参考文献列表将被自动按字母顺序排序。请注意，这可能会移除斜体格式（如期刊名）。"
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
