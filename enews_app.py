import streamlit as st
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import re
import base64
from io import BytesIO
from bs4 import BeautifulSoup
import traceback

# --- 樣式設定 ---
STYLE_H1 = "font-size: 28px; font-weight: bold; line-height: 1.8; margin-bottom: 20px;"
STYLE_H2 = "font-size: 22px; font-weight: bold; line-height: 2; margin-top: 25px; margin-bottom: 10px;"
STYLE_H3 = "font-size: 20px; font-weight: bold; line-height: 2; margin-top: 20px; margin-bottom: 10px;"
STYLE_P = "font-size: 20px; line-height: 2; margin-bottom: 15px;"
# 圖片的HTML樣式依然保留，方便使用者複製貼上
STYLE_IMG_TAG = 'style="display: block; margin-left: auto; margin-right: auto; max-width: 90%; height: auto; border-radius: 8px;"'


# --- 核心邏輯函式 ---

def get_heading_level(text, is_first_p):
    """使用啟發式規則判斷文字段落的標題級別"""
    stripped_text = text.strip()
    if not stripped_text: return None
    if is_first_p: return 'h1'
    if re.match(r'^([一二三四五六七八九十]+[、．\.])|(\d+\.)|(\([一二三四五六七八九十]+\))|(第[一二三四五六七八九十]+[章節])', stripped_text): return 'h2'
    if len(stripped_text) < 35 and stripped_text[-1] not in ['。', '，', '？', '！', '；', ':', '：', ')', '）', '.', '"', '”']: return 'h3'
    return 'p'

def generate_meta_description(html_parts):
    """從HTML片段中找到第一個<p>標籤的內容，生成meta description"""
    for part in html_parts:
        if part.strip().startswith('<p'):
            soup = BeautifulSoup(part, 'html.parser')
            p_text = soup.get_text().strip()
            meta_content = (p_text[:150] + '...') if len(p_text) > 150 else p_text
            meta_content = meta_content.replace('"', "'")
            return f'<meta name="description" content="{meta_content}">'
    return '<meta name="description" content="一篇精彩的文章內容。">'

def get_paragraph_text_v7(p: Paragraph):
    """(v7) 終極文字提取函式，直接遍歷XML"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    try:
        text_nodes = p._p.xpath('.//w:t/text()', namespaces=ns)
        return ''.join(text_nodes)
    except Exception:
        return ''.join(run.text for run in p.runs)

def process_paragraph_v9(p, is_first_paragraph_ref, image_counter):
    """(v9) 處理一個段落，將圖片替換為更清晰的註解"""
    local_html_parts = []
    # 處理圖片：只增加計數並生成註解
    if p._p.xpath('.//a:blip/@r:embed'):
        image_counter['count'] += 1
        # 產生更清晰、更不可能錯過的註解
        placeholder = (
            f'\n\n<!-- ########## 請在此處插入第 {image_counter["count"]} 張圖片 ########## -->\n'
            f'<!-- 步驟：(1)從下方下載圖片 (2)上傳到您的平台取得網址 (3)用網址替換掉下面的"請貼上圖片網址" -->\n'
            f'<img src="請貼上圖片網址" alt="文章插圖" {STYLE_IMG_TAG}>\n'
            f'<!-- ########## 第 {image_counter["count"]} 張圖片結束 ########## -->\n\n'
        )
        local_html_parts.append(placeholder)
    
    # 處理文字
    text = get_paragraph_text_v7(p).strip()
    if text:
        level = get_heading_level(text, is_first_paragraph_ref[0])
        if level:
            style_map = {'h1': STYLE_H1, 'h2': STYLE_H2, 'h3': STYLE_H3, 'p': STYLE_P}
            local_html_parts.append(f'<{level} style="{style_map[level]}">{text}</{level}>')
            if is_first_paragraph_ref[0]:
                is_first_paragraph_ref[0] = False
    return local_html_parts

def iter_block_items(parent):
    """生成器，可以遍歷文件或表格儲存格中的段落和表格"""
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def process_docx_v9(file_stream):
    """(v9) 最終優化版，分離圖片和文字"""
    doc = docx.Document(file_stream)
    html_parts = []
    images_data = []
    is_first_paragraph_ref = [True]
    image_counter = {'count': 0} # 使用字典來傳遞引用
    
    # 提取所有圖片數據
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            try:
                img_data = rel.target_part.blob
                ext = rel.target_part.content_type.split('/')[-1]
                if ext == 'jpeg': ext = 'jpg'
                images_data.append({'data': img_data, 'ext': ext})
            except Exception:
                pass
    
    # 遍歷文件內容生成HTML
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            html_parts.extend(process_paragraph_v9(block, is_first_paragraph_ref, image_counter))
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for paragraph_in_cell in iter_block_items(cell):
                        if isinstance(paragraph_in_cell, Paragraph):
                            html_parts.extend(process_paragraph_v9(paragraph_in_cell, is_first_paragraph_ref, image_counter))
    return html_parts, images_data


def build_final_html_v9(html_parts, meta_description, title="您的文章標題"):
    """(v9) 將所有HTML片段組合成一個完整的HTML文件，移除外層div"""
    for part in html_parts:
        if part.strip().startswith('<h1'):
            soup = BeautifulSoup(part, 'html.parser')
            title = soup.get_text().strip()
            break
    body_content = "\n".join(html_parts)
    # 移除 <div class="container">，並將其樣式合併到 body
    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    {meta_description}
    <title>{title}</title>
</head>
<body>
{body_content}
</body>
</html>"""

# --- Streamlit UI 介面 ---

st.set_page_config(page_title="E-news 自動化排版工具", page_icon="⚙️", layout="wide")
st.title("E-news 自動化排版工具 ⚙️")
st.markdown("歡迎polaris的小夥伴，請直接上傳您的文章檔案吧！")

uploaded_file = st.file_uploader("點擊此處上傳檔案", type=["docx"])

if uploaded_file is not None:
    try:
        with st.spinner("正在分析您的文件，請稍候..."):
            html_parts, images_data = process_docx_v9(uploaded_file)

        if html_parts or images_data:
            st.success("分析完成！")
            
            # --- 顯示HTML結果 ---
            st.subheader("📋 複製您的 HTML 程式碼")
            meta_tag = generate_meta_description(html_parts)
            final_html = build_final_html_v9(html_parts, meta_tag, title=uploaded_file.name)
            st.text_area("下方的程式碼已移除外層邊界，並將圖片位置用註解標示出來。", final_html, height=400)
            
            # --- 顯示圖片下載區塊 ---
            if images_data:
                st.subheader("🖼️ 下載文章中的圖片")
                st.info("請將下方圖片依序上傳至您的平台，再將網址貼回上方HTML中對應的位置。")
                cols = st.columns(4)
                for i, img in enumerate(images_data):
                    col = cols[i % 4]
                    with col:
                        col.image(img['data'], caption=f"第 {i+1} 張圖片", use_column_width=True)
                        col.download_button(
                            label=f"下載第 {i+1} 張圖",
                            data=img['data'],
                            file_name=f"image_{i+1}.{img['ext']}",
                            mime=f"image/{img['ext']}"
                        )
            
            # --- 顯示預覽 ---
            st.subheader("👀 即時預覽 (無圖片)")
            st.info("此處預覽不包含圖片，僅供確認文字排版。")
            st.components.v1.html(final_html, height=600, scrolling=True)

        else:
            st.error("處理完成，但未能從文件中提取任何內容。")

    except Exception as e:
        st.error("處理檔案時發生錯誤！")
        st.code(traceback.format_exc())
