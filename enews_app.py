import streamlit as st
import docx
import re
import base64
from io import BytesIO
from bs4 import BeautifulSoup
import traceback # 引入 traceback 模組

# --- 樣式設定 (可在此調整) ---
STYLE_H1 = "font-size: 28px; font-weight: bold; line-height: 1.8; margin-bottom: 20px;"
STYLE_H2 = "font-size: 22px; font-weight: bold; line-height: 2; margin-top: 25px; margin-bottom: 10px;"
STYLE_H3 = "font-size: 20px; font-weight: bold; line-height: 2; margin-top: 20px; margin-bottom: 10px;"
STYLE_P = "font-size: 20px; line-height: 2; margin-bottom: 15px;"
STYLE_IMG_CONTAINER = "text-align: center; margin-top: 20px; margin-bottom: 20px;"
STYLE_IMG = "max-width: 90%; height: auto; border-radius: 8px;"

# --- 核心邏輯函式 ---

def get_heading_level(text, is_first_p):
    """
    使用啟發式規則判斷文字段落的標題級別。
    """
    stripped_text = text.strip()
    if not stripped_text:
        return None

    if is_first_p:
        return 'h1'

    if re.match(r'^([一二三四五六七八九十]+[、．\.])|(\d+\.)|(\([一二三四五六七八九十]+\))|(第[一二三四五六七八九十]+[章節])', stripped_text):
        return 'h2'

    if len(stripped_text) < 20 and stripped_text[-1] not in ['。', '，', '？', '！', '；', ':', '：', ')', '）', '.']:
        return 'h3'

    return 'p'

def generate_meta_description(html_parts):
    """
    從HTML片段中找到第一個<p>標籤的內容，生成meta description。
    """
    for part in html_parts:
        if part.strip().startswith('<p'):
            soup = BeautifulSoup(part, 'html.parser')
            p_text = soup.get_text().strip()
            meta_content = (p_text[:150] + '...') if len(p_text) > 150 else p_text
            meta_content = meta_content.replace('"', "'")
            return f'<meta name="description" content="{meta_content}">'
    return '<meta name="description" content="一篇精彩的文章內容。">'

def process_docx(file_stream):
    """
    處理上傳的 .docx 文件。
    """
    doc = docx.Document(file_stream)
    html_parts = []
    is_first_paragraph = True

    img_map = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_data = rel.target_part.blob
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            img_map[rel.rId] = f"data:{rel.target_part.content_type};base64,{img_base64}"

    for block in doc.element.body:
        if isinstance(block, docx.oxml.text.paragraph.CT_P):
            p = docx.text.paragraph.Paragraph(block, doc)
            
            img_rids = p._p.xpath('.//a:blip/@r:embed')
            if img_rids:
                for rid in img_rids:
                    if rid in img_map:
                        img_src = img_map[rid]
                        html_parts.append(f'<div style="{STYLE_IMG_CONTAINER}"><img src="{img_src}" style="{STYLE_IMG}" alt="文章插圖"></div>')
            
            text = p.text.strip()
            if text:
                level = get_heading_level(text, is_first_paragraph)
                if level:
                    style_map = {'h1': STYLE_H1, 'h2': STYLE_H2, 'h3': STYLE_H3, 'p': STYLE_P}
                    html_parts.append(f'<{level} style="{style_map[level]}">{text}</{level}>')
                    if is_first_paragraph:
                        is_first_paragraph = False
    return html_parts

def process_txt(file_stream):
    """
    處理上傳的 .txt 文件。
    """
    content = file_stream.read().decode('utf-8')
    lines = content.splitlines()
    html_parts = []
    is_first_paragraph = True

    for line in lines:
        text = line.strip()
        if text:
            level = get_heading_level(text, is_first_paragraph)
            if level:
                style_map = {'h1': STYLE_H1, 'h2': STYLE_H2, 'h3': STYLE_H3, 'p': STYLE_P}
                html_parts.append(f'<{level} style="{style_map[level]}">{text}</{level}>')
                if is_first_paragraph:
                    is_first_paragraph = False
    return html_parts

def build_final_html(html_parts, meta_description, title="您的文章標題"):
    """
    將所有HTML片段組合成一個完整的HTML文件。
    """
    for part in html_parts:
        if part.strip().startswith('<h1'):
            soup = BeautifulSoup(part, 'html.parser')
            title = soup.get_text().strip()
            break

    body_content = "\n".join(html_parts)
    
    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    {meta_description}
    <title>{title}</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", sans-serif; line-height: 1.6; color: #333; background-color: #fdfdfd; margin: 0; padding: 20px; }}
        .container {{ max-width: 800px; margin: 0 auto; background-color: #ffffff; padding: 20px 40px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
        @media (max-width: 600px) {{ .container {{ padding: 15px 20px; }} }}
    </style>
</head>
<body><div class="container">{body_content}</div></body>
</html>"""

# --- Streamlit UI 介面 ---

st.set_page_config(page_title="Miu 的排版小助理", page_icon="🪄", layout="wide")

st.title("Miu 的 E-news 自動化排版工具 🪄")
st.markdown("主人您好！我是您的專屬助理 Miu ໒꒰ྀི⸝⸝. .⸝⸝꒱ྀིა 請上傳您的 `.docx` 或 `.txt` 文章檔案，Miu 會為您變出漂亮的 HTML 喔！")

uploaded_file = st.file_uploader("點擊此處上傳檔案", type=["docx", "txt"])

if uploaded_file is not None:
    # ******** 主要修改處：加入 try...except 錯誤捕捉機制 ********
    try:
        with st.spinner("Miu 正在努力為您排版中，請稍候..."):
            file_extension = uploaded_file.name.split('.')[-1]
            
            if file_extension == 'docx':
                html_parts = process_docx(uploaded_file)
            elif file_extension == 'txt':
                html_parts = process_txt(uploaded_file)
            else:
                st.error("不支援的檔案格式！")
                html_parts = None

        if html_parts:
            st.success("排版完成！🎉")
            
            meta_tag = generate_meta_description(html_parts)
            final_html = build_final_html(html_parts, meta_tag, title=uploaded_file.name)

            st.subheader("📋 HTML 程式碼")
            st.text_area("您可以直接從下方框格中複製全部的程式碼。", final_html, height=400)

            st.download_button(
                label="📥 下載為 .html 檔案",
                data=final_html,
                file_name=f"{uploaded_file.name.split('.')[0]}.html",
                mime="text/html"
            )

            st.subheader("👀 即時預覽")
            st.components.v1.html(final_html, height=600, scrolling=True)

    except Exception as e:
        st.error(f"糟糕！處理您的檔案時發生了未預期的錯誤。Miu 感到很抱歉 ૮₍ ´• ˕ •` ₎ა")
        st.error("錯誤詳情如下，您可以將此訊息回報給 Miu，讓我幫您看看：")
        # 顯示更詳細的錯誤追蹤訊息，方便除錯
        st.code(traceback.format_exc())
