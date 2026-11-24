import streamlit as st
import whisper
import tempfile
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re

st.set_page_config(
    page_title="語音轉文字",
    page_icon="🎤",
    layout="centered"
)

st.title("🎤 語音轉文字工具")
st.caption("上傳 MP3 / WAV / M4A → 自動轉文字 → 產生 Word 檔下載")

# ---------------------------- 快取模型 ----------------------------
@st.cache_resource
def load_model(model_size: str):
    return whisper.load_model(model_size)

# ---------------------------- 側邊欄 ----------------------------
with st.sidebar:
    st.header("設定")
    model_size = st.selectbox(
        "模型大小（建議 base）",
        options=["tiny", "base", "small"],
        index=1,
        help="tiny 最快、記憶體最少；base 平衡；small 較準但較慢"
    )
    language = st.selectbox(
        "語言",
        options=["zh", "en", "auto"],
        index=0
    )

# ---------------------------- 上傳檔案 ----------------------------
uploaded_file = st.file_uploader(
    "請上傳語音檔",
    type=["mp3", "wav", "m4a", "ogg", "flac"]
)

if uploaded_file is not None:
    st.audio(uploaded_file)

    if st.button("開始轉換", type="primary", use_container_width=True):
        progress = st.progress(0, text="準備中...")

        try:
            # 儲存暫存檔
            suffix = os.path.splitext(uploaded_file.name)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded_file.getvalue())
                audio_path = tmp.name

            progress.progress(20, text="載入模型中...")
            model = load_model(model_size)

            progress.progress(40, text="正在轉寫語音...")
            result = model.transcribe(
                audio_path,
                language=None if language == "auto" else language,
                fp16=False          # CPU 必須關掉 fp16
            )
            text = result["text"].strip()

            progress.progress(70, text="產生 Word 檔中...")

            # ---------- 產生 Word ----------
            doc = Document()

            def set_font(run, size=12, name="微軟正黑體"):
                run.font.name = name
                run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
                run.font.size = Pt(size)

            # 標題
            title = doc.add_heading("語音轉文字結果", level=1)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 資訊
            p = doc.add_paragraph()
            run = p.add_run(f"轉換時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            set_font(run, 10)
            run.font.color.rgb = RGBColor(100, 100, 100)

            p = doc.add_paragraph()
            run = p.add_run(f"原始檔案：{uploaded_file.name}")
            set_font(run, 10)
            run.font.color.rgb = RGBColor(100, 100, 100)

            doc.add_paragraph()
            doc.add_heading("轉寫內容", level=2)

            # 簡單切段（提升可讀性）
            # 先用常見句號切，沒有的話就依字數切
            sentences = re.split(r'([。！？.!?])', text)
            current = ""
            for i in range(0, len(sentences), 2):
                part = sentences[i]
                punct = sentences[i+1] if i+1 < len(sentences) else ""
                current += part + punct
                if len(current) >= 90 or punct in "。！？.!?":
                    if current.strip():
                        para = doc.add_paragraph()
                        run = para.add_run(current.strip())
                        set_font(run, 12)
                        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                        para.paragraph_format.space_after = Pt(10)
                        para.paragraph_format.first_line_indent = Inches(0.3)
                    current = ""
            if current.strip():
                para = doc.add_paragraph()
                run = para.add_run(current.strip())
                set_font(run, 12)
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                para.paragraph_format.space_after = Pt(10)
                para.paragraph_format.first_line_indent = Inches(0.3)

            output_name = "語音轉文字結果.docx"
            doc.save(output_name)

            progress.progress(100, text="完成！")

            # 顯示結果
            st.subheader("轉寫結果")
            st.write(text)

            # 下載按鈕
            with open(output_name, "rb") as f:
                st.download_button(
                    label="⬇️ 下載 Word 檔案",
                    data=f,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            st.success("轉換完成！")

        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")
            st.info("請到右下角 Manage app → Logs 查看詳細錯誤訊息")
        finally:
            if 'audio_path' in locals() and os.path.exists(audio_path):
                os.remove(audio_path)
