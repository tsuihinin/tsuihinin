import streamlit as st
import whisper
import tempfile
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from funasr import AutoModel
import re

# ---------------------------- 頁面設定 ----------------------------
st.set_page_config(
    page_title="語音轉文字 + 標點還原",
    page_icon="🎤",
    layout="centered"
)

st.title("🎤 語音轉文字（含標點符號）")
st.markdown("上傳 MP3 / WAV / M4A，自動轉文字 + 還原標點 + 產生 Word 檔")

# ---------------------------- 快取模型（很重要，避免每次重新下載） ----------------------------
@st.cache_resource
def load_whisper_model(model_size="base"):
    return whisper.load_model(model_size)

@st.cache_resource
def load_punc_model():
    # 使用 FunASR 的中文標點模型
    return AutoModel(model="ct-punc", device="cpu", disable_update=True)

# ---------------------------- 側邊欄設定 ----------------------------
with st.sidebar:
    st.header("設定")
    model_size = st.selectbox(
        "Whisper 模型大小",
        options=["tiny", "base", "small"],
        index=1,
        help="base 速度與準確度較平衡，tiny 最快但準確度較低"
    )
    language = st.selectbox(
        "語言",
        options=["zh", "en", "auto"],
        index=0,
        help="zh = 中文，en = 英文，auto = 自動偵測"
    )

# ---------------------------- 上傳音檔 ----------------------------
uploaded_file = st.file_uploader(
    "上傳語音檔（支援 mp3 / wav / m4a）",
    type=["mp3", "wav", "m4a", "ogg"]
)

if uploaded_file is not None:
    st.audio(uploaded_file, format=f"audio/{uploaded_file.type.split('/')[-1]}")

    if st.button("開始轉換", type="primary"):
        with st.spinner("正在處理中，請稍候..."):
            # 1. 儲存暫存音檔
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                tmp.write(uploaded_file.read())
                audio_path = tmp.name

            try:
                # 2. Whisper 轉寫
                whisper_model = load_whisper_model(model_size)
                result = whisper_model.transcribe(
                    audio_path,
                    language=None if language == "auto" else language,
                    fp16=False
                )
                raw_text = result["text"].strip()

                # 3. 標點符號還原
                punc_model = load_punc_model()
                punc_result = punc_model.generate(input=raw_text)
                punctuated_text = punc_result[0]["text"]

                # 4. 顯示結果
                st.subheader("轉寫結果（已加入標點）")
                st.write(punctuated_text)

                # 5. 產生 Word 檔
                doc = Document()

                # 設定中文字型
                def set_run_font(run, font_name="微軟正黑體", size=12):
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                    run.font.size = Pt(size)

                # 標題
                title = doc.add_heading("語音轉文字結果", level=1)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # 基本資訊
                info = doc.add_paragraph()
                run = info.add_run(f"轉換時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                set_run_font(run, size=10)
                run.font.color.rgb = RGBColor(100, 100, 100)

                info2 = doc.add_paragraph()
                run2 = info2.add_run(f"原始檔案：{uploaded_file.name}")
                set_run_font(run2, size=10)
                run2.font.color.rgb = RGBColor(100, 100, 100)

                doc.add_paragraph()

                # 內容
                doc.add_heading("轉寫內容（已還原標點）", level=2)

                # 把長文切成段落（提升可讀性）
                paragraphs = re.split(r'([。！？])', punctuated_text)
                current_para = ""
                for i in range(0, len(paragraphs), 2):
                    sentence = paragraphs[i]
                    punct = paragraphs[i+1] if i+1 < len(paragraphs) else ""
                    current_para += sentence + punct
                    if len(current_para) > 80 or punct in "。！？":
                        p = doc.add_paragraph()
                        run = p.add_run(current_para.strip())
                        set_run_font(run, size=12)
                        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                        p.paragraph_format.space_after = Pt(8)
                        p.paragraph_format.first_line_indent = Inches(0.3)
                        current_para = ""

                if current_para.strip():
                    p = doc.add_paragraph()
                    run = p.add_run(current_para.strip())
                    set_run_font(run, size=12)
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                    p.paragraph_format.space_after = Pt(8)
                    p.paragraph_format.first_line_indent = Inches(0.3)

                # 儲存並提供下載
                output_filename = "語音轉文字結果.docx"
                doc.save(output_filename)

                with open(output_filename, "rb") as f:
                    st.download_button(
                        label="⬇️ 下載 Word 檔案",
                        data=f,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                st.success("轉換完成！")

            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")
            finally:
                # 清理暫存檔
                if os.path.exists(audio_path):
                    os.remove(audio_path)
