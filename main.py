import streamlit as st
import pandas as pd
import openai
import os
import glob
import base64
import fitz  # PyMuPDF (PDFを画像にするライブラリ)
from docx import Document
from dotenv import load_dotenv

# --- 設定読み込み ---
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

# dataフォルダ設定
current_dir = os.getcwd()
DATA_FOLDER = os.path.join(current_dir, "data")

# --- 関数：画像をGPT-4oに送って文字にしてもらう (Cloud OCR) ---
def ocr_with_gpt4o(image_bytes, api_key):
    """
    画像のバイナリデータをGPT-4oに送信し、書かれているテキストを抽出させる
    """
    base64_image = base64.b64encode(image_bytes).decode('utf-8')
    
    client = openai.Client(api_key=api_key)
    try:
        response = client.chat.completions.create(
            model="gpt-4o",  # Vision機能が使えるモデル
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "この画像は建設工事の見積書や内訳書です。書かれている文字、数値、表の内容をすべて正確にマークダウン形式のテキストとして書き起こしてください。"},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}",
                                "detail": "high"  # 細かい文字も読めるように高画質モード
                            },
                        },
                    ],
                }
            ],
            max_tokens=2000,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"(画像読み取りエラー: {str(e)})"

# --- 関数：ファイル読み込み ---
def extract_text_from_files(folder_path, api_key):
    combined_text = ""
    file_count = 0
    debug_logs = []

    if not os.path.exists(folder_path):
        return "フォルダなし", 0, ["dataフォルダが見つかりません"]

    # PDF, Word, Excelを検索
    files = []
    for ext in ['*.pdf', '*.docx', '*.xlsx']:
        files.extend(glob.glob(os.path.join(folder_path, ext)))

    if not files:
        return "ファイルなし", 0, ["ファイルが見つかりません"]

    # 進捗バーの準備
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, file_path in enumerate(files):
        file_name = os.path.basename(file_path)
        status_text.text(f"読込中 ({idx+1}/{len(files)}): {file_name}")
        
        try:
            # 1. PDFの場合 (PyMuPDFを使用)
            if file_path.endswith('.pdf'):
                doc = fitz.open(file_path)
                text = f"\n\n--- ファイル名: {file_name} (PDF) ---\n"
                
                for page_num, page in enumerate(doc):
                    # まずテキスト抽出を試みる
                    extracted_text = page.get_text()
                    
                    # 文字がほとんどない場合(50文字未満)は「画像PDF」と判断
                    if len(extracted_text.strip()) < 50:
                        debug_logs.append(f"ℹ️ {file_name} (p.{page_num+1}) は画像と判断し、GPT-4oで読み取ります...")
                        
                        # ページを画像(Pixmap)に変換
                        pix = page.get_pixmap(dpi=200) # 200dpi程度で十分
                        img_bytes = pix.tobytes("jpeg")
                        
                        # GPT-4oに画像を送って読ませる
                        vision_text = ocr_with_gpt4o(img_bytes, api_key)
                        text += f"\n[Page {page_num+1} (Vision Read)]\n{vision_text}\n"
                    else:
                        text += extracted_text + "\n"
                
                combined_text += text
                file_count += 1

            # 2. Wordの場合
            elif file_path.endswith('.docx'):
                doc = Document(file_path)
                text = f"\n\n--- ファイル名: {file_name} (Word) ---\n"
                for para in doc.paragraphs:
                    text += para.text + "\n"
                combined_text += text
                file_count += 1

            # 3. Excelの場合
            elif file_path.endswith('.xlsx'):
                xls = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                text = f"\n\n--- ファイル名: {file_name} (Excel) ---\n"
                for sheet_name, df in xls.items():
                    df = df.fillna("")
                    text += f"Sheet: {sheet_name}\n"
                    text += df.to_markdown(index=False) + "\n"
                combined_text += text
                file_count += 1
        
        except Exception as e:
            debug_logs.append(f"❌ エラー: {file_name} - {str(e)}")

        # 進捗更新
        progress_bar.progress((idx + 1) / len(files))

    status_text.empty()
    progress_bar.empty()
    return combined_text, file_count, debug_logs


# --- アプリ画面構成 ---
st.set_page_config(page_title="建設コンサル見積作成支援AI (Vision)", layout="wide")
st.title("🏗️ 建設コンサル見積作成支援 (GPT-4o Vision版)")

# サイドバー
with st.sidebar:
    st.header("⚙️ 設定")
    if API_KEY:
        st.success("✅ APIキー: 読込完了")
    else:
        st.error("🚫 APIキー: 未設定")
    
    if os.path.exists(DATA_FOLDER):
        st.success(f"✅ dataフォルダ: {len(glob.glob(os.path.join(DATA_FOLDER, '*.*')))}ファイル")
    else:
        st.error("🚫 dataフォルダが見つかりません")

# メイン画面
st.subheader("1. 案件情報の入力")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("案件名", value="道路改良工事")
    location = st.text_input("施工場所", value="A市B町")
with col2:
    work_items = st.text_area("作業内容", height=100)

# 実行ボタン
if st.button("見積案を作成する", type="primary"):
    if not API_KEY or not os.path.exists(DATA_FOLDER):
        st.error("設定を確認してください。")
    else:
        openai.api_key = API_KEY
        
        with st.spinner('資料を解析中... (画像PDFの場合は時間がかかります)'):
            # データを読み込み（ここでGPT-4o Visionが走ります）
            context_data, count, logs = extract_text_from_files(DATA_FOLDER, API_KEY)
            
            # ログ表示
            if logs:
                with st.expander("処理ログを確認する"):
                    for log in logs:
                        st.write(log)
            
            # データ量制限
            if len(context_data) > 100000:
                context_data = context_data[:100000] + "\n...(省略)..."
            
            if count > 0:
                st.success(f"過去資料 {count} 件の内容を解析しました。見積作成を開始します。")
            else:
                st.warning("有効なデータがありませんでした。")

        # 見積作成プロンプト
        system_prompt = """
        あなたは建設コンサルタントの積算技術者です。
        提供された【過去データ】（画像解析結果を含む）に基づき、新規案件の見積書を作成してください。
        
        【指示】
        ・過去データに類似工種があれば、その単価を優先採用し、摘要に「過去実績より」と記載すること。
        ・データ読み取り結果に誤字（OCRミス）があっても、文脈から正しい建設用語や数値に補正して判断すること。
        """
        
        user_prompt = f"""
        【案件名】: {project_name}
        【場所】: {location}
        【作業内容】: {work_items}
        【過去データ】:
        {context_data}
        """
        
        with st.spinner('見積書を作成中...'):
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini", # 集計はminiで行いコスト節約
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.1,
                )
                st.markdown(response.choices[0].message.content)
            except Exception as e:
                st.error(f"APIエラー: {e}")