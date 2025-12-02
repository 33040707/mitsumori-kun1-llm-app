import streamlit as st
import pandas as pd
import openai
import os
import glob
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv
# === OCR用ライブラリのインポート ===
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    OCR_AVAILABLE = True
    # 【重要】Tesseract-OCRをインストールした場所を指定してください
    # 以下は標準的なインストール例です。ご自身の環境に合わせて変更が必要です。
    # もしパスが通っていれば、この行はコメントアウトしても動く場合があります。
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
except ImportError:
    OCR_AVAILABLE = False
    print("OCRライブラリが見つかりません。pip install pytesseract pdf2image pillow を実行してください。")

# --- 設定読み込み ---
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

current_dir = os.getcwd()
DATA_FOLDER = os.path.join(current_dir, "data")

# --- 関数定義：OCR対応版 ---
def extract_text_from_files(folder_path):
    combined_text = ""
    file_count = 0
    debug_logs = []

    extensions = ['*.pdf', '*.docx', '*.xlsx']
    files = []

    if folder_path and os.path.exists(folder_path):
        for ext in extensions:
            files.extend(glob.glob(os.path.join(folder_path, ext)))
    else:
        return "dataフォルダが見つかりません。", 0, ["フォルダなし"]

    if not files:
        return "ファイルなし", 0, ["ファイルが見つかりません"]

    for file_path in files:
        file_name = os.path.basename(file_path)
        try:
            # 1. PDFの場合（OCR対応処理）
            if file_path.endswith('.pdf'):
                reader = PdfReader(file_path)
                text = f"\n\n--- ファイル名: {file_name} (PDF) ---\n"
                
                # まずは通常のテキスト抽出を試みる
                raw_text = ""
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        raw_text += extracted + "\n"
                
                # テキストが極端に少ない(50文字未満)場合は、画像PDFとみなしてOCRを試みる
                if len(raw_text.strip()) < 50:
                    debug_logs.append(f"ℹ️ {file_name} はテキスト情報が少ないため、OCR処理を試みます。時間がかかります...")
                    
                    if OCR_AVAILABLE:
                        try:
                            # PDFを画像に変換 (Popplerが必要)
                            # ※Popplerのパスが環境変数に通っていない場合、poppler_path引数での指定が必要になることがあります
                            images = convert_from_path(file_path, dpi=300)
                            ocr_result_text = ""
                            
                            progress_bar = st.progress(0)
                            for i, img in enumerate(images):
                                debug_logs.append(f"  - {i+1}/{len(images)}ページ目をOCR解析中...")
                                # 画像から日本語(jpn)の文字を読み取る
                                ocr_result_text += pytesseract.image_to_string(img, lang='jpn') + "\n"
                                progress_bar.progress((i + 1) / len(images))
                            progress_bar.empty()

                            if ocr_result_text.strip():
                                text += ocr_result_text
                                debug_logs.append(f"✅ {file_name} のOCR解析に成功しました。")
                            else:
                                text += "(OCRを実行しましたが文字を認識できませんでした)\n" + raw_text
                                debug_logs.append(f"⚠️ {file_name} のOCRを実行しましたが、有効な文字を認識できませんでした。")
                        except Exception as e_ocr:
                            text += "(OCR処理中にエラーが発生しました)\n" + raw_text
                            err_msg = str(e_ocr).lower()
                            if "tesseract is not installed" in err_msg or "found" in err_msg:
                                debug_logs.append(f"❌ OCRエラー: Tesseractが見つかりません。パス設定を確認してください。\n詳細: {e_ocr}")
                            elif "poppler" in err_msg:
                                debug_logs.append(f"❌ OCRエラー: Popplerが見つかりません。インストールとパス設定を確認してください。\n詳細: {e_ocr}")
                            else:
                                debug_logs.append(f"❌ {file_name} のOCR処理エラー: {e_ocr}")
                    else:
                         text += "(OCRライブラリが不足しているため画像文字は読めません)\n" + raw_text
                         debug_logs.append(f"⚠️ {file_name} は画像PDFの可能性がありますが、OCRライブラリが導入されていないためスキップします。")
                else:
                    # 通常のテキスト抽出で十分な文字が取れた場合
                    text += raw_text
                
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
            debug_logs.append(f"❌ 読込エラー: {file_name} - {str(e)}")

    return combined_text, file_count, debug_logs


# --- アプリ本体 ---
st.set_page_config(page_title="建設コンサル向け見積作成支援AI (OCR強化版)", layout="wide")
st.title("🏗️ 建設コンサル見積作成支援システム ")

# --- サイドバー ---
with st.sidebar:
    st.header("⚙️ 設定・状態")
    if API_KEY:
        st.success("✅ APIキー: OK")
    else:
        st.error("🚫 APIキー: 未設定")
    
    if os.path.exists(DATA_FOLDER):
        files = glob.glob(os.path.join(DATA_FOLDER, "*.*"))
        st.success(f"✅ dataフォルダ: {len(files)}ファイル")
    else:
        st.error("🚫 dataフォルダが見つかりません")
    
    st.markdown("---")
    st.markdown("### OCR機能ステータス")
    if OCR_AVAILABLE:
        st.success("✅ OCRライブラリ: 導入済み")
        st.caption("※TesseractとPopplerの外部設定が必要です。")
    else:
        st.warning("⚠️ OCRライブラリ: 未導入")
        st.caption("画像PDFは読めません。")

# --- メインエリア ---
st.subheader("1. 新規案件の条件入力")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("案件名", value="")
    location = st.text_input("施工場所", value="")
with col2:
    work_items = st.text_area("作業内容", height=100, placeholder="作業内容を入力...")

# データ確認ボタン
st.subheader("2. 参照データの確認 (デバッグ用)")
if st.button("フォルダ内のデータを読み込んで中身を確認する"):
    with st.spinner('データ解析中 (OCR処理が入ると時間がかかります)...'):
        context_data, count, logs = extract_text_from_files(DATA_FOLDER)
        
        if logs:
            st.write("--- 処理ログ ---")
            for log in logs:
                if "❌" in log: st.error(log)
                elif "⚠️" in log: st.warning(log)
                elif "ℹ️" in log: st.info(log)
                else: st.success(log)
        
        st.info(f"{count} 件のファイルを読み込みました。")
        with st.expander("クリックしてAIに送られるテキスト全文を確認する"):
            st.text(context_data)

# 見積作成ボタン
st.subheader("3. 見積作成実行")
if st.button("見積案を作成する", type="primary"):
    if not API_KEY or not os.path.exists(DATA_FOLDER):
        st.error("設定を確認してください。")
    else:
        openai.api_key = API_KEY
        with st.spinner('データ読込＆AI計算中 (OCR処理が入ると数分かかる場合があります)...'):
            # データ読み込み
            context_data, count, logs = extract_text_from_files(DATA_FOLDER)
            
            # 文字数制限 (10万文字)
            if len(context_data) > 100000:
                context_data = context_data[:100000] + "\n...(以下省略)..."
            
            # プロンプト
            system_prompt = """
            あなたは建設コンサルタントのベテラン積算技術者です。
            提供される【参照する社内過去データ】に基づき、新規案件の官公庁向け予算見積書案を作成してください。
            
            【最優先事項】
            参照データ内に類似の工種、単価、歩掛がある場合は、必ずそれらを優先して採用し、適用した根拠（例：「○○工事のデータより採用」）を摘要欄に明記してください。
            データが不鮮明な場合（OCRの誤認識など）は、文脈からベテランの知見で合理的な数値を推定・補正してください。
            """
            
            user_prompt = f"""
            【案件名】: {project_name}
            【場所】: {location}
            【作業内容】: {work_items}
            【参照する社内過去データ (OCR処理済)】:
            {context_data}
            """
            
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.1,
                )
                st.markdown(response.choices[0].message.content)
            except Exception as e:
                st.error(f"APIエラー: {e}")