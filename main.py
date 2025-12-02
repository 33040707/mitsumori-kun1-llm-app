import streamlit as st
import pandas as pd
import openai
import os
import glob
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv

# --- 設定読み込み ---
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")

# dataフォルダの設定
current_dir = os.getcwd()
DATA_FOLDER = os.path.join(current_dir, "data")

# --- 関数定義：エラーハンドリング強化版 ---
def extract_text_from_files(folder_path):
    combined_text = ""
    file_count = 0
    debug_logs = []  # エラーログ用

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
            # 1. PDFの場合
            if file_path.endswith('.pdf'):
                reader = PdfReader(file_path)
                text = f"\n\n--- ファイル名: {file_name} (PDF) ---\n"
                page_texts = []
                for i, page in enumerate(reader.pages):
                    extracted = page.extract_text()
                    if extracted:
                        page_texts.append(extracted)
                    else:
                        debug_logs.append(f"⚠️ {file_name} の {i+1}ページ目は文字が抽出できませんでした（画像PDFの可能性があります）。")
                
                if not page_texts:
                    text += "(このPDFからは文字情報を取得できませんでした)"
                else:
                    text += "\n".join(page_texts)
                
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
                # engine='openpyxl' を明示的に指定
                xls = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                text = f"\n\n--- ファイル名: {file_name} (Excel) ---\n"
                for sheet_name, df in xls.items():
                    # NaN（空白）を空文字に置換して読みやすくする
                    df = df.fillna("")
                    text += f"Sheet: {sheet_name}\n"
                    text += df.to_markdown(index=False) + "\n"
                combined_text += text
                file_count += 1

        except Exception as e:
            error_msg = f"❌ 読込エラー: {file_name} - {str(e)}"
            debug_logs.append(error_msg)
            # Excel特有のエラーヒント
            if "openpyxl" in str(e):
                debug_logs.append("💡 ヒント: pip install openpyxl を実行してください。")
            if "Permission denied" in str(e):
                debug_logs.append("💡 ヒント: ファイルが開かれたままになっていませんか？閉じてから再試行してください。")

    return combined_text, file_count, debug_logs


# --- アプリ本体 ---
st.set_page_config(page_title="建設コンサル見積作成支援AI (Debug版)", layout="wide")
st.title("🏗️ 建設コンサル見積作成支援システム (Debug Mode)")

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

# --- メインエリア ---
st.subheader("1. 新規案件の条件入力")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("案件名", value="テスト案件")
    location = st.text_input("施工場所", value="テスト市")
with col2:
    work_items = st.text_area("作業内容", height=100, placeholder="作業内容を入力...")

# データ読み込みテストボタン（実行前に確認できるように分離）
st.subheader("2. 参照データの確認 (デバッグ用)")
if st.button("フォルダ内のデータを読み込んで中身を確認する"):
    with st.spinner('データ解析中...'):
        context_data, count, logs = extract_text_from_files(DATA_FOLDER)
        
        # エラーログの表示
        if logs:
            st.error("以下の問題が発生しました:")
            for log in logs:
                st.write(log)
        
        # 読み取れたテキストの表示
        st.info(f"{count} 件のファイルを読み込みました。")
        with st.expander("クリックしてAIに送られるテキスト全文を確認する"):
            st.text(context_data)
            if len(context_data) < 100:
                st.warning("⚠️ テキストが非常に少ないか、空です。PDFが画像（スキャン）データの可能性があります。")

# 見積作成ボタン
st.subheader("3. 見積作成実行")
if st.button("見積案を作成する", type="primary"):
    if not API_KEY:
        st.error("APIキー設定を確認してください。")
    else:
        openai.api_key = API_KEY
        
        # データ再読み込み
        context_data, count, logs = extract_text_from_files(DATA_FOLDER)
        
        # 文字数制限を緩和 (10万文字まで)
        if len(context_data) > 100000:
            context_data = context_data[:100000] + "\n...(以下省略)..."
            st.warning("⚠️ データ量が非常に多いため、一部を省略しました。")

        system_prompt = """
#役割
あなたは建設コンサルタントの積算技術者です。
過去の参照データに基づき、新規案件の見積書を作成してください。

#最優先指示
1. 【参照データ】の中に、類似の工種や単価がある場合は、**計算ルールよりも優先して**その単価を採用してください。
2. 参照データにない項目のみ、後述の【積算ルール】に従って計算してください。

#積算ルール
（省略：ユーザーの指定した計算式・単価表）
•   技術者単価: 令和7年度単価適用
... (中略) ...
        """

        user_prompt = f"""
        【案件名】: {project_name}
        【場所】: {location}
        【作業内容】:
        {work_items}

        【参照する社内過去データ (RAG)】:
        {context_data}
        """

        with st.spinner('AIが計算中...'):
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini", # または gpt-4o
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.1,
                )
                st.markdown(response.choices[0].message.content)
            except Exception as e:
                st.error(f"APIエラー: {e}")