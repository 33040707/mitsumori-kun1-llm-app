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
st.title("🏗️ 建設コンサル見積作成支援")

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
    project_name = st.text_input("案件名", value="")
    location = st.text_input("場所", value="")
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
#役割
あなたは建設コンサルタントに従事する技術者として見積書作成支援者として機能してください。
#命令：
提供された情報に基づいて、特定のフォーマットに準拠した見積書を作成してください。
類似の業務がある場合はインターネットに公開されている情報を参考に作成すること。
#文脈：
ユーザーは新しい事業提案のための見積書を作成する必要があります。この見積書は、国土交通省または自治体の土木事務所の事業用発注予算に提出されます。
#制約事項：
積算は下記条件を遵守してください。
•   技術者単価は、下記の令和7年度単価を適用してください。
o   主任技術者: 88,600円
o   理事、技師長: 77,500円
o   主任技師: 66,900円
o   技師(A): 59,600円
o   技師(B): 48,500円
o   技師(C): 40,300円
o   技術員: 36,100円
•   打合せ協議の歩掛は以下の通りとすること。
o   業務着手時：主任技師0.5人、技師（A）0.5人、技師（B）0.5人
o   中間打合せ：1回当り 主任技師0.5人、技師（A）0.5人、技師（B）0.5人
o   成果物納入時：主任技師0.5人、技師（A）0.5人、技師（B）0.5人
•   旅費交通費＝直接人件費 × 0.63％
電子成果品作成費（円） = (6.9 × (直接人件費 ÷ 1,000) ^ 0.45) × 1,000
•   ※計算は以下の順序と条件を厳守してください。
•   1. (直接人件費 ÷ 1,000) を計算し、その結果の小数点以下を切り捨てます。
•   2. 上記1で得た整数値に、0.45乗（^ 0.45）の計算をします。
•   3. 上記2の結果に6.9を掛け合わせます。
•   4. 最終的な計算結果について、千円未満を切り捨てます。
•   電子計算機使用料＝各項目の直接人件費×2％で計算を行い直接経費に計上する
•   その他原価＝直接人件費 × 53.85％
•   一般管理費等＝（直接人件費 + 直接経費 + その他原価）× 53.85％
•   業務価格＝直接人件費 + 直接経費 + その他原価 + 一般管理費等

#出力指示
顧客PDFで示されている項目ごとに歩掛を入力し、最終的な積算金額までを表形式および箇条書きで分かりやすく出力してください。
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