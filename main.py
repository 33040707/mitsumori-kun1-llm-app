import streamlit as st
import pandas as pd
import openai
import os
import glob
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv

# --- 設定読み込み ---
# .envファイルからAPIキーのみロード
load_dotenv()

# APIキーを取得
API_KEY = os.getenv("OPENAI_API_KEY")

# 【修正箇所】Googleドライブではなく、プロジェクト内の「data」フォルダを指定
# app.py と同じ場所にある 'data' フォルダを参照します
current_dir = os.getcwd()
DATA_FOLDER = os.path.join(current_dir, "data")

# --- 関数定義：各種ファイルからテキストを抽出する ---
def extract_text_from_files(folder_path):
    """指定フォルダ内のPDF, Excel, Wordからテキストを抽出して結合する"""
    combined_text = ""
    file_count = 0

    # 対応する拡張子
    extensions = ['*.pdf', '*.docx', '*.xlsx']
    files = []

    # フォルダ内の全ファイルを検索
    if folder_path and os.path.exists(folder_path):
        for ext in extensions:
            files.extend(glob.glob(os.path.join(folder_path, ext)))
    else:
        return "dataフォルダが見つかりません。作成してください。", 0

    if not files:
        return "dataフォルダ内にファイルが見つかりませんでした。", 0

    for file_path in files:
        file_name = os.path.basename(file_path)
        try:
            # 1. PDFの場合
            if file_path.endswith('.pdf'):
                reader = PdfReader(file_path)
                text = f"\n\n--- ファイル名: {file_name} (PDF) ---\n"
                for page in reader.pages:
                    text += page.extract_text() + "\n"
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
                xls = pd.read_excel(file_path, sheet_name=None)
                text = f"\n\n--- ファイル名: {file_name} (Excel) ---\n"
                for sheet_name, df in xls.items():
                    text += f"Sheet: {sheet_name}\n"
                    text += df.to_markdown(index=False) + "\n"
                combined_text += text
                file_count += 1

        except Exception as e:
            st.warning(f"読込エラー: {file_name} - {e}")

    return combined_text, file_count


# --- アプリ本体 ---
st.set_page_config(page_title="建設コンサル向け見積作成支援AI (Pro)", layout="wide")
st.title("🏗️ 建設コンサル見積作成支援システム (RAG対応版)")

# --- サイドバー：設定確認 ---
with st.sidebar:
    st.header("⚙️ システム設定状況")

    # APIキーの読み込み確認
    if API_KEY:
        st.success("✅ APIキー: 読込完了")
    else:
        st.error("🚫 APIキー: 未設定 (.envを確認)")

    # フォルダパスの読み込み確認
    st.markdown("### 📂 参照データフォルダ")
    if os.path.exists(DATA_FOLDER):
        st.success("✅ dataフォルダ: 接続完了")
        # フォルダの中身のファイル数を表示
        files = glob.glob(os.path.join(DATA_FOLDER, "*.*"))
        st.caption(f"ファイル数: {len(files)} 件")
    else:
        st.error("🚫 dataフォルダ: 未検出")
        st.warning("プロジェクト内に 'data' フォルダを作成してください。")

# --- メインエリア ---
st.subheader("1. 新規案件の条件入力")
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("案件名")
    location = st.text_input("施工場所")
with col2:
    work_items = st.text_area("作業内容・条件", height=100,
                              placeholder="例：\n・擁壁工（H=3.0m, L=20m）\n・場所打ち杭\n・過去のA地区の実績を参考にしたい")

if st.button("見積案を作成する", type="primary"):
    if not API_KEY:
        st.error("APIキーが設定されていません。.envファイルを確認してください。")
    elif not work_items:
        st.warning("作業内容を入力してください。")
    elif not os.path.exists(DATA_FOLDER):
        st.error("プロジェクト内に 'data' フォルダが見つかりません。作成してデータを入れてください。")
    else:
        # APIキーを設定
        openai.api_key = API_KEY

        with st.spinner('dataフォルダ内の資料を読み込み中...'):
            # RAG処理：フォルダ内のファイルをテキスト化
            context_data, count = extract_text_from_files(DATA_FOLDER)

            if count == 0:
                st.warning("有効なデータが見つかりませんでした。AIの一般知識のみで回答します。")
            else:
                st.success(f"過去資料 {count} 件を参照しました。")

        # トークン数制限対策
        if len(context_data) > 30000:
            context_data = context_data[:30000] + "\n...(データ量が多すぎるため省略)..."
            st.warning("⚠️ 参照データが多すぎるため、一部のみをAIに渡しました。")

        # プロンプト作成
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
        【作業内容】:
        {work_items}

        【参照する社内過去データ】:
        {context_data}
        """

        with st.spinner('AIが見積を計算中...'):
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=0.1,
                )

                result = response.choices[0].message.content

                st.subheader("2. 作成結果")
                st.markdown(result)

            except Exception as e:
                st.error(f"AI生成エラー: {e}")