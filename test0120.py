import pandas as pd
from rapidfuzz import fuzz, process
from docx import Document
from io import BytesIO
from pydocx import PyDocX  # .doc ファイルを扱うためのライブラリ
import streamlit as st
from PyPDF2 import PdfReader  # PDFからテキストを抽出するためのライブラリ

# Excel ファイルを読み込む関数
def load_excel(file):
    # Handle dynamic column names and ensure proper format
    df = pd.read_excel(file, engine="openpyxl")
    if df.columns.size < 1:
        raise ValueError("The Excel file must contain at least one column with terms.")
    return df

# Word、DOC または PDF ファイルからテキストを抽出する関数
def extract_text_from_file(file, file_type):
    if file_type == "docx":
        doc = Document(file)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs])
    elif file_type == "doc":
        return PyDocX.to_text(file)
    elif file_type == "pdf":
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    else:
        return ""

# Fuzzy Matching を用いて類似語を検出する関数
def find_similar_terms(text, terms, threshold):
    words = text.split()
    detected_terms = []

    for word in words:
        # Extract multiple matches with a limit for better matching accuracy
        matches = process.extract(word, terms, scorer=fuzz.partial_ratio, limit=10)
        for match in matches:
            if match[1] >= threshold:  # Include matches above the threshold
                detected_terms.append((word, match[0], match[1]))

    return detected_terms

# 修正を適用して新しい Word ファイルを作成する関数
def create_correction_table(detected):
    correction_table = pd.DataFrame(detected, columns=["原稿内の語", "類似する用語", "類似度"])
    return correction_table

# 正誤表を使用して修正を適用する関数
def apply_corrections_with_table(text, correction_df):
    required_columns = ['誤った用語', '正しい用語']
    if not all(col in correction_df.columns for col in required_columns):
        raise ValueError(f"正誤表に必要な列が不足しています: {required_columns}")

    for _, row in correction_df.iterrows():
        incorrect = row['誤った用語']
        correct = row['正しい用語']
        text = text.replace(incorrect, correct)
    return text

# 利用漢字表を使用して修正を適用する関数
def apply_kanji_table(text, kanji_df):
    required_columns = ['ひらがな', '漢字']
    if not all(col in kanji_df.columns for col in required_columns):
        raise ValueError(f"利用漢字表に必要な列が不足しています: {required_columns}")

    for _, row in kanji_df.iterrows():
        hiragana = row['ひらがな']
        kanji = row['漢字']
        text = text.replace(hiragana, kanji)
    return text

# Streamlit アプリケーション
st.title("用語チェックアプリ")

st.write("以下のファイルを個別にアップロードしてください:")
word_file = st.file_uploader("原稿ファイル (Word, DOC, PDF):", type=["docx", "doc", "pdf"])
terms_file = st.file_uploader("用語集ファイル (Excel):", type=["xlsx"])
correction_file = st.file_uploader("正誤表ファイル (Excel, 任意):", type=["xlsx"])
kanji_file = st.file_uploader("利用漢字表ファイル (Excel, 任意):", type=["xlsx"])

if word_file and terms_file:
    # 用語集をDataFrameとして読み込み
    try:
        terms_df = load_excel(terms_file)
    except Exception as e:
        st.error(f"用語集ファイルの読み込み中にエラーが発生しました: {e}")

    if terms_df.empty:
        st.error("用語集ファイルが空です。少なくとも1つの用語を含む必要があります。")
    else:
        terms = terms_df.iloc[:, 0].dropna().astype(str).tolist()

        # 原稿ファイルの読み込み
        file_type = word_file.name.split(".")[-1]
        original_text = extract_text_from_file(word_file, file_type)

        # 類似度の閾値を入力
        threshold = st.slider("類似度の閾値を設定してください (50-100):", min_value=50, max_value=99, value=70)
        detected = find_similar_terms(original_text, terms, threshold)

        # 結果を表示
        if detected:
            st.success("類似語が検出されました！")
            correction_table = create_correction_table(detected)
            st.dataframe(correction_table)

            # 修正箇所を表形式でダウンロード
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                correction_table.to_excel(writer, index=False, sheet_name="Corrections")
            st.download_button(
                label="修正箇所をダウンロード",
                data=output.getvalue(),
                file_name="correction_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        else:
            st.warning("類似する語は検出されませんでした。")

        # 正誤表がアップロードされている場合、修正を適用
        if correction_file:
            try:
                correction_df = load_excel(correction_file)
                original_text = apply_corrections_with_table(original_text, correction_df)
                st.success("正誤表を適用しました！")
            except Exception as e:
                st.error(f"正誤表の処理中にエラーが発生しました: {e}")

        # 利用漢字表がアップロードされている場合、修正を適用
        if kanji_file:
            try:
                kanji_df = load_excel(kanji_file)
                original_text = apply_kanji_table(original_text, kanji_df)
                st.success("利用漢字表を適用しました！")
            except Exception as e:
                st.error(f"利用漢字表の処理中にエラーが発生しました: {e}")

else:
    st.warning("原稿ファイルと用語集ファイルの両方をアップロードしてください！")
