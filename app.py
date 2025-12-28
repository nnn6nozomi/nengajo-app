import streamlit as st
import pandas as pd
from pdf_generator import generate_nengajo_pdf, generate_preview_image
import io
import os

# ページ設定
st.set_page_config(page_title="年賀状作成アプリ", layout="wide")

# ==========================================
# 🕵️‍♀️ フォントファイル検証エリア
# ==========================================
st.title("📮 年賀状 宛名印刷アプリ")

TARGET_FONT_NAME = "brush.ttf"

if os.path.exists(TARGET_FONT_NAME):
    st.success(f"✅ フォント「{TARGET_FONT_NAME}」を認識中。書き初め風で出力します。")
else:
    st.error(f"⚠️ フォント「{TARGET_FONT_NAME}」が見つかりません（現在は標準フォントになります）")
    st.info(f"💡 ヒント: 筆文字にしたい場合は、フォントファイルを「{TARGET_FONT_NAME}」という名前にして置いてください。")

st.markdown("---")

# ==========================================
# アプリ本編
# ==========================================

col1, col2 = st.columns([1.5, 1])

if 'df_edited' not in st.session_state:
    st.session_state.df_edited = None

# --- 左カラム：データ編集 ---
with col1:
    st.subheader("1. 住所録データの読み込み")

    # ▼▼▼ ガイドエリア（ここを強化しました） ▼▼▼
    with st.expander("📌 データ作成ガイド（AI用プロンプトあり）"):
        st.markdown("### 1. 自分でExcelを作る場合")
        st.markdown("""
        1行目に**「名前」**と**「住所」**という列を作ってください。
        * **名前**: 宛名（例：山田 太郎）
        * **住所**: 郵便番号込みの住所（例：100-0001 東京都...）
        """)
        
        st.markdown("---")
        
        st.markdown("### 2. AI(ChatGPT)に作らせる場合")
        st.write("手元の住所リスト（メールやメモなど）を、以下の文章と一緒にChatGPTに貼り付けると、このアプリ用の表を一瞬で作ってくれます。")
        
        # コピーしやすいようにコードブロックで表示
        st.code("""
あなたは優秀なデータ編集アシスタントです。
以下のテキストデータから「氏名」と「住所（郵便番号含む）」を抽出し、Excelに貼り付けられる形式の表（マークダウンの表）を作成してください。

【出力フォーマットのルール】
1. 列名は必ず「名前」「住所」「印刷状態」の3列にしてください。
2. 「住所」列には、郵便番号と住所を繋げて記載してください（例：100-0001 東京都...）。
3. 「印刷状態」列には、すべて「印刷対象」と入力してください。
4. 連名（夫婦など）の場合は、行を分けて1行につき1名にしてください。

【データ】
(ここに住所録のテキストを貼り付けてください)
        """, language="text")
        st.caption("👆 右上のコピーボタンを押して、ChatGPTに入力してください。")
    # ▲▲▲ ここまで ▲▲▲

    uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsx)", type=["xlsx"])

    target_df = pd.DataFrame() 

    if uploaded_file is not None:
        if st.session_state.df_edited is None:
            try:
                df_raw = pd.read_excel(uploaded_file)
                
                # 必須列チェック
                required_cols = ["名前", "住所"]
                missing_cols = [c for c in required_cols if c not in df_raw.columns]
                
                if missing_cols:
                    st.error(f"⚠️ Excelに「{', '.join(missing_cols)}」の列が見つかりません。")
                else:
                    if '印刷状態' in df_raw.columns:
                        df_raw.insert(0, "印刷", df_raw['印刷状態'] == '印刷対象')
                    else:
                        df_raw.insert(0, "印刷", True)
                    st.session_state.df_edited = df_raw
            except Exception as e:
                st.error(f"読み込みエラー: {e}")

        # リスト表示
        if st.session_state.df_edited is not None:
            edited_df = st.data_editor(
                st.session_state.df_edited,
                column_config={
                    "印刷": st.column_config.CheckboxColumn(
                        "印刷",
                        default=True,
                    )
                },
                hide_index=True,
                use_container_width=True,
                num_rows="fixed"
            )
            
            st.session_state.df_edited = edited_df
            target_df = edited_df[edited_df['印刷'] == True]
            st.write(f"🖨️ 現在の印刷対象: **{len(target_df)}** 件")

    # --- プレビュー選択 ---
    if st.session_state.df_edited is not None:
        st.markdown("---")
        st.subheader("2. 仕上がりプレビュー")
        
        current_df = st.session_state.df_edited
        preview_options = current_df.apply(lambda x: f"{'✅' if x['印刷'] else '⬜'} {x['名前']} ({str(x.get('住所',''))[:6]}...)", axis=1)
        
        selected_index = st.selectbox(
            "確認したい宛名を選択:",
            current_df.index,
            format_func=lambda i: preview_options[i]
        )

        # --- 右カラム：プレビュー ---
        with col2:
            st.subheader("🖼️ プレビュー")
            if selected_index is not None:
                record = current_df.iloc[selected_index]
                name = str(record["名前"])
                address = str(record["住所"])
                
                with st.spinner('プレビュー画像を生成中...'):
                    img = generate_preview_image(name, address)
                    st.image(img, caption=f"「{name}」様のイメージ", use_container_width=True)

    # --- PDF作成ボタン ---
    with col1:
        st.markdown("---")
        st.subheader("3. 印刷用PDFの作成")
        
        if uploaded_file is not None:
            if len(target_df) > 0:
                if st.button("PDFをダウンロード (選択した宛名のみ)", type="primary"):
                    with st.spinner('PDFを作成しています...'):
                        pdf_data = generate_nengajo_pdf(target_df.to_dict(orient="records"))
                        
                        st.download_button(
                            label="📥 PDFファイルを保存",
                            data=pdf_data,
                            file_name="nengajo_print.pdf",
                            mime="application/pdf"
                        )
                        st.success("作成完了！")
            else:
                st.warning("印刷する人が選択されていません。")