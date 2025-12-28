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
    st.error(f"⚠️ フォント「{TARGET_FONT_NAME}」が見つかりません（標準フォントになります）")
    st.info(f"💡 ヒント: 筆文字にするには、フォントファイルを「{TARGET_FONT_NAME}」という名前に変更して置いてください。")

st.markdown("---")

col1, col2 = st.columns([1.5, 1])

if 'df_edited' not in st.session_state:
    st.session_state.df_edited = None

# ==========================================
# 👈 左カラム：データ編集
# ==========================================
with col1:
    st.subheader("1. 住所録データの読み込み")

    # ▼▼▼ ガイドエリア（複数人の連名対応版） ▼▼▼
    with st.expander("📌 データ作成ガイド（AIプロンプトあり）"):
        st.markdown("### 1. 自分でExcelを作る場合")
        st.markdown("""
        以下の列を作ってください（「連名」は空欄でもOK）。
        * **名前**: 世帯主のお名前（例：山田 太郎）
        * **連名**: 2人目以降のお名前。
            * 夫婦など複数いる場合は**スペースで区切って**ください（例：「花子 一郎」）。
            * 名字が同じなら、下のお名前だけでOKです。
        * **住所**: 郵便番号込みの住所（例：100-0001 東京都...）
        """)
        
        st.markdown("---")
        
        st.markdown("### 2. AI(ChatGPT)に作らせる場合")
        st.write("以下の指示文をコピーして、ChatGPTに住所録テキストと一緒に貼り付けてください。")
        
        st.code("""
あなたはデータ処理のプロフェッショナルです。
添付のデータ（住所録）を、私が作成した年賀状アプリで読み込める形式に変換・補完してください。

## 現状の課題
1. **郵便番号の欠落**: 住所はあるが郵便番号がない行があり、このままではアプリではがきへの印字ができません。
2. **連名の処理**: 元データでは家族の名前がどうなっているか確認し（例：「備考」にある、または「名前」に併記されているなど）、アプリ用に分離する必要があります。

## 依頼内容
以下の手順でデータを加工し、新しいExcelファイルを出力してください。

### 手順1. 郵便番号の補完
「住所」列の情報を元に、正しい**7桁の郵便番号（ハイフンあり）**を特定してください。
※番地まで特定できない場合は、その市町村の代表番号など、推測できる範囲で埋めてください。

### 手順2. アプリ用フォーマットへの整形
以下の6列を持つテーブルを作成してください（順序厳守）。

| 列順 | 列名 | 内容・加工ルール |
| :--- | :--- | :--- |
| A | NO. | 元の「No.」があれば使用。なければ連番を振る。 |
| B | 名前 | **世帯主（1人目）の氏名**のみを記載してください。 |
| C | 連名 | 2人目以降の名前（配偶者やお子様）を記載してください。<br>・名字が同じ場合は**「下のお名前のみ」**にしてください。<br>・複数人いる場合は**スペース区切り**で横に並べてください（例：「花子 一郎」）。<br>・いない場合は空欄にしてください。 |
| D | 印刷状態 | 「印刷する」→「印刷対象」、「空欄」→「未印刷」に置換。<br>既に「印刷済」「除外」のものは維持してください。 |
| E | グループ | 元の「グループ」情報をそのまま転記してください。 |
| F | 住所 | **重要：** 手順1で特定した郵便番号を、住所の先頭に結合してください。<br>例: `595-0031 大阪府泉大津市...`<br>（アプリがこの形式から郵便番号を自動抽出します） |

### 手順3. 出力
- シート名は「名簿」にしてください。
- ファイル形式は `.xlsx` で出力してください。

処理をお願いします。
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
                    # 印刷列の追加
                    if '印刷状態' in df_raw.columns:
                        df_raw.insert(0, "印刷", df_raw['印刷状態'] == '印刷対象')
                    else:
                        df_raw.insert(0, "印刷", True)
                    
                    # 連名列がない場合は空で作っておく（エラー回避）
                    if '連名' not in df_raw.columns:
                        df_raw['連名'] = ""

                    st.session_state.df_edited = df_raw
            except Exception as e:
                st.error(f"読み込みエラー: {e}")

        # リスト表示
        if st.session_state.df_edited is not None:
            # 編集用テーブルの表示
            edited_df = st.data_editor(
                st.session_state.df_edited,
                column_config={
                    "印刷": st.column_config.CheckboxColumn("印刷", default=True),
                    "連名": st.column_config.TextColumn("連名", help="複数人の場合はスペースで区切ってください（例：花子 一郎）")
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
        
        # セレクトボックスの表示名を作る（連名があれば表示）
        def make_label(row):
            renmei_txt = str(row.get('連名',''))
            if renmei_txt == "nan": renmei_txt = ""
            renmei_part = f"・{renmei_txt}" if renmei_txt else ""
            return f"{'✅' if row['印刷'] else '⬜'} {row['名前']}{renmei_part}"

        preview_options = current_df.apply(make_label, axis=1)
        
        selected_index = st.selectbox(
            "確認したい宛名を選択:",
            current_df.index,
            format_func=lambda i: preview_options[i]
        )

        # ==========================================
        # 👉 右カラム：プレビュー画面
        # ==========================================
        with col2:
            st.subheader("🖼️ プレビュー")
            if selected_index is not None:
                record = current_df.iloc[selected_index]
                name = str(record["名前"])
                address = str(record["住所"])
                
                # 連名を取得（なければ空文字）
                renmei = str(record.get("連名", ""))
                if renmei == "nan": renmei = ""
                
                with st.spinner('プレビュー画像を生成中...'):
                    # 連名データも渡して画像生成
                    img = generate_preview_image(name, address, renmei)
                    st.image(img, caption=f"「{name}」様のイメージ", use_container_width=True)

    # ==========================================
    # 📥 PDFダウンロードボタン
    # ==========================================
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