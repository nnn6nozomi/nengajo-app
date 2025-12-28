import pandas as pd
import shutil
import os
import datetime
from openpyxl import load_workbook

class DataManager:
    REQUIRED_COLUMNS = ["NO.", "名前", "印刷状態", "グループ", "住所"]
    VALID_STATUS = ["未印刷", "印刷対象", "印刷済", "除外"]

    def __init__(self):
        self.df = None
        self.file_path = None

    def load_excel(self, file_object):
        """Excelファイルを読み込み、バリデーションを行う"""
        try:
            # StreamlitのUploadedFileオブジェクトから読み込む
            self.df = pd.read_excel(file_object, dtype={"NO.": str, "住所": str})
            
            # 1. カラムチェック
            missing_cols = [col for col in self.REQUIRED_COLUMNS if col not in self.df.columns]
            if missing_cols:
                return False, f"必須列が不足しています: {', '.join(missing_cols)}"

            # 2. NO. 重複チェック
            if self.df["NO."].duplicated().any():
                duplicated = self.df[self.df["NO."].duplicated()]["NO."].tolist()
                return False, f"NO. に重複があります: {duplicated}"

            # 3. 印刷状態チェック
            invalid_status = self.df[~self.df["印刷状態"].isin(self.VALID_STATUS)]
            if not invalid_status.empty:
                return False, f"不正な印刷状態が含まれています (行: {invalid_status.index.tolist()})"

            # ファイルパスの保持（保存用、UploadedFileにはname属性がある）
            self.file_path = file_object.name
            return True, "読み込み成功"

        except Exception as e:
            return False, f"読み込みエラー: {str(e)}"

    def save_excel(self, df_to_save, original_file_path):
        """データを保存し、バックアップを作成する"""
        if original_file_path is None:
            return False, "保存先ファイルが指定されていません。"

        # バックアップ作成
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{os.path.splitext(original_file_path)[0]}_backup_{timestamp}.xlsx"
        
        try:
            # 元ファイルが存在する場合のみバックアップ（初回アップロード直後はローカルパスの扱いが難しいが、
            # ここではローカルアプリ想定で、同名ファイルがカレントにあればバックアップする挙動とする）
            if os.path.exists(original_file_path):
                shutil.copy(original_file_path, backup_path)
            
            # 保存実行
            # Excelを開いたままだとPermissionErrorになる
            df_to_save.to_excel(original_file_path, index=False)
            return True, f"保存完了（バックアップ: {backup_path}）"
            
        except PermissionError:
            return False, "ファイルが開かれているため保存できません。Excelを閉じて再試行してください。"
        except Exception as e:
            return False, f"保存エラー: {str(e)}"

    def get_filtered_data(self, groups, statuses, search_name):
        """フィルタリングとソートを行う"""
        if self.df is None:
            return pd.DataFrame()
        
        temp_df = self.df.copy()

        # フィルタ
        if groups:
            temp_df = temp_df[temp_df["グループ"].isin(groups)]
        if statuses:
            temp_df = temp_df[temp_df["印刷状態"].isin(statuses)]
        if search_name:
            temp_df = temp_df[temp_df["名前"].astype(str).str.contains(search_name, na=False)]

        return temp_df