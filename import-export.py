# -*- coding: utf-8 -*-
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta

def format_elapsed_time(minutes):
    """分を日、時、分の形式にフォーマット"""
    if pd.isna(minutes) or minutes <= 0:
        return "0日 00:00"
    
    days = minutes // 1440
    hours = (minutes % 1440) // 60
    remaining_minutes = minutes % 60
    return f"{int(days)}日 {int(hours):02}:{int(remaining_minutes):02}"

# CSVファイルをインポートして処理する関数
def import_csv():
    # ファイル選択ダイアログの表示
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="CSVファイルを選択してください", filetypes=[("CSV files", "*.csv")])
    
    if not file_path:
        print("キャンセルされました。")
        return

    # CSVファイルを読み込む
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"CSVファイルの読み込み中にエラーが発生しました: {e}")
        return

    # 必要な列が存在するかチェック
    required_columns = ["Reg.Time", "End Time", "Compl.Time"]
    for col in required_columns:
        if col not in df.columns:
            print(f"列 '{col}' がCSVファイルにありません。")
            return

    # 所要時間とステータスを計算して新しい列に追加
    elapsed_time_list = []
    completion_status_list = []
    
    for _, row in df.iterrows():
        reg_time = pd.to_datetime(row["Reg.Time"], errors='coerce')
        compl_time = pd.to_datetime(row["Compl.Time"], errors='coerce')

        if pd.notna(reg_time) and pd.notna(compl_time):
            elapsed_time = compl_time - reg_time
            total_minutes = elapsed_time.total_seconds() / 60
            days = total_minutes // 1440
            hours = (total_minutes % 1440) // 60
            minutes = total_minutes % 60
            elapsed_time_list.append(f"{int(days)}日 {int(hours):02}:{int(minutes):02}")
            completion_status_list.append("完了")
        else:
            elapsed_time_list.append("")  # 未完了の場合のデフォルト値
            completion_status_list.append("未完了")
    
    # 新しい列をデータフレームに追加
    df["Elapsed Time"] = elapsed_time_list
    df["Completion"] = completion_status_list

    # 必要のない列を削除
    df = df.drop(columns=["End Time", "Compl.Time"])  # 必要なら削除対象を調整

    # 処理済みのデータを保存
    processed_path = os.path.join(os.environ["USERPROFILE"], "Downloads", "processed_data.csv")
    try:
        df.to_csv(processed_path, index=False, encoding="utf-8-sig")
        print(f"処理済みデータが保存されました: {processed_path}")
    except Exception as e:
        print(f"データの保存中にエラーが発生しました: {e}")

# データをエクスポートする関数
def export_csv():
    # インポートしたデータを処理済みフォルダから読み込む
    processed_path = os.path.join(os.environ["USERPROFILE"], "Downloads", "processed_data.csv")
    if not os.path.exists(processed_path):
        print("処理済みデータが見つかりません。最初にインポートしてください。")
        return

    df = pd.read_csv(processed_path)

    # エクスポート先を指定
    export_path = os.path.join(os.environ["USERPROFILE"], "Downloads", "export.csv")
    df.to_csv(export_path, index=False, encoding="shift_jis")
    print(f"データがエクスポートされました: {export_path}")

if __name__ == "__main__":
    while True:
        print("1: CSVをインポートして処理する\n2: データをエクスポートする\n3: 終了")
        choice = input("選択してください: ")
        if choice == "1":
            import_csv()
        elif choice == "2":
            export_csv()
        elif choice == "3":
            print("終了します。")
            break
        else:
            print("無効な選択です。再試行してください。")


