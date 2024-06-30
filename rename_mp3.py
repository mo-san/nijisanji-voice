import os
import unicodedata
from typing import Optional, Dict, List, Tuple
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import argparse


def normalize_file_name(file_name: str) -> str:
    """NFKC正規化を行う"""
    return unicodedata.normalize('NFKC', file_name)


def parse_file_name(file_name: str) -> Optional[Dict[str, object]]:
    """ファイル名を解析して必要な情報を抽出する"""
    if file_name.startswith('EX_'):
        parts = file_name[3:-4].split('_')
        if len(parts) == 2:
            return {
                'Character': parts[0],
                'Suffix': parts[1],
                'IsEX': True
            }
    else:
        parts = file_name[:-4].split('_')
        if len(parts) == 2:
            return {
                'Character': parts[0],
                'Suffix': parts[1],
                'IsEX': False
            }
    return None


def generate_new_file_name(character: str, suffix: str, is_ex: bool) -> str:
    """新しいファイル名を生成する"""
    number = "02" if is_ex else "01"
    ex_suffix = " EX" if is_ex else ""
    return f"[{suffix}]{character} - {number} {suffix}{ex_suffix}.mp3"


def get_renamed_files(path: str) -> List[Tuple[str, str]]:
    """リネーム後のファイル名のリストを取得する"""
    renamed_files = []
    for file_name in os.listdir(path):
        if file_name.endswith('.mp3'):
            normalized_file_name = normalize_file_name(file_name)
            parsed_name = parse_file_name(normalized_file_name)

            if not parsed_name:
                print(f"Skipped: '{file_name}' - does not match expected pattern")
                continue

            new_name = generate_new_file_name(parsed_name['Character'], parsed_name['Suffix'], parsed_name['IsEX'])
            renamed_files.append((file_name, new_name))
    return renamed_files


def rename_files(path: str, renamed_files: List[Tuple[str, str]], root: tk.Tk) -> None:
    """実際にファイルをリネームする"""
    for old_name, new_name in renamed_files:
        old_path = os.path.join(path, old_name)
        new_path = os.path.join(path, new_name)
        os.rename(old_path, new_path)
    messagebox.showinfo("完了", "ファイルのリネームが完了しました。")
    root.destroy()


def preview_renamed_files(path: str) -> None:
    """リネーム後のファイル名をGUIでプレビューする"""
    renamed_files = get_renamed_files(path)

    # GUIのセットアップ
    root = tk.Tk()
    root.title("ファイル名リネームプレビュー")

    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    tree = ttk.Treeview(frame, columns=('Old Name', 'New Name'), show='headings')
    tree.heading('Old Name', text='元のファイル名')
    tree.heading('New Name', text='新しいファイル名')

    for old_name, new_name in renamed_files:
        tree.insert('', tk.END, values=(old_name, new_name))

    tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # スクロールバーを追加
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

    # リサイズ設定
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)
    tree.columnconfigure(0, weight=1)
    tree.columnconfigure(1, weight=1)

    # 確認ボタンを追加
    button_frame = ttk.Frame(root, padding=10)
    button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    confirm_button = ttk.Button(button_frame, text="リネームを実行",
                                command=lambda: rename_files(path, renamed_files, root))
    confirm_button.grid(row=0, column=0, padx=5, pady=5)

    cancel_button = ttk.Button(button_frame, text="キャンセル", command=root.destroy)
    cancel_button.grid(row=0, column=1, padx=5, pady=5)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)

    root.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ファイル名をリネームするスクリプト")
    parser.add_argument("--directory", type=str, required=True, help="処理するディレクトリのパス")

    args = parser.parse_args()
    directory: str = args.directory

    preview_renamed_files(directory)
