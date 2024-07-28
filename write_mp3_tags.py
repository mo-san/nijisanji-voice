import os
import argparse
from pathlib import Path
import unicodedata
from typing import Optional, TypedDict, List, Tuple
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3NoHeaderError


class ID3Tags(TypedDict):
    """ID3タグの情報を保持する型定義"""
    track_name: str
    artist_name: str
    album_name: str
    track_number: int


def normalize_file_name(file_name: str) -> str:
    """NFKC正規化を行う"""
    return unicodedata.normalize('NFKC', file_name)


def parse_file_name(file_name: str) -> Optional[ID3Tags]:
    """
    ファイル名を解析してID3タグの情報を抽出する

    期待されるファイル名のパターン:
    [カテゴリ]キャラクター名 - 番号 タイトル (EX).mp3

    例:
    [ジューンブライド2024ボイス]ソフィア・ヴァレンタイン - 01 ジューンブライド2024ボイス.mp3
    [夜更かしボイス]五十嵐梨花 - 02 夜更かしボイス EX.mp3
    """
    if not file_name.endswith('.mp3'):
        return None

    base_name = file_name[:-4]
    parts = base_name.split(' - ')
    if len(parts) != 2:
        return None

    prefix_part, track_part = parts
    if not prefix_part.startswith('[') or ']' not in prefix_part:
        return None

    album_name = prefix_part[1:prefix_part.index(']')]
    artist_name = prefix_part[prefix_part.index(']') + 1:]

    track_info = track_part.split(' ', 1)
    if len(track_info) != 2:
        return None

    try:
        track_number = int(track_info[0])
    except ValueError:
        return None

    track_name = track_info[1].replace(" EX", "")
    return ID3Tags(
        track_name=track_name,
        artist_name=artist_name,
        album_name=album_name,
        track_number=track_number
    )


def write_id3_tags(file_path: str, tags: ID3Tags) -> None:
    """ID3タグを書き込む"""
    try:
        audio = EasyID3(file_path)
    except ID3NoHeaderError:
        audio = EasyID3()
        audio.save(file_path)

    audio['title'] = tags['track_name']
    audio['artist'] = tags['artist_name']
    audio['album'] = tags['album_name']
    audio['tracknumber'] = str(tags['track_number'])
    audio.save(file_path)
    print(f"Processed: {file_path}")


def process_files(path: str) -> List[Tuple[str, ID3Tags]]:
    """ディレクトリ内のファイルにID3タグを書き込む"""
    processed_files = []
    for file_path in Path(path).rglob('*.mp3'):
        tags = parse_file_name(file_path.name)

        if not tags:
            print(f"Skipped: {file_path} - does not match expected pattern")
            continue

        processed_files.append((file_path, tags))
    return processed_files


def preview_id3_tags(processed_files: List[Tuple[str, ID3Tags]]) -> None:
    """ID3タグのプレビューをGUIで表示する"""

    def execute_writes():
        for i, (file_path, tags) in enumerate(processed_files, 1):
            write_id3_tags(file_path, tags)
            progress_var.set(i)
            current_file_var.set(f"Processing: {os.path.basename(file_path)}")
            root.update_idletasks()
        messagebox.showinfo("完了", "ID3タグの書き込みが完了しました。")
        root.destroy()

    # GUIのセットアップ
    root = tk.Tk()
    root.title("ID3タグプレビュー")

    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # ウィンドウの大きさを1.5倍に設定
    default_width = 800
    default_height = 600
    window_width = int(default_width * 1.5)
    window_height = int(default_height * 1.5)
    root.geometry(f"{window_width}x{window_height}")

    tree = ttk.Treeview(frame, columns=('File Path', 'Title', 'Artist', 'Album', 'Track Number'), show='headings')
    tree.heading('File Path', text='ファイルパス', command=lambda: sortby(tree, 'File Path', False))
    tree.heading('Title', text='タイトル', command=lambda: sortby(tree, 'Title', False))
    tree.heading('Artist', text='アーティスト', command=lambda: sortby(tree, 'Artist', False))
    tree.heading('Album', text='アルバム', command=lambda: sortby(tree, 'Album', False))
    tree.heading('Track Number', text='トラック番号', command=lambda: sortby(tree, 'Track Number', False))

    for file_path, tags in processed_files:
        tree.insert('', tk.END, values=(
            file_path, tags['track_name'], tags['artist_name'], tags['album_name'], tags['track_number']))

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

    # 列の幅を調整
    tree.column('File Path', width=300)
    tree.column('Title', width=150)
    tree.column('Artist', width=100)
    tree.column('Album', width=100)
    tree.column('Track Number', width=10)

    # 進捗バーを追加
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(processed_files))
    progress_bar.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

    current_file_var = tk.StringVar()
    current_file_label = ttk.Label(root, textvariable=current_file_var)
    current_file_label.grid(row=2, column=0, padx=10, pady=5, sticky=(tk.W, tk.E))

    # ボタンを追加
    button_frame = ttk.Frame(root, padding=10)
    button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    execute_button = ttk.Button(button_frame, text="タグ書き込みを実行", command=execute_writes)
    execute_button.grid(row=0, column=0, padx=5, pady=5)

    cancel_button = ttk.Button(button_frame, text="キャンセル", command=root.destroy)
    cancel_button.grid(row=0, column=1, padx=5, pady=5)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)
    root.rowconfigure(2, weight=0)
    root.rowconfigure(3, weight=0)

    # Treeviewのセルを部分的にコピー可能にする
    tree.bind("<Double-1>", on_double_click)

    root.mainloop()

def sortby(tree, col, descending):
    """Treeviewの並べ替えを行う"""
    data = [(tree.set(child, col), child) for child in tree.get_children('')]
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], '', ix)
    tree.heading(col, command=lambda: sortby(tree, col, int(not descending)))

def on_double_click(event):
    """セルをダブルクリックで部分選択とコピーを可能にする"""
    item_id = event.widget.identify_row(event.y)
    column = event.widget.identify_column(event.x)
    value = event.widget.item(item_id, "values")[int(column[1:]) - 1]
    show_copy_popup(value)

def show_copy_popup(value):
    """部分選択とコピーのためのポップアップを表示"""
    popup = tk.Toplevel()
    popup.title("部分選択とコピー")
    text = tk.Text(popup, wrap="word")
    text.insert("1.0", value)
    text.pack(expand=True, fill="both")
    text.bind("<Control-c>", lambda e: popup.clipboard_append(text.selection_get()))
    close_button = ttk.Button(popup, text="閉じる", command=popup.destroy)
    close_button.pack()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MP3ファイルにID3タグを付けるスクリプト")
    parser.add_argument("--directory", type=str, required=True, help="処理するディレクトリのパス")

    args = parser.parse_args()
    directory: str = args.directory

    files_to_process = process_files(directory)

    preview_id3_tags(files_to_process)
