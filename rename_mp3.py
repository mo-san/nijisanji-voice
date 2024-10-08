from pathlib import Path
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
        return None

    parts = file_name[:-4].split('_')
    if len(parts) == 2:
        return {
            'Character': parts[0],
            'Suffix': parts[1],
            'IsEX': False
        }
    if len(parts) == 3 and parts[0] == '01':
        return {
            'Character': parts[1],
            'Suffix': parts[2],
            'IsEX': False
        }
    return None


def generate_new_file_name(parsed_name: Dict[str, object]) -> str:
    """新しいファイル名を生成する"""
    character = parsed_name['Character']
    suffix = parsed_name['Suffix']
    is_ex = parsed_name['IsEX']

    number = "02" if is_ex else "01"
    ex_suffix = " EX" if is_ex else ""
    return f"[{suffix}]{character} - {number} {suffix}{ex_suffix}.mp3"


def get_renamed_files(directory_path: Path, recursive: bool = False) -> list[tuple[Path, Path]]:
    """リネーム後のファイル名のリストを取得する"""
    renamed_files = []

    for file_path in directory_path.rglob("*.mp3") if recursive else directory_path.glob("*.mp3"):
        normalized_file_name = normalize_file_name(file_path.name)
        parsed_name = parse_file_name(normalized_file_name)

        if not parsed_name:
            print(f"Skipped: '{file_path}' - does not match expected pattern")
            continue

        new_name = generate_new_file_name(parsed_name)
        new_path = file_path.parent / new_name
        renamed_files.append((file_path, new_path))
    return renamed_files


def rename_files(path: str, renamed_files: List[Tuple[str, str]], root: tk.Tk, dry_run: bool = False) -> None:
    """実際にファイルをリネームする"""
    for old_name, new_name in renamed_files:
        old_path = Path(path) / old_name
        new_path = Path(path) / new_name
        if dry_run:
            print(f"Dry-run: Would rename {old_path} to {new_path}")
        else:
            old_path.rename(new_path)

    if not dry_run:
        messagebox.showinfo("完了", "ファイルのリネームが完了しました。")
    root.destroy()


def setup_preview_gui(root, directory_path, renamed_files, dry_run):
    """リネーム後のファイル名をGUIでプレビューするためのセットアップ"""
    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    # ウィンドウの大きさを1.5倍に設定
    default_width = 800
    default_height = 600
    window_width = int(default_width * 1.5)
    window_height = int(default_height * 1.5)
    root.geometry(f"{window_width}x{window_height}")

    tree = ttk.Treeview(frame, columns=('Old Name', 'New Name'), show='headings')
    tree.heading('Old Name', text='元のファイル名', command=lambda: sortby(tree, 'Old Name', False))
    tree.heading('New Name', text='新しいファイル名', command=lambda: sortby(tree, 'New Name', False))

    for old_name, new_name in renamed_files:
        tree.insert('', tk.END, values=(old_name, new_name))

    tree.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    # スクロールバーを追加
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky=tk.N + tk.S)

    # リサイズ設定
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)
    tree.columnconfigure(0, weight=1)
    tree.columnconfigure(1, weight=1)

    # 確認ボタンを追加
    button_frame = ttk.Frame(root, padding=10)
    button_frame.grid(row=1, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    confirm_button_text = "リネームを実行 (dry-run のため実際には書き込まれません)" if dry_run else "リネームを実行"
    confirm_button = ttk.Button(button_frame, text=confirm_button_text, command=lambda: rename_files(directory_path, renamed_files, root, dry_run))
    confirm_button.grid(row=0, column=0, padx=5, pady=5)

    cancel_button = ttk.Button(button_frame, text="キャンセル", command=root.destroy)
    cancel_button.grid(row=0, column=1, padx=5, pady=5)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)

    # Treeviewのセルを部分的にコピー可能にする
    tree.bind("<Double-1>", on_double_click)


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


def preview_renamed_files(directory_path: Path, recursive: bool, dry_run: bool) -> None:
    """リネーム後のファイル名をGUIでプレビューする"""
    renamed_files = get_renamed_files(directory_path, recursive)

    # GUIのセットアップ
    root = tk.Tk()
    root.title("ファイル名リネームプレビュー")

    setup_preview_gui(root, directory_path, renamed_files, dry_run)

    root.mainloop()


def main():
    parser = argparse.ArgumentParser(description="ファイル名をリネームするスクリプト")
    parser.add_argument("--directory", type=str, required=True, help="処理するディレクトリのパス")
    parser.add_argument("--recursive", action="store_true", help="指定するとサブディレクトリを再帰的に処理する")
    parser.add_argument("--dry-run", action="store_true", help="指定すると実際にはリネームせずに処理をシミュレートする")

    args = parser.parse_args()
    directory: str = args.directory
    recursive: bool = args.recursive
    dry_run: bool = args.dry_run

    preview_renamed_files(Path(directory), recursive, dry_run)


if __name__ == "__main__":
    main()
