from pathlib import Path
import unicodedata
from typing import Optional, Dict, List, Tuple
import argparse
import re

# tkinterのインポートを試行し、失敗した場合はCUIモードで動作
try:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False


def normalize_file_name(file_name: str) -> str:
    """NFKC正規化を行う"""
    return unicodedata.normalize("NFKC", file_name)


def is_already_properly_formatted(file_name: str) -> bool:
    """ファイル名がすでに適切にフォーマットされているかチェックする"""
    # 目標フォーマット: [アルバム名]アーティスト名 - 01/02 トラック名[ EX].mp3
    # 通常版: [アルバム名]アーティスト名 - 01 トラック名.mp3
    # EX版: [アルバム名]アーティスト名 - 02 トラック名 EX.mp3

    # 正規表現パターン
    pattern = r"^\[([^\]]+)\](.+?) - (01|02) \1( EX)?\.mp3$"

    match = re.match(pattern, file_name)
    if not match:
        return False

    album_name, artist_name, number, ex_suffix = match.groups()

    # EX版の場合は番号が02で EX サフィックスが必要
    if ex_suffix == " EX":
        return number == "02"
    # 通常版の場合は番号が01で EX サフィックスがない
    else:
        return number == "01"


def parse_file_name(file_name: str) -> Optional[Dict[str, object]]:
    """ファイル名を解析して必要な情報を抽出する"""
    # EX_アーティスト名_アルバム名/トラック名.mp3 形式
    if file_name.startswith("EX_"):
        parts = file_name[3:-4].split("_")
        if len(parts) == 2:
            return {
                "Character": parts[0],  # アーティスト名
                "Suffix": parts[1],  # アルバム名/トラック名
                "IsEX": True,
            }
        return None

    parts = file_name[:-4].split("_")
    # アーティスト名_アルバム名/トラック名.mp3 形式
    if len(parts) == 2:
        return {
            "Character": parts[0],  # アーティスト名
            "Suffix": parts[1],  # アルバム名/トラック名
            "IsEX": False,
        }
    # 01_アーティスト名_アルバム名/トラック名.mp3 形式
    if len(parts) == 3 and parts[0] == "01":
        return {
            "Character": parts[1],  # アーティスト名
            "Suffix": parts[2],  # アルバム名/トラック名
            "IsEX": False,
        }
    return None


def generate_new_file_name(parsed_name: Dict[str, object]) -> str:
    """新しいファイル名を生成する"""
    # Character = アーティスト名, Suffix = アルバム名/トラック名として扱う
    character = parsed_name["Character"]  # アーティスト名
    suffix = parsed_name["Suffix"]  # アルバム名/トラック名
    is_ex = parsed_name["IsEX"]

    number = "02" if is_ex else "01"
    ex_suffix = " EX" if is_ex else ""
    return f"[{suffix}]{character} - {number} {suffix}{ex_suffix}.mp3"


def get_renamed_files(
    directory_path: Path, recursive: bool = False
) -> list[tuple[Path, Path]]:
    """リネーム後のファイル名のリストを取得する"""
    renamed_files = []

    for file_path in (
        directory_path.rglob("*.mp3") if recursive else directory_path.glob("*.mp3")
    ):
        normalized_file_name = normalize_file_name(file_path.name)
        parsed_name = parse_file_name(normalized_file_name)

        if not parsed_name:
            # ファイル名の解析に失敗した場合、すでに適切にフォーマットされているかチェック
            if is_already_properly_formatted(normalized_file_name):
                print(f"Already formatted: '{file_path}' - すでにリネーム済み")
            else:
                print(f"Skipped: '{file_path}' - does not match expected pattern")
            continue

        new_name = generate_new_file_name(parsed_name)
        new_path = file_path.parent / new_name

        # 新しいファイル名が現在のファイル名と同じ場合もすでにリネーム済みとして扱う
        if new_path.name == file_path.name:
            print(f"Already formatted: '{file_path}' - すでにリネーム済み")
            continue

        renamed_files.append((file_path, new_path))
    return renamed_files


def rename_files_gui(
    path: str, renamed_files: List[Tuple[str, str]], root, dry_run: bool = False
) -> None:
    """GUIモードでファイルをリネームする"""
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


def rename_files_cui(
    directory_path: Path, renamed_files: List[Tuple[Path, Path]], dry_run: bool = False
) -> None:
    """CUIモードでファイルをリネームする"""
    for old_path, new_path in renamed_files:
        if dry_run:
            print(f"Dry-run: Would rename {old_path} to {new_path}")
        else:
            old_path.rename(new_path)

    if not dry_run:
        print("ファイルのリネームが完了しました。")


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

    tree = ttk.Treeview(frame, columns=("Old Name", "New Name"), show="headings")
    tree.heading(
        "Old Name",
        text="元のファイル名",
        command=lambda: sortby(tree, "Old Name", False),
    )
    tree.heading(
        "New Name",
        text="新しいファイル名",
        command=lambda: sortby(tree, "New Name", False),
    )

    for old_name, new_name in renamed_files:
        tree.insert("", tk.END, values=(old_name, new_name))

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

    confirm_button_text = (
        "リネームを実行 (dry-run のため実際には書き込まれません)"
        if dry_run
        else "リネームを実行"
    )
    confirm_button = ttk.Button(
        button_frame,
        text=confirm_button_text,
        command=lambda: rename_files_gui(directory_path, renamed_files, root, dry_run),
    )
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
    data = [(tree.set(child, col), child) for child in tree.get_children("")]
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], "", ix)
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


def display_cui_table(renamed_files: List[Tuple[Path, Path]], dry_run: bool) -> None:
    """CUIでリネーム後のファイル名を表形式で表示する"""
    if not renamed_files:
        print("リネーム対象のファイルが見つかりませんでした。")
        return

    print("\n" + "=" * 80)
    print("ファイル名リネームプレビュー")
    print("=" * 80)

    # ヘッダー表示
    print(f"{'元のファイル名':<40} {'新しいファイル名':<40}")
    print("-" * 80)

    # ファイル一覧表示
    for old_path, new_path in renamed_files:
        old_name = old_path.name
        new_name = new_path.name
        print(f"{old_name:<40} {new_name:<40}")

    print("-" * 80)
    print(f"合計: {len(renamed_files)}ファイル")
    print("=" * 80)


def preview_renamed_files_cui(
    directory_path: Path, recursive: bool, dry_run: bool
) -> None:
    """CUIモードでリネーム後のファイル名をプレビューする"""
    renamed_files = get_renamed_files(directory_path, recursive)

    display_cui_table(renamed_files, dry_run)

    if not renamed_files:
        return

    # ユーザーに確認
    action_text = (
        "実行しますか？(dry-run のため実際には書き込まれません)"
        if dry_run
        else "リネームを実行しますか？"
    )
    while True:
        response = input(f"\n{action_text} [y/N]: ").strip().lower()
        if response in ["y", "yes"]:
            rename_files_cui(directory_path, renamed_files, dry_run)
            break
        elif response in ["n", "no", ""]:
            print("キャンセルしました。")
            break
        else:
            print("y または n で答えてください。")


def preview_renamed_files(directory_path: Path, recursive: bool, dry_run: bool) -> None:
    """リネーム後のファイル名をプレビューする（GUI/CUI自動切り替え）"""
    if GUI_AVAILABLE:
        # GUIモード
        renamed_files = get_renamed_files(directory_path, recursive)
        root = tk.Tk()
        root.title("ファイル名リネームプレビュー")
        setup_preview_gui(root, directory_path, renamed_files, dry_run)
        root.mainloop()
    else:
        # CUIモード
        print("GUIが利用できないため、CUIモードで実行します。")
        preview_renamed_files_cui(directory_path, recursive, dry_run)


def main():
    parser = argparse.ArgumentParser(description="ファイル名をリネームするスクリプト")
    parser.add_argument(
        "--directory", type=str, required=True, help="処理するディレクトリのパス"
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="指定するとサブディレクトリを再帰的に処理する",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="指定すると実際にはリネームせずに処理をシミュレートする",
    )
    parser.add_argument(
        "--cui", action="store_true", help="指定するとGUIではなくCUIモードで実行する"
    )

    args = parser.parse_args()
    directory: str = args.directory
    recursive: bool = args.recursive
    dry_run: bool = args.dry_run
    force_cui: bool = args.cui

    if force_cui:
        # CUIモードを強制
        preview_renamed_files_cui(Path(directory), recursive, dry_run)
    else:
        # 通常の自動判定
        preview_renamed_files(Path(directory), recursive, dry_run)


if __name__ == "__main__":
    main()
