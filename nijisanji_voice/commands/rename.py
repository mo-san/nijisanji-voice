from pathlib import Path
from typing import Optional, Dict, List, Tuple

from ..lib.cli_utils import (
    get_user_confirmation,
    display_completion_message,
    display_cui_table_header,
    display_cui_table_footer,
    setup_common_argument_parser,
    print_no_files_message,
    natural_sort_key,
    pad_text,
)
from ..lib.gui_utils import (
    GUI_AVAILABLE,
    sortby,
    on_double_click,
    setup_window_size,
    setup_scrollbar_for_tree,
    setup_button_frame,
    setup_frame_resize_behavior,
)

# 共通モジュールのインポート
from ..lib.mp3_utils import (
    normalize_file_name,
    get_mp3_files,
    handle_file_parsing_failure,
)

# tkinterの再インポート（GUIモードで使用）
if GUI_AVAILABLE:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox


def parse_file_name(file_name: str) -> Optional[Dict[str, object]]:
    """ファイル名を解析して必要な情報を抽出する"""
    # EX Another_アーティスト名_アルバム名/トラック名.mp3 形式
    if file_name.startswith("EX Another_"):
        parts = file_name[11:-4].split("_")
        if len(parts) == 2:
            return {
                "Character": parts[0],  # アーティスト名
                "Suffix": parts[1],  # アルバム名/トラック名
                "IsEX": True,
                "IsAnother": True,
            }
        return None
    
    # EX_アーティスト名_アルバム名/トラック名.mp3 形式
    if file_name.startswith("EX_"):
        parts = file_name[3:-4].split("_")
        if len(parts) == 2:
            return {
                "Character": parts[0],  # アーティスト名
                "Suffix": parts[1],  # アルバム名/トラック名
                "IsEX": True,
                "IsAnother": False,
            }
        return None

    parts = file_name[:-4].split("_")
    # アーティスト名_アルバム名/トラック名.mp3 形式
    if len(parts) == 2:
        return {
            "Character": parts[0],  # アーティスト名
            "Suffix": parts[1],  # アルバム名/トラック名
            "IsEX": False,
            "IsAnother": False,
        }
    # 01_アーティスト名_アルバム名/トラック名.mp3 形式 (track numbers 01-06)
    if len(parts) == 3 and parts[0].isdigit() and parts[0] in ["01", "02", "03", "04", "05", "06"]:
        return {
            "Character": parts[1],  # アーティスト名
            "Suffix": parts[2],  # アルバム名/トラック名
            "IsEX": False,
            "IsAnother": False,
            "TrackNumber": parts[0],  # トラック番号
        }
    return None


def generate_new_file_name(parsed_name: Dict[str, object]) -> str:
    """新しいファイル名を生成する"""
    # Character = アーティスト名, Suffix = アルバム名/トラック名として扱う
    character = parsed_name["Character"]  # アーティスト名
    suffix = parsed_name["Suffix"]  # アルバム名/トラック名
    is_ex = parsed_name["IsEX"]
    is_another = parsed_name.get("IsAnother", False)

    # TrackNumberが指定されている場合はそれを使用
    if "TrackNumber" in parsed_name:
        number = parsed_name["TrackNumber"]
    else:
        # EX Anotherの場合は03、EXの場合は02、それ以外は01
        if is_another:
            number = "03"
        elif is_ex:
            number = "02"
        else:
            number = "01"
    
    # サフィックス生成: EX Anotherの場合は " EX(Another)"、EXの場合は " EX"、それ以外は ""
    if is_another:
        ex_suffix = " EX(Another)"
    elif is_ex:
        ex_suffix = " EX"
    else:
        ex_suffix = ""
    
    return f"[{suffix}]{character} - {number} {suffix}{ex_suffix}.mp3"


def get_renamed_files(directory_path: Path, recursive: bool = False) -> list[tuple[Path, Path]]:
    """リネーム後のファイル名のリストを取得する"""
    renamed_files = []

    for file_path in get_mp3_files(directory_path, recursive):
        normalized_file_name = normalize_file_name(file_path.name)
        parsed_name = parse_file_name(normalized_file_name)

        if not parsed_name:
            handle_file_parsing_failure(file_path, normalized_file_name)
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

    display_completion_message(dry_run, "ファイルのリネーム")


def setup_preview_gui(root, directory_path, renamed_files, dry_run):
    """リネーム後のファイル名をGUIでプレビューするためのセットアップ"""
    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    # ウィンドウサイズを設定
    setup_window_size(root)

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
    setup_scrollbar_for_tree(frame, tree)

    # リサイズ設定
    setup_frame_resize_behavior(frame, tree)

    # ボタンを追加
    confirm_button_text = (
        "リネームを実行 (dry-run のため実際には書き込まれません)" if dry_run else "リネームを実行"
    )
    buttons_config = [
        (
            confirm_button_text,
            lambda: rename_files_gui(directory_path, renamed_files, root, dry_run),
        ),
        ("キャンセル", root.destroy),
    ]
    setup_button_frame(root, 1, buttons_config)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)

    # Treeviewのセルを部分的にコピー可能にする
    tree.bind("<Double-1>", on_double_click)


def display_cui_table(renamed_files: List[Tuple[Path, Path]], dry_run: bool) -> None:
    """CUIでリネーム後のファイル名を表形式で表示する"""
    if not renamed_files:
        print_no_files_message("リネーム対象")
        return

    display_cui_table_header("ファイル名リネームプレビュー")

    # 自然順でソート
    sorted_files = sorted(renamed_files, key=lambda x: natural_sort_key(str(x[0])))

    # カラム幅を定義
    old_name_width = 40
    new_name_width = 40

    # ヘッダー表示
    old_header = pad_text("元のファイル名", old_name_width)
    new_header = pad_text("新しいファイル名", new_name_width)
    print(f"{old_header} {new_header}")
    print("-" * 80)

    # ファイル一覧表示
    for old_path, new_path in sorted_files:
        old_name = pad_text(old_path.name, old_name_width)
        new_name = pad_text(new_path.name, new_name_width)
        print(f"{old_name} {new_name}")

    display_cui_table_footer(len(renamed_files))


def preview_renamed_files_cui(directory_path: Path, recursive: bool, dry_run: bool) -> None:
    """CUIモードでリネーム後のファイル名をプレビューする"""
    renamed_files = get_renamed_files(directory_path, recursive)

    display_cui_table(renamed_files, dry_run)

    if not renamed_files:
        return

    # ユーザーに確認
    if get_user_confirmation("リネーム", dry_run):
        rename_files_cui(directory_path, renamed_files, dry_run)
    else:
        print("キャンセルしました。")


def preview_renamed_files(directory_path: Path, recursive: bool, dry_run: bool) -> None:
    """リネーム後のファイル名をプレビューする（GUI/CUI自動切り替え）"""
    if GUI_AVAILABLE:
        # GUIモード
        renamed_files = get_renamed_files(directory_path, recursive)
        # 対象ファイルが0件の場合はCUIと同じメッセージを表示して正常終了
        if not renamed_files:
            print_no_files_message("リネーム対象")
            return
        root = tk.Tk()
        root.title("ファイル名リネームプレビュー")
        setup_preview_gui(root, directory_path, renamed_files, dry_run)
        root.mainloop()
    else:
        # CUIモード
        print("GUIが利用できないため、CUIモードで実行します。")
        preview_renamed_files_cui(directory_path, recursive, dry_run)


def main():
    parser = setup_common_argument_parser("ファイル名をリネームするスクリプト")

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
