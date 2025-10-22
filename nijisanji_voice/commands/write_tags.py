from pathlib import Path
from typing import Optional, List, Tuple
from mutagen.easyid3 import EasyID3
from mutagen.id3._util import ID3NoHeaderError

from ..lib.cli_utils import (
    get_user_confirmation,
    display_completion_message,
    display_cui_table_header,
    display_cui_table_footer,
    setup_common_argument_parser,
    print_no_files_message,
    truncate_text,
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
    ID3Tags,
    normalize_file_name,
    get_mp3_files,
    handle_file_parsing_failure,
)

# tkinterの再インポート（GUIモードで使用）
if GUI_AVAILABLE:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox


def extract_track_info(track_part: str, artist_name: str, album_name: str) -> Tuple[int, str]:
    """
    トラック番号とトラック名を抽出する

    期待されるトラック名のパターン:
    番号 タイトル
    または
    タイトル

    例:
    01 サンプルボイス
    XXXボイス EX
    XXXボイス EX(Another)

    Args:
        track_part (str): トラック名の部分
        artist_name (str): アーティスト名
        album_name (str): アルバム名

    Returns:
        Tuple[int, str]: トラック番号とトラック名
    """
    track_info = track_part.split(" ", 1)

    if len(track_info) != 2:
        return 1, f"[{album_name}] {artist_name}"

    try:
        track_number = int(track_info[0])
        track_name = track_info[1]
    except ValueError:
        track_number = 1
        track_name = track_part

    # EX(Another)の処理
    if track_name.endswith(" EX(Another)"):
        track_name = track_name[:-12] + f" [{album_name}] {artist_name} EX(Another)"
    # EXの処理
    elif track_name.endswith(" EX"):
        track_name = track_name[:-3] + f" [{album_name}] {artist_name} EX"
    else:
        track_name = f"[{album_name}] {artist_name}"

    return track_number, track_name


def parse_file_name(file_name: str) -> Optional[ID3Tags]:
    """
    ファイル名を解析してID3タグの情報を抽出する

    期待されるファイル名のパターン:
    [アルバム名]アーティスト名 - 番号 トラック名 (EX).mp3
    または
    [アルバム名]アーティスト名 - トラック名.mp3

    先頭に (YYYY-MM) の日付が付く場合も許容し、その場合はアルバム名に含める。

    例:
    [XXXボイス]ヴィオラ - 01 花園ボイス.mp3
    [ダミーボイス]ザック - 02 雷鳴ボイス EX.mp3
    [星空ボイス]アルテミス - 星空ボイス.mp3
    (2001-01) [サンプルボイス Vol.3]ノヴァ - 01 光明ボイス Vol.3.mp3
    """
    if not file_name.endswith(".mp3"):
        return None

    date_prefix = ""
    original_file_name = file_name
    
    # 先頭に "(YYYY-MM) " がある場合は日付を保存して取り除く
    if file_name.startswith("(") and len(file_name) >= 9 and file_name[8] == ")":
        # フォーマット確認 (YYYY-MM)
        date_part = file_name[1:8]
        if date_part[:4].isdigit() and date_part[4] == "-" and date_part[5:7].isdigit():
            date_prefix = file_name[:9]  # "(YYYY-MM)" を保存
            # 閉じ括弧の次がスペースなら除去
            if len(file_name) > 9 and file_name[9] == " ":
                file_name = file_name[10:]
            else:
                # スペースが無い場合は括弧までを除去
                file_name = file_name[9:]

    base_name = file_name[:-4]
    parts = base_name.split(" - ")
    if len(parts) != 2:
        return None

    prefix_part, track_part = parts
    if not prefix_part.startswith("[") or "]" not in prefix_part:
        return None

    album_name = prefix_part[1 : prefix_part.index("]")]
    artist_name = prefix_part[prefix_part.index("]") + 1 :]

    # 日付プレフィックスがある場合はアルバム名に追加
    if date_prefix:
        album_name = f"{date_prefix} [{album_name}]"

    track_number, track_name = extract_track_info(track_part, artist_name, album_name)

    return ID3Tags(
        track_name=track_name,
        artist_name=artist_name,
        album_name=album_name,
        track_number=track_number,
    )


def write_id3_tags(file_path: Path, tags: ID3Tags, dry_run: bool = False) -> None:
    """ID3タグを書き込む"""
    if dry_run:
        print(f"Dry-run: Would process {file_path} with tags {tags}")
        return

    try:
        audio = EasyID3(file_path)
    except ID3NoHeaderError:
        audio = EasyID3()
        audio.save(file_path)

    audio["title"] = tags["track_name"]
    audio["artist"] = tags["artist_name"]
    audio["album"] = tags["album_name"]
    audio["tracknumber"] = str(tags["track_number"])
    audio.save(file_path)
    print(f"Processed: {file_path}")


def process_files(path: Path, recursive: bool = False) -> list[tuple[Path, ID3Tags]]:
    """ディレクトリ内のファイルにID3タグを書き込む"""
    processed_files = []
    for file_path in get_mp3_files(path, recursive):
        normalized_file_name = normalize_file_name(file_path.name)
        tags = parse_file_name(normalized_file_name)

        if not tags:
            handle_file_parsing_failure(file_path, normalized_file_name)
            continue

        processed_files.append((file_path, tags))
    return processed_files


def setup_preview_gui(root, processed_files, execute_writes, dry_run):
    """ID3タグのプレビューをGUIで表示するためのセットアップ"""
    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    # ウィンドウサイズを設定
    setup_window_size(root)

    tree = ttk.Treeview(
        frame,
        columns=("File Path", "Title", "Artist", "Album", "Track Number"),
        show="headings",
    )
    tree.heading(
        "File Path",
        text="ファイルパス",
        command=lambda: sortby(tree, "File Path", False),
    )
    tree.heading("Title", text="タイトル", command=lambda: sortby(tree, "Title", False))
    tree.heading("Artist", text="アーティスト", command=lambda: sortby(tree, "Artist", False))
    tree.heading("Album", text="アルバム", command=lambda: sortby(tree, "Album", False))
    tree.heading(
        "Track Number",
        text="トラック番号",
        command=lambda: sortby(tree, "Track Number", False),
    )

    for file_path, tags in processed_files:
        tree.insert(
            "",
            tk.END,
            values=(
                file_path,
                tags["track_name"],
                tags["artist_name"],
                tags["album_name"],
                tags["track_number"],
            ),
        )

    tree.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    # スクロールバーを追加
    setup_scrollbar_for_tree(frame, tree)

    # リサイズ設定
    setup_frame_resize_behavior(frame, tree)

    # 列の幅を調整
    tree.column("File Path", width=300)
    tree.column("Title", width=150)
    tree.column("Artist", width=100)
    tree.column("Album", width=100)
    tree.column("Track Number", width=10)

    # 進捗バーを追加
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(processed_files))
    progress_bar.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W + tk.E)

    current_file_var = tk.StringVar()
    current_file_label = ttk.Label(root, textvariable=current_file_var)
    current_file_label.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W + tk.E)

    # ボタンを追加
    execute_button_text = (
        "タグ書き込みを実行 (dry-run のため実際には書き込まれません)"
        if dry_run
        else "タグ書き込みを実行"
    )
    buttons_config = [
        (execute_button_text, execute_writes),
        ("キャンセル", root.destroy),
    ]
    setup_button_frame(root, 3, buttons_config)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=0)
    root.rowconfigure(2, weight=0)
    root.rowconfigure(3, weight=0)

    # Treeviewのセルを部分的にコピー可能にする
    tree.bind("<Double-1>", on_double_click)

    return progress_var, current_file_var


def preview_id3_tags(processed_files: List[Tuple[Path, ID3Tags]], dry_run: bool) -> None:
    """ID3タグのプレビューを表示する（GUI/CUI自動切り替え）"""
    if GUI_AVAILABLE:
        # GUIモード
        preview_id3_tags_gui(processed_files, dry_run)
    else:
        # CUIモード
        print("GUIが利用できないため、CUIモードで実行します。")
        preview_id3_tags_cui(processed_files, dry_run)


def preview_id3_tags_gui(processed_files: List[Tuple[Path, ID3Tags]], dry_run: bool) -> None:
    """ID3タグのプレビューをGUIで表示する"""

    # 対象ファイルが0件の場合はCUIと同じメッセージを表示して正常終了
    if not processed_files:
        print_no_files_message()
        return

    def execute_writes():
        for i, (file_path, tags) in enumerate(processed_files, 1):
            write_id3_tags(file_path, tags, dry_run)
            progress_var.set(i)
            current_file_var.set(f"Processing: {file_path}")
            root.update_idletasks()
        messagebox.showinfo("完了", "ID3タグの書き込みが完了しました。")
        root.destroy()

    # GUIのセットアップ
    root = tk.Tk()
    root.title("ID3タグプレビュー")

    progress_var, current_file_var = setup_preview_gui(
        root, processed_files, execute_writes, dry_run
    )

    root.mainloop()


def display_cui_table(processed_files: List[Tuple[Path, ID3Tags]], dry_run: bool) -> None:
    """CUIでID3タグ情報を表形式で表示する"""
    if not processed_files:
        print_no_files_message()
        return

    display_cui_table_header("ID3タグ書き込みプレビュー", 120)

    # ヘッダー表示
    print(
        f"{'ファイルパス':<40} {'タイトル':<30} {'アーティスト':<20} {'アルバム':<20} {'トラック':<8}"
    )
    print("-" * 120)

    # ファイル一覧表示
    for file_path, tags in processed_files:
        file_name = file_path.name
        title = truncate_text(tags["track_name"], 30)
        artist = truncate_text(tags["artist_name"], 20)
        album = truncate_text(tags["album_name"], 20)
        track = str(tags["track_number"])
        print(f"{file_name:<40} {title:<30} {artist:<20} {album:<20} {track:<8}")

    display_cui_table_footer(len(processed_files), 120)


def write_tags_cui(processed_files: List[Tuple[Path, ID3Tags]], dry_run: bool) -> None:
    """CUIモードでID3タグを書き込む"""
    for file_path, tags in processed_files:
        write_id3_tags(file_path, tags, dry_run)

    display_completion_message(dry_run, "ID3タグの書き込み")


def preview_id3_tags_cui(processed_files: List[Tuple[Path, ID3Tags]], dry_run: bool) -> None:
    """CUIモードでID3タグのプレビューを表示する"""
    display_cui_table(processed_files, dry_run)

    if not processed_files:
        return

    # ユーザーに確認
    if get_user_confirmation("ID3タグの書き込み", dry_run):
        write_tags_cui(processed_files, dry_run)
    else:
        print("キャンセルしました。")


def main():
    parser = setup_common_argument_parser("MP3ファイルにID3タグを付けるスクリプト")

    args = parser.parse_args()
    directory: str = args.directory
    recursive: bool = args.recursive
    dry_run: bool = args.dry_run
    force_cui: bool = args.cui

    files_to_process = process_files(Path(directory), recursive)

    if force_cui:
        # CUIモードを強制
        preview_id3_tags_cui(files_to_process, dry_run)
    else:
        # 通常の自動判定
        preview_id3_tags(files_to_process, dry_run)


if __name__ == "__main__":
    main()
