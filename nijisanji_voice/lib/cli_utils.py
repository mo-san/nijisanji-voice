"""
CLI関連の共通ユーティリティ関数

このモジュールは rename_mp3.py と write_mp3_tags.py で共通して使用される
コマンドライン処理機能を提供します。
"""

import argparse
import unicodedata


def get_user_confirmation(action_text: str, dry_run: bool = False) -> bool:
    """
    ユーザーに確認を求める

    Args:
        action_text: 実行する処理の説明
        dry_run: dry-runモードかどうか

    Returns:
        True: 実行する, False: キャンセル
    """
    if dry_run:
        prompt = f"実行しますか？(dry-run のため実際には書き込まれません) [y/N]: "
    else:
        prompt = f"{action_text}を実行しますか？ [y/N]: "

    while True:
        response = input(f"\n{prompt}").strip().lower()
        if response in ["y", "yes"]:
            return True
        elif response in ["n", "no", ""]:
            return False
        else:
            print("y または n で答えてください。")


def display_completion_message(dry_run: bool, action_name: str) -> None:
    """
    完了メッセージを表示する

    Args:
        dry_run: dry-runモードかどうか
        action_name: 実行した処理の名前
    """
    if not dry_run:
        print(f"{action_name}が完了しました。")


def display_cui_table_header(title: str, width: int = 80) -> None:
    """
    CUIテーブルのヘッダーを表示する

    Args:
        title: テーブルのタイトル
        width: テーブルの幅
    """
    print("\n" + "=" * width)
    print(title)
    print("=" * width)


def display_cui_table_footer(count: int, width: int = 80) -> None:
    """
    CUIテーブルのフッターを表示する

    Args:
        count: 処理対象の件数
        width: テーブルの幅
    """
    print("-" * width)
    print(f"合計: {count}ファイル")
    print("=" * width)


def setup_common_argument_parser(description: str) -> argparse.ArgumentParser:
    """
    共通のコマンドライン引数解析器を設定する

    Args:
        description: スクリプトの説明

    Returns:
        設定されたArgumentParser
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("-d", "--directory", type=str, required=True, help="処理するディレクトリのパス")
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="指定するとサブディレクトリを再帰的に処理する",
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="指定すると実際には書き込まずに処理をシミュレートする",
    )
    parser.add_argument(
        "-c", "--cui", action="store_true", help="指定するとGUIではなくCUIモードで実行する"
    )
    return parser


def print_no_files_message(file_type: str = "処理対象") -> None:
    """ファイルが見つからなかった場合のメッセージを表示"""
    print(f"{file_type}のファイルが見つかりませんでした。")


def truncate_text(text: str, max_length: int) -> str:
    """テキストを指定された長さに切り詰める"""
    if len(text) > max_length:
        return text[: max_length - 3] + "..."
    return text


def natural_sort_key(text: str) -> list:
    """
    自然順ソート用のキーを生成する

    数字を含む文字列を自然な順序でソートするためのキーを返す。
    例: ["file1.mp3", "file2.mp3", "file10.mp3"] → 正しい順序でソート

    Args:
        text: ソート対象の文字列

    Returns:
        ソート用のキーリスト
    """
    import re

    def atoi(s: str):
        return int(s) if s.isdigit() else s

    return [atoi(c) for c in re.split(r"(\d+)", text)]


def get_display_width(text: str) -> int:
    """
    文字列の表示幅を計算する（全角文字を2文字としてカウント）

    Args:
        text: 表示幅を計算する文字列

    Returns:
        表示幅（全角文字は2、半角文字は1としてカウント）
    """
    import unicodedata

    width = 0
    for char in text:
        # East Asian Width プロパティを取得
        # 'F' (Fullwidth) と 'W' (Wide) は全角文字
        if unicodedata.east_asian_width(char) in ("F", "W"):
            width += 2
        else:
            width += 1
    return width


def pad_text(text: str, target_width: int) -> str:
    """
    文字列を指定された表示幅にパディングする（全角文字を考慮）

    Args:
        text: パディングする文字列
        target_width: 目標の表示幅

    Returns:
        パディングされた文字列
    """
    current_width = get_display_width(text)
    padding = target_width - current_width
    if padding > 0:
        return text + " " * padding
    return text


def truncate_text_with_width(text: str, max_width: int) -> str:
    """
    文字列を指定された表示幅に切り詰める（全角文字を考慮）

    Args:
        text: 切り詰める文字列
        max_width: 最大表示幅

    Returns:
        切り詰められた文字列（必要に応じて "..." を追加）
    """
    if get_display_width(text) <= max_width:
        return text

    # "..." の表示幅は3
    ellipsis = "..."
    ellipsis_width = 3
    target_width = max_width - ellipsis_width

    result = ""
    current_width = 0

    for char in text:
        char_width = 2 if unicodedata.east_asian_width(char) in ("F", "W") else 1
        if current_width + char_width > target_width:
            break
        result += char
        current_width += char_width

    return result + ellipsis
