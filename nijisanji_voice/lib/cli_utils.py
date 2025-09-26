"""
CLI関連の共通ユーティリティ関数

このモジュールは rename_mp3.py と write_mp3_tags.py で共通して使用される
コマンドライン処理機能を提供します。
"""

import argparse


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
    parser.add_argument("--directory", type=str, required=True, help="処理するディレクトリのパス")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="指定するとサブディレクトリを再帰的に処理する",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="指定すると実際には書き込まずに処理をシミュレートする",
    )
    parser.add_argument(
        "--cui", action="store_true", help="指定するとGUIではなくCUIモードで実行する"
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
