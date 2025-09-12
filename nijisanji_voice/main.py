"""
nijisanji-voice メインエントリーポイント

コマンドライン引数に応じて rename または write-tags コマンドを実行します。
"""

import argparse
import sys


from .commands import rename, write_tags


def create_main_parser():
    """メインのコマンドライン引数パーサーを作成"""
    parser = argparse.ArgumentParser(
        prog="nijisanji-voice",
        description="にじさんじボイスファイル管理ツール",
        epilog="例: nijisanji-voice rename --directory /path/to/files --dry-run",
    )

    subparsers = parser.add_subparsers(dest="command", help="利用可能なコマンド", metavar="COMMAND")

    # rename コマンド
    rename_parser = subparsers.add_parser("rename", help="MP3ファイルをリネーム")
    rename_parser.add_argument(
        "--directory", type=str, required=True, help="処理するディレクトリのパス"
    )
    rename_parser.add_argument(
        "--recursive", action="store_true", help="サブディレクトリを再帰的に処理"
    )
    rename_parser.add_argument(
        "--dry-run", action="store_true", help="実際には変更せずにシミュレーション実行"
    )
    rename_parser.add_argument("--cui", action="store_true", help="GUIではなくCUIモードで実行")

    # write-tags コマンド
    write_tags_parser = subparsers.add_parser("write-tags", help="MP3ファイルにID3タグを書き込み")
    write_tags_parser.add_argument(
        "--directory", type=str, required=True, help="処理するディレクトリのパス"
    )
    write_tags_parser.add_argument(
        "--recursive", action="store_true", help="サブディレクトリを再帰的に処理"
    )
    write_tags_parser.add_argument(
        "--dry-run", action="store_true", help="実際には変更せずにシミュレーション実行"
    )
    write_tags_parser.add_argument("--cui", action="store_true", help="GUIではなくCUIモードで実行")

    return parser


def main():
    """メインエントリーポイント関数"""
    parser = create_main_parser()
    args = parser.parse_args()

    # コマンドが指定されていない場合はヘルプを表示
    if not args.command:
        parser.print_help()
        sys.exit(1)

    # 各コマンドに応じて処理を実行
    try:
        if args.command == "rename":
            # sys.argv を書き換えて rename コマンドに渡す
            # rename コマンドは従来の引数解析を期待している
            sys.argv = ["nijisanji-voice-rename"] + ["--directory", args.directory]
            if args.recursive:
                sys.argv.append("--recursive")
            if args.dry_run:
                sys.argv.append("--dry-run")
            if args.cui:
                sys.argv.append("--cui")

            rename.main()

        elif args.command == "write-tags":
            # sys.argv を書き換えて write_tags コマンドに渡す
            sys.argv = ["nijisanji-voice-write-tags"] + ["--directory", args.directory]
            if args.recursive:
                sys.argv.append("--recursive")
            if args.dry_run:
                sys.argv.append("--dry-run")
            if args.cui:
                sys.argv.append("--cui")

            write_tags.main()

    except KeyboardInterrupt:
        print("\n操作がキャンセルされました。")
        sys.exit(130)
    except Exception as e:
        print(f"エラーが発生しました: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
