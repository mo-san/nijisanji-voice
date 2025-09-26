"""
MP3ファイル処理用の共通ユーティリティ関数とデータ型定義

このモジュールは rename_mp3.py と write_mp3_tags.py で共通して使用される
ファイル名処理や型定義を提供します。
"""

import re
import unicodedata
from pathlib import Path
from typing import TypedDict


class ID3Tags(TypedDict):
    """ID3タグの情報を保持する型定義"""

    track_name: str
    artist_name: str
    album_name: str
    track_number: int


def normalize_file_name(file_name: str) -> str:
    """NFKC正規化を行う"""
    return unicodedata.normalize("NFKC", file_name)


def is_already_properly_formatted(file_name: str) -> bool:
    """ファイル名がすでに適切にフォーマットされているかチェックする"""
    # 目標フォーマット: [アルバム名]アーティスト名 - 01/02 トラック名[ EX].mp3
    # 通常版: [アルバム名]アーティスト名 - 01 トラック名.mp3
    # EX版: [アルバム名]アーティスト名 - 02 トラック名 EX.mp3
    # 先頭に任意で (YYYY-MM) の日付が付与される場合も許容する。

    # 正規表現パターン（先頭に日付が任意で付与される場合を許容）
    pattern = r"^(?:\(\d{4}-\d{2}\)\s)?\[([^\]]+)\](.+?) - (01|02) \1( EX)?\.mp3$"

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


def get_mp3_files(directory_path: Path, recursive: bool = False):
    """指定されたディレクトリからMP3ファイルを取得する"""
    if recursive:
        return directory_path.rglob("*.mp3")
    else:
        return directory_path.glob("*.mp3")


def handle_file_parsing_failure(file_path: Path, normalized_file_name: str) -> None:
    """
    ファイル名の解析に失敗した場合の共通処理
    適切にフォーマット済みかどうかをチェックして適切なメッセージを表示
    """
    if is_already_properly_formatted(normalized_file_name):
        print(f"Already formatted: '{file_path}' - すでにリネーム済み")
    else:
        print(f"Skipped: {file_path} - does not match expected pattern")
