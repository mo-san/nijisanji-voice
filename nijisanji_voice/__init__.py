"""
nijisanji-voice: MP3ファイルのリネームとID3タグ書き込みツール

このパッケージはにじさんじボイスファイルの管理を効率化するためのツールです。
"""

__version__ = "0.1.0"
__author__ = "nijisanji-voice project"

# 主要な機能をエクスポート
from .commands import rename, write_tags

__all__ = ["rename", "write_tags"]
