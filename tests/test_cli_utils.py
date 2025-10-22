"""CLI utility functions tests"""

from nijisanji_voice.lib.cli_utils import (
    natural_sort_key,
    get_display_width,
    pad_text,
    truncate_text_with_width,
)


class TestNaturalSortKey:
    """natural_sort_key関数のテスト"""

    def test_natural_sort_numbers(self):
        """数字を含む文字列の自然順ソート"""
        filenames = ["file10.mp3", "file2.mp3", "file1.mp3", "file20.mp3"]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == ["file1.mp3", "file2.mp3", "file10.mp3", "file20.mp3"]

    def test_natural_sort_mixed(self):
        """文字と数字が混在する複雑なケース"""
        filenames = [
            "track10_song.mp3",
            "track2_song.mp3",
            "track1_song.mp3",
            "track100_song.mp3",
        ]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == [
            "track1_song.mp3",
            "track2_song.mp3",
            "track10_song.mp3",
            "track100_song.mp3",
        ]

    def test_natural_sort_prefix(self):
        """接頭辞付きファイル名のソート"""
        filenames = [
            "2025-01_album_10.mp3",
            "2025-01_album_2.mp3",
            "2025-01_album_1.mp3",
        ]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == [
            "2025-01_album_1.mp3",
            "2025-01_album_2.mp3",
            "2025-01_album_10.mp3",
        ]

    def test_natural_sort_text_only(self):
        """数字を含まない文字列のソート"""
        filenames = ["charlie.mp3", "alpha.mp3", "bravo.mp3"]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == ["alpha.mp3", "bravo.mp3", "charlie.mp3"]

    def test_natural_sort_multiple_numbers(self):
        """複数の数字を含むファイル名のソート"""
        filenames = ["v1_track10.mp3", "v1_track2.mp3", "v2_track1.mp3", "v1_track1.mp3"]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == [
            "v1_track1.mp3",
            "v1_track2.mp3",
            "v1_track10.mp3",
            "v2_track1.mp3",
        ]

    def test_natural_sort_real_world_example(self):
        """実際のファイル名に近いケース"""
        filenames = [
            "セレスティア_月光ボイス_10_夜想曲.mp3",
            "セレスティア_月光ボイス_2_星の唄.mp3",
            "セレスティア_月光ボイス_1_銀河の調べ.mp3",
            "セレスティア_月光ボイス_20_流星の詩.mp3",
        ]
        sorted_files = sorted(filenames, key=natural_sort_key)
        assert sorted_files == [
            "セレスティア_月光ボイス_1_銀河の調べ.mp3",
            "セレスティア_月光ボイス_2_星の唄.mp3",
            "セレスティア_月光ボイス_10_夜想曲.mp3",
            "セレスティア_月光ボイス_20_流星の詩.mp3",
        ]



class TestGetDisplayWidth:
    """get_display_width関数のテスト"""

    def test_ascii_only(self):
        """ASCII文字のみの場合"""
        assert get_display_width("hello") == 5
        assert get_display_width("test123") == 7

    def test_fullwidth_only(self):
        """全角文字のみの場合"""
        assert get_display_width("エリアナ") == 8  # 4文字 × 2
        assert get_display_width("ファンタジー") == 12  # 6文字 × 2

    def test_mixed_characters(self):
        """半角と全角が混在する場合"""
        assert get_display_width("エリアナ_test") == 13  # 4×2 + 1 + 4×1
        assert get_display_width("星空ボイス_01") == 13  # 5×2 + 1 + 2×1

    def test_empty_string(self):
        """空文字列の場合"""
        assert get_display_width("") == 0


class TestPadText:
    """pad_text関数のテスト"""

    def test_ascii_padding(self):
        """ASCII文字のパディング"""
        result = pad_text("test", 10)
        assert len(result) == 10
        assert result == "test      "

    def test_fullwidth_padding(self):
        """全角文字のパディング"""
        result = pad_text("エリアナ", 10)
        # 表示幅8 + スペース2 = 10
        assert get_display_width(result) == 10
        assert result == "エリアナ  "

    def test_mixed_padding(self):
        """混在文字のパディング"""
        result = pad_text("エリアナ_01", 20)
        # 表示幅11 + スペース9 = 20
        assert get_display_width(result) == 20

    def test_exact_width(self):
        """目標幅と一致する場合"""
        text = "test"
        result = pad_text(text, 4)
        assert result == text

    def test_over_width(self):
        """目標幅を超える場合（パディングなし）"""
        text = "verylongtext"
        result = pad_text(text, 5)
        assert result == text


class TestTruncateTextWithWidth:
    """truncate_text_with_width関数のテスト"""

    def test_ascii_no_truncate(self):
        """ASCII文字で切り詰め不要"""
        text = "hello"
        result = truncate_text_with_width(text, 10)
        assert result == "hello"

    def test_ascii_truncate(self):
        """ASCII文字の切り詰め"""
        text = "verylongtext"
        result = truncate_text_with_width(text, 10)
        assert result == "verylon..."
        assert get_display_width(result) == 10

    def test_fullwidth_no_truncate(self):
        """全角文字で切り詰め不要"""
        text = "エリアナ"
        result = truncate_text_with_width(text, 10)
        assert result == "エリアナ"

    def test_fullwidth_truncate(self):
        """全角文字の切り詰め"""
        text = "セレスティア星空ボイス"
        result = truncate_text_with_width(text, 10)
        # 表示幅10以下: 全角3文字(6) + "..."(3) = 9
        assert get_display_width(result) <= 10
        assert result.endswith("...")

    def test_mixed_truncate(self):
        """混在文字の切り詰め"""
        text = "エリアナ_very_long_album_name"
        result = truncate_text_with_width(text, 15)
        assert get_display_width(result) <= 15
        assert result.endswith("...")

    def test_exact_width_boundary(self):
        """境界値テスト: ちょうど目標幅"""
        text = "エリアナ"  # 表示幅8
        result = truncate_text_with_width(text, 8)
        assert result == text
        assert get_display_width(result) == 8
