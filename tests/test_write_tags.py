"""
Tests for write_tags module
"""

import pytest
from nijisanji_voice.commands.write_tags import parse_file_name, extract_track_info
from nijisanji_voice.lib.mp3_utils import ID3Tags


class TestExtractTrackInfo:
    """Tests for extract_track_info function."""

    def test_extract_track_info_with_number(self):
        """Test extracting track info with track number."""
        track_number, track_name = extract_track_info("01 サンプルボイス", "アルテミス", "ABCボイス")
        
        assert track_number == 1
        assert track_name == "[ABCボイス] アルテミス"

    def test_extract_track_info_ex_track(self):
        """Test extracting track info for EX track."""
        track_number, track_name = extract_track_info("02 サンプルボイス EX", "アルテミス", "ABCボイス")
        
        assert track_number == 2
        assert track_name == "サンプルボイス [ABCボイス] アルテミス EX"

    def test_extract_track_info_ex_another_track(self):
        """Test extracting track info for EX(Another) track."""
        track_number, track_name = extract_track_info("03 月光ボイス EX(Another)", "セレスティア", "月光ボイス")

        assert track_number == 3
        assert track_name == "月光ボイス [月光ボイス] セレスティア EX(Another)"

    def test_extract_track_info_without_number(self):
        """Test extracting track info without track number."""
        track_number, track_name = extract_track_info("サンプルボイス", "アルテミス", "ABCボイス")
        
        assert track_number == 1
        assert track_name == "[ABCボイス] アルテミス"

    def test_extract_track_info_invalid_number(self):
        """Test extracting track info with invalid track number."""
        track_number, track_name = extract_track_info("XX サンプルボイス", "アルテミス", "ABCボイス")
        
        assert track_number == 1
        assert track_name == "[ABCボイス] アルテミス"

    def test_extract_track_info_single_word(self):
        """Test extracting track info with single word."""
        track_number, track_name = extract_track_info("ボイス", "アルテミス", "ABCボイス")
        
        assert track_number == 1
        assert track_name == "[ABCボイス] アルテミス"


class TestParseFileNameWriteTags:
    """Tests for parse_file_name function in write_tags context."""

    def test_parse_standard_filename(self):
        """Test parsing standard filename format."""
        filename = "[ABCボイス]アルテミス - 01 サンプルボイス.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "[ABCボイス] アルテミス"
        assert result["artist_name"] == "アルテミス"
        assert result["album_name"] == "ABCボイス"
        assert result["track_number"] == 1

    def test_parse_ex_filename(self):
        """Test parsing EX filename format."""
        filename = "[ABCボイス]アルテミス - 02 サンプルボイス EX.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "サンプルボイス [ABCボイス] アルテミス EX"
        assert result["artist_name"] == "アルテミス"
        assert result["album_name"] == "ABCボイス"
        assert result["track_number"] == 2

    def test_parse_ex_another_filename(self):
        """Test parsing EX(Another) filename format."""
        filename = "[月光ボイス]セレスティア - 03 月光ボイス EX(Another).mp3"
        result = parse_file_name(filename)

        assert result is not None
        assert result["track_name"] == "月光ボイス [月光ボイス] セレスティア EX(Another)"
        assert result["artist_name"] == "セレスティア"
        assert result["album_name"] == "月光ボイス"
        assert result["track_number"] == 3

    def test_parse_filename_without_track_number(self):
        """Test parsing filename without track number."""
        filename = "[星空ボイス]アルテミス - 星空ボイス.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "[星空ボイス] アルテミス"
        assert result["artist_name"] == "アルテミス"
        assert result["album_name"] == "星空ボイス"
        assert result["track_number"] == 1

    def test_parse_filename_with_date_prefix(self):
        """Test parsing filename with date prefix."""
        filename = "(2001-01) [サンプルボイス Vol.3]ノヴァ - 01 光明ボイス Vol.3.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "[(2001-01) [サンプルボイス Vol.3]] ノヴァ"
        assert result["artist_name"] == "ノヴァ"
        assert result["album_name"] == "(2001-01) [サンプルボイス Vol.3]"
        assert result["track_number"] == 1

    def test_parse_filename_complex_names(self):
        """Test parsing filename with complex Japanese names."""
        filename = "[XXXボイス]ヴィオラ - 01 花園ボイス.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "[XXXボイス] ヴィオラ"
        assert result["artist_name"] == "ヴィオラ"
        assert result["album_name"] == "XXXボイス"
        assert result["track_number"] == 1

    def test_parse_filename_with_ex_complex(self):
        """Test parsing complex EX filename."""
        filename = "[ダミーボイス]ザック - 02 雷鳴ボイス EX.mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "雷鳴ボイス [ダミーボイス] ザック EX"
        assert result["artist_name"] == "ザック"
        assert result["album_name"] == "ダミーボイス"
        assert result["track_number"] == 2

    def test_parse_invalid_filename(self):
        """Test parsing invalid filename formats."""
        invalid_filenames = [
            "invalid_format.mp3",
            "missing_brackets.mp3", 
            "[album]artist track.mp3",  # Missing dash
            "not_an_mp3_file.txt",
        ]
        
        for filename in invalid_filenames:
            result = parse_file_name(filename)
            assert result is None, f"Should return None for invalid filename: {filename}"
        
        # Test case with empty artist name (should be valid but produce empty artist)
        result = parse_file_name("[album] - track.mp3")
        assert result is not None
        assert result["artist_name"] == ""
        assert result["album_name"] == "album"

    def test_parse_filename_edge_cases(self):
        """Test parsing filename edge cases."""
        # Empty track name part
        filename = "[テストボイス]テスト - 01 .mp3"
        result = parse_file_name(filename)
        
        assert result is not None
        assert result["track_name"] == "[テストボイス] テスト"
        assert result["track_number"] == 1


class TestID3TagsCompat:
    """Tests for ID3Tags compatibility with new track_name format."""

    def test_id3_tags_structure_new_format(self):
        """Test ID3Tags structure with new track_name format."""
        tags: ID3Tags = {
            "track_name": "[Test Album] Test Artist",
            "artist_name": "Test Artist",
            "album_name": "Test Album",
            "track_number": 1,
        }

        assert tags["track_name"] == "[Test Album] Test Artist"
        assert tags["artist_name"] == "Test Artist"
        assert tags["album_name"] == "Test Album"
        assert tags["track_number"] == 1

        # Verify all required keys are present
        required_keys = {"track_name", "artist_name", "album_name", "track_number"}
        assert set(tags.keys()) == required_keys

    def test_id3_tags_structure_ex_format(self):
        """Test ID3Tags structure with EX track format."""
        tags: ID3Tags = {
            "track_name": "サンプルボイス [Test Album] Test Artist EX",
            "artist_name": "Test Artist", 
            "album_name": "Test Album",
            "track_number": 2,
        }

        assert tags["track_name"] == "サンプルボイス [Test Album] Test Artist EX"
        assert tags["artist_name"] == "Test Artist"
        assert tags["album_name"] == "Test Album"
        assert tags["track_number"] == 2

    def test_parse_filename_with_date_prefix_advanced(self):
        """Test parsing filename with date prefix - advanced cases."""
        # 複数の日付形式のテスト
        test_cases = [
            {
                "filename": "(2025-09) [ABCボイス]アルテミス - 01 サンプルボイス.mp3",
                "expected_album": "(2025-09) [ABCボイス]",
                "expected_track": "[(2025-09) [ABCボイス]] アルテミス",
                "expected_artist": "アルテミス"
            },
            {
                "filename": "(2024-12) [テストボイス]テストアーティスト - 02 テストボイス EX.mp3", 
                "expected_album": "(2024-12) [テストボイス]",
                "expected_track": "テストボイス [(2024-12) [テストボイス]] テストアーティスト EX",
                "expected_artist": "テストアーティスト"
            }
        ]
        
        for case in test_cases:
            result = parse_file_name(case["filename"])
            assert result is not None, f"Failed to parse: {case['filename']}"
            assert result["album_name"] == case["expected_album"]
            assert result["track_name"] == case["expected_track"] 
            assert result["artist_name"] == case["expected_artist"]

    def test_parse_filename_without_date_prefix_advanced(self):
        """Test parsing filename without date prefix - ensure no date formatting."""
        test_cases = [
            {
                "filename": "[ABCボイス]アルテミス - 01 サンプルボイス.mp3",
                "expected_album": "ABCボイス",
                "expected_track": "[ABCボイス] アルテミス",
                "expected_artist": "アルテミス"
            },
            {
                "filename": "[テストボイス]テストアーティスト - 02 テストボイス EX.mp3",
                "expected_album": "テストボイス", 
                "expected_track": "テストボイス [テストボイス] テストアーティスト EX",
                "expected_artist": "テストアーティスト"
            }
        ]
        
        for case in test_cases:
            result = parse_file_name(case["filename"])
            assert result is not None, f"Failed to parse: {case['filename']}"
            assert result["album_name"] == case["expected_album"]
            assert result["track_name"] == case["expected_track"]
            assert result["artist_name"] == case["expected_artist"]