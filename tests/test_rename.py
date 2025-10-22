"""
Tests for nijisanji_voice.commands.rename module.
"""

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from nijisanji_voice.commands.rename import (
    parse_file_name,
    generate_new_file_name,
    get_renamed_files,
)


class TestParseFileName:
    """Tests for parse_file_name function."""

    def test_parse_ex_pattern(self):
        """Test parsing EX_ pattern files."""
        result = parse_file_name("EX_エリアナ_ファンタジーボイス.mp3")
        assert result is not None
        assert result["Character"] == "エリアナ"
        assert result["Suffix"] == "ファンタジーボイス"
        assert result["IsEX"] is True
        assert "TrackNumber" not in result

    def test_parse_ex_another_pattern(self):
        """Test parsing EX Another_ pattern files."""
        result = parse_file_name("EX Another_セレスティア_月光ボイス.mp3")
        assert result is not None
        assert result["Character"] == "セレスティア"
        assert result["Suffix"] == "月光ボイス"
        assert result["IsEX"] is True
        assert result["IsAnother"] is True
        assert "TrackNumber" not in result

    def test_parse_standard_pattern(self):
        """Test parsing standard artist_album pattern."""
        result = parse_file_name("リュウ_ドラゴンボイス.mp3")
        assert result is not None
        assert result["Character"] == "リュウ"
        assert result["Suffix"] == "ドラゴンボイス"
        assert result["IsEX"] is False
        assert "TrackNumber" not in result

    def test_parse_numbered_pattern_01(self):
        """Test parsing 01_artist_album pattern."""
        result = parse_file_name("01_アルテミス_星空ボイス.mp3")
        assert result is not None
        assert result["Character"] == "アルテミス"
        assert result["Suffix"] == "星空ボイス"
        assert result["IsEX"] is False
        assert result["TrackNumber"] == "01"

    def test_parse_numbered_patterns_02_to_06(self):
        """Test parsing track numbers 02-06."""
        test_cases = [
            ("02_ザック_雷鳴ボイス.mp3", "02"),
            ("03_ノヴァ_光明ボイス.mp3", "03"),
            ("04_カイ_海風ボイス.mp3", "04"),
            ("05_ヴィオラ_花園ボイス.mp3", "05"),
            ("06_フェニックス_炎舞ボイス.mp3", "06"),
        ]

        for filename, expected_track in test_cases:
            result = parse_file_name(filename)
            assert result is not None, f"Failed to parse {filename}"
            assert result["IsEX"] is False
            assert result["TrackNumber"] == expected_track

    def test_parse_welcome_voice_pattern(self):
        """Test parsing Welcome Voice patterns."""
        result = parse_file_name("02_サンプル音声_オリオン.mp3")
        assert result is not None
        assert (
            result["Character"] == "サンプル音声"
        )  # Note: current logic treats middle part as Character
        assert result["Suffix"] == "オリオン"
        assert result["IsEX"] is False
        assert result["TrackNumber"] == "02"

    def test_parse_invalid_patterns(self):
        """Test parsing invalid file patterns."""
        invalid_cases = [
            "07_artist_album.mp3",  # Invalid track number (not 01-06)
            "",  # Empty string
            "abc",  # Too short (less than 4 chars)
            "EX_only_one_part.mp3",  # EX with wrong number of parts
            "too_many_parts_here_test.mp3",  # Too many parts (4 parts)
            "nounderscores.mp3",  # No underscores at all
        ]

        for filename in invalid_cases:
            result = parse_file_name(filename)
            assert result is None, f"Should not parse invalid filename: {filename}"

    def test_parse_valid_two_part_patterns(self):
        """Test that valid two-part patterns are correctly parsed."""
        # These should be valid according to the current logic
        valid_cases = [
            "single_part.mp3",  # 2 parts - valid standard pattern
            "artist_album.mp3",  # 2 parts - valid standard pattern
        ]

        for filename in valid_cases:
            result = parse_file_name(filename)
            assert result is not None, f"Should parse valid filename: {filename}"
            assert result["IsEX"] is False
            assert "TrackNumber" not in result

    def test_parse_non_mp3_extension(self):
        """Test parsing files with non-MP3 extensions (function strips last 4 chars regardless)."""
        # The function doesn't validate .mp3 extension, it just strips last 4 characters
        # This is by design since it's only called on MP3 files
        # "artist_album.txt" becomes "artist_album" after stripping ".txt"
        result = parse_file_name("artist_album.txt")
        assert result is not None
        assert result["Character"] == "artist"
        assert result["Suffix"] == "album"  # Full "album" after stripping ".txt"

    def test_parse_edge_cases(self):
        """Test parsing edge cases with special characters."""
        # Test with Japanese characters
        result = parse_file_name("01_レイナ_魔法ボイス.mp3")
        assert result is not None
        assert result["Character"] == "レイナ"
        assert result["Suffix"] == "魔法ボイス"
        assert result["TrackNumber"] == "01"

        # Test with mixed characters
        result = parse_file_name("EX_エクシオン_戦闘ボイス.mp3")
        assert result is not None
        assert result["Character"] == "エクシオン"
        assert result["Suffix"] == "戦闘ボイス"
        assert result["IsEX"] is True
    
    def test_parse_irregular_unicode_characters(self):
        """Test parsing files with irregular Unicode character patterns."""
        # 濁点分離パターン
        # は+濁点 -> ば (should be normalized by the system before parsing)
        dakuten_decomposed = "01_た\u3099ち\u3099つ\u3099_は\u3099ひ\u3099ふ\u3099.mp3"
        result = parse_file_name(dakuten_decomposed)
        assert result is not None
        # The normalize_file_name function should handle this, but let's test raw parsing too
        assert result["TrackNumber"] == "01"
        
        # 半濁点分離パターン  
        # は+半濁点 -> ぱ
        handakuten_decomposed = "EX_は\u309Aひ\u309A_ダミーボイス.mp3"
        result = parse_file_name(handakuten_decomposed)
        assert result is not None
        assert result["IsEX"] is True
        
        # カタカナ濁点分離パターン
        katakana_decomposed = "02_ハ\u3099ヒ\u3099_タ\u3099チ\u3099.mp3"
        result = parse_file_name(katakana_decomposed)
        assert result is not None
        assert result["TrackNumber"] == "02"
        
        # 混合パターン（一部分離、一部結合）
        mixed_composition = "03_た\u3099パ_は\u3099ヂ.mp3"
        result = parse_file_name(mixed_composition)
        assert result is not None
        assert result["TrackNumber"] == "03"
        
        # 長音記号の異形パターン
        long_vowel_tilde = "01_モックボイス～_テスト.mp3"  # 波ダッシュ
        result = parse_file_name(long_vowel_tilde)
        assert result is not None
        assert result["TrackNumber"] == "01"
        
        long_vowel_dash = "01_モックボイスー_テスト.mp3"  # 長音記号
        result = parse_file_name(long_vowel_dash)
        assert result is not None
        assert result["TrackNumber"] == "01"


class TestGenerateNewFileName:
    """Tests for generate_new_file_name function."""

    def test_generate_ex_file_name(self):
        """Test generating EX file names."""
        parsed = {"Character": "エリアナ", "Suffix": "ファンタジーボイス", "IsEX": True}
        result = generate_new_file_name(parsed)
        expected = "[ファンタジーボイス]エリアナ - 02 ファンタジーボイス EX.mp3"
        assert result == expected

    def test_generate_ex_another_file_name(self):
        """Test generating EX Another file names."""
        parsed = {"Character": "セレスティア", "Suffix": "月光ボイス", "IsEX": True, "IsAnother": True}
        result = generate_new_file_name(parsed)
        expected = "[月光ボイス]セレスティア - 03 月光ボイス EX(Another).mp3"
        assert result == expected

    def test_generate_standard_file_name(self):
        """Test generating standard file names."""
        parsed = {"Character": "リュウ", "Suffix": "ドラゴンボイス", "IsEX": False}
        result = generate_new_file_name(parsed)
        expected = "[ドラゴンボイス]リュウ - 01 ドラゴンボイス.mp3"
        assert result == expected

    def test_generate_with_track_number(self):
        """Test generating file names with specific track numbers."""
        test_cases = [
            ("01", "[星空ボイス]アルテミス - 01 星空ボイス.mp3"),
            ("02", "[雷鳴ボイス]ザック - 02 雷鳴ボイス.mp3"),
            ("03", "[光明ボイス]ノヴァ - 03 光明ボイス.mp3"),
            ("04", "[海風ボイス]カイ - 04 海風ボイス.mp3"),
            ("05", "[花園ボイス]ヴィオラ - 05 花園ボイス.mp3"),
            ("06", "[炎舞ボイス]フェニックス - 06 炎舞ボイス.mp3"),
        ]

        for track_num, expected in test_cases:
            parsed = {
                "Character": expected.split("]")[1].split(" - ")[
                    0
                ],  # Extract character from expected
                "Suffix": expected.split("[")[1].split("]")[0],  # Extract suffix from expected
                "IsEX": False,
                "TrackNumber": track_num,
            }
            result = generate_new_file_name(parsed)
            assert result == expected

    def test_generate_welcome_voice_pattern(self):
        """Test generating Welcome Voice file names."""
        parsed = {
            "Character": "サンプル音声",
            "Suffix": "オリオン",
            "IsEX": False,
            "TrackNumber": "02",
        }
        result = generate_new_file_name(parsed)
        expected = "[オリオン]サンプル音声 - 02 オリオン.mp3"
        assert result == expected

    def test_track_number_overrides_ex_logic(self):
        """Test that TrackNumber takes precedence over EX logic."""
        # EX file with explicit track number should use that number
        parsed = {
            "Character": "メルクリウス",
            "Suffix": "紫水晶ボイス",
            "IsEX": True,  # This would normally give "02"
            "TrackNumber": "05",  # But this should override
        }
        result = generate_new_file_name(parsed)
        expected = "[紫水晶ボイス]メルクリウス - 05 紫水晶ボイス EX.mp3"
        assert result == expected


class TestGetRenamedFiles:
    """Integration tests for get_renamed_files function."""

    def test_get_renamed_files_basic(self, sample_mp3_files):
        """Test basic file renaming detection."""
        temp_dir = list(sample_mp3_files.values())[0].parent

        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            mock_get_files.return_value = [sample_mp3_files["numbered_02"]]

            renamed_files = get_renamed_files(temp_dir, recursive=False)

            assert len(renamed_files) == 1
            old_path, new_path = renamed_files[0]
            assert old_path.name == "02_ザック_雷鳴ボイス.mp3"
            assert new_path.name == "[雷鳴ボイス]ザック - 02 雷鳴ボイス.mp3"

    def test_get_renamed_files_already_formatted(self, sample_mp3_files, capsys):
        """Test handling of already formatted files."""
        temp_dir = list(sample_mp3_files.values())[0].parent

        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            mock_get_files.return_value = [sample_mp3_files["already_formatted_01"]]

            renamed_files = get_renamed_files(temp_dir, recursive=False)

            # Should return empty list for already formatted files
            assert len(renamed_files) == 0

            # Should print already formatted message
            captured = capsys.readouterr()
            assert "Already formatted" in captured.out
            assert "すでにリネーム済み" in captured.out

    def test_get_renamed_files_invalid_pattern(self, sample_mp3_files, capsys):
        """Test handling of files with invalid patterns."""
        temp_dir = list(sample_mp3_files.values())[0].parent

        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            mock_get_files.return_value = [sample_mp3_files["invalid_wrong_parts"]]

            renamed_files = get_renamed_files(temp_dir, recursive=False)

            # Should return empty list for invalid files
            assert len(renamed_files) == 0

            # Should print skipped message
            captured = capsys.readouterr()
            assert "Skipped" in captured.out
            assert "does not match expected pattern" in captured.out

    def test_get_renamed_files_recursive(self, nested_temp_dir):
        """Test recursive file discovery."""
        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            # Mock finding all files recursively
            expected_files = list(nested_temp_dir.rglob("*.mp3"))
            mock_get_files.return_value = expected_files

            renamed_files = get_renamed_files(nested_temp_dir, recursive=True)

            # Should find files from all directories
            assert len(renamed_files) == 3
            mock_get_files.assert_called_once_with(nested_temp_dir, True)

    def test_get_renamed_files_non_recursive(self, nested_temp_dir):
        """Test non-recursive file discovery."""
        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            # Mock finding only root level files
            root_files = list(nested_temp_dir.glob("*.mp3"))
            mock_get_files.return_value = root_files

            renamed_files = get_renamed_files(nested_temp_dir, recursive=False)

            # Should find only root level files
            assert len(renamed_files) == 1
            mock_get_files.assert_called_once_with(nested_temp_dir, False)
    
    def test_get_renamed_files_with_unicode_normalization(self, unicode_filenames):
        """Test file renaming with Unicode normalization."""
        temp_dir = list(unicode_filenames.values())[0].parent
        
        with patch("nijisanji_voice.commands.rename.get_mp3_files") as mock_get_files:
            # Test with different Unicode normalization files
            test_files = [
                unicode_filenames["dakuten_composed"],     # 正常な濁点文字
                unicode_filenames["dakuten_decomposed"],   # 分離された濁点文字  
                unicode_filenames["handakuten_decomposed"], # 分離された半濁点文字
                unicode_filenames["mixed_normalization"],   # 混合パターン
            ]
            mock_get_files.return_value = test_files
            
            renamed_files = get_renamed_files(temp_dir, recursive=False)
            
            # すべてのファイルが適切にパースされ、リネーム対象として検出されるべき
            assert len(renamed_files) == 4
            
            # 各ファイルが適切にリネームされていることを確認
            for old_path, new_path in renamed_files:
                assert old_path.suffix == ".mp3"
                assert new_path.suffix == ".mp3"
                # 新しいファイル名が適切なフォーマットになっていることを確認
                assert "[" in new_path.name and "]" in new_path.name
                assert " - " in new_path.name
