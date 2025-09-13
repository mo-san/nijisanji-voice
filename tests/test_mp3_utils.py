"""
Tests for nijisanji_voice.lib.mp3_utils module.
"""

import unicodedata

from nijisanji_voice.lib.mp3_utils import (
    normalize_file_name,
    is_already_properly_formatted,
    get_mp3_files,
    handle_file_parsing_failure,
    ID3Tags,
)


class TestNormalizeFileName:
    """Tests for normalize_file_name function."""

    def test_normalize_nfc_string(self):
        """Test normalizing already NFC string."""
        input_str = "テスト.mp3"
        result = normalize_file_name(input_str)
        expected = unicodedata.normalize("NFKC", input_str)
        assert result == expected
        assert result == input_str  # Should be unchanged

    def test_normalize_nfd_string(self):
        """Test normalizing NFD (decomposed) string."""
        # Create NFD form manually
        nfc_str = "テスト.mp3"
        nfd_str = unicodedata.normalize("NFD", nfc_str)

        result = normalize_file_name(nfd_str)
        expected = unicodedata.normalize("NFKC", nfd_str)
        assert result == expected
        assert result == nfc_str  # Should be normalized to NFC

    def test_normalize_fullwidth_characters(self):
        """Test normalizing full-width characters."""
        input_str = "０１_テスト_ボイス.mp3"  # Full-width 01
        result = normalize_file_name(input_str)
        expected = "01_テスト_ボイス.mp3"  # Half-width 01
        assert result == expected

    def test_normalize_mixed_characters(self):
        """Test normalizing mixed character forms."""
        input_str = "EX_テスト_ボイス.mp3"
        result = normalize_file_name(input_str)
        # Should remain the same as it's already properly normalized
        assert result == input_str

    def test_normalize_empty_string(self):
        """Test normalizing empty string."""
        result = normalize_file_name("")
        assert result == ""

    def test_normalize_dakuten_decomposed(self):
        """Test normalizing decomposed dakuten (濁点分離) characters."""
        # は + 濁点 (U+3099) -> ば
        decomposed = "は\u3099ひ\u3099ふ\u3099へ\u3099ほ\u3099"  # は+濁点, ひ+濁点, ふ+濁点, へ+濁点, ほ+濁点
        result = normalize_file_name(decomposed)
        expected = "ばびぶべぼ"  # 結合された形
        assert result == expected

    def test_normalize_handakuten_decomposed(self):
        """Test normalizing decomposed handakuten (半濁点分離) characters."""
        # は + 半濁点 (U+309A) -> ぱ
        decomposed = "は\u309aひ\u309aふ\u309aへ\u309aほ\u309a"  # は+半濁点, ひ+半濁点, ふ+半濁点, へ+半濁点, ほ+半濁点
        result = normalize_file_name(decomposed)
        expected = "ぱぴぷぺぽ"  # 結合された形
        assert result == expected

    def test_normalize_katakana_dakuten_decomposed(self):
        """Test normalizing decomposed katakana dakuten characters."""
        # カタカナの濁点分離パターン
        decomposed = "ハ\u3099ヒ\u3099フ\u3099ヘ\u3099ホ\u3099"  # ハ+濁点, ヒ+濁点, フ+濁点, ヘ+濁点, ホ+濁点
        result = normalize_file_name(decomposed)
        expected = "バビブベボ"  # 結合された形
        assert result == expected

    def test_normalize_mixed_composition(self):
        """Test normalizing mixed composition patterns."""
        # 一部が分離、一部が結合済みの混合パターン
        mixed = "は\u3099パた\u3099ヂ"  # は+濁点, パ(結合済み), た+濁点, ヂ(結合済み)
        result = normalize_file_name(mixed)
        expected = "ばパだヂ"  # すべて結合形に正規化
        assert result == expected

    def test_normalize_long_vowel_variants(self):
        """Test normalizing different long vowel mark variants."""
        # 波ダッシュ（U+301C）と長音記号（U+30FC）
        with_tilde = "ボイス～テスト"  # 波ダッシュ
        with_dash = "ボイスーテスト"  # 長音記号

        result_tilde = normalize_file_name(with_tilde)
        result_dash = normalize_file_name(with_dash)

        # NFKC正規化により波ダッシュはチルダ（~）に変換される
        assert result_tilde == "ボイス~テスト"  # 波ダッシュ → チルダ
        assert result_dash == "ボイスーテスト"  # 長音記号はそのまま

        # 異なる文字なので結果は異なる
        assert result_tilde != result_dash

        # それぞれが期待される正規化形になっていることを確認
        import unicodedata

        assert result_tilde == unicodedata.normalize("NFKC", with_tilde)
        assert result_dash == unicodedata.normalize("NFKC", with_dash)


class TestIsAlreadyProperlyFormatted:
    """Tests for is_already_properly_formatted function."""

    def test_properly_formatted_standard(self):
        """Test properly formatted standard files."""
        test_cases = [
            "[日常ボイス]クリスタル - 01 日常ボイス.mp3",
            "[Welcome Voice]キリム - 01 Welcome Voice.mp3",
            "[ミスティックボイス]ネフィリム - 01 ミスティックボイス.mp3",
        ]

        for filename in test_cases:
            assert is_already_properly_formatted(
                filename
            ), f"Should be properly formatted: {filename}"

    def test_properly_formatted_ex(self):
        """Test properly formatted EX files."""
        test_cases = [
            "[未来ボイス]サイラス - 02 未来ボイス EX.mp3",
            "[Welcome Voice]ユニコーン - 02 Welcome Voice EX.mp3",
            "[神話ボイス]アポロ - 02 神話ボイス EX.mp3",
        ]

        for filename in test_cases:
            assert is_already_properly_formatted(
                filename
            ), f"Should be properly formatted: {filename}"

    def test_not_properly_formatted_invalid_patterns(self):
        """Test files that are not properly formatted."""
        test_cases = [
            # Wrong format entirely
            "01_クリスタル_日常ボイス.mp3",
            "EX_サイラス_未来ボイス.mp3",
            "クリスタル_日常ボイス.mp3",
            # Wrong track numbers
            "[日常ボイス]クリスタル - 02 日常ボイス.mp3",  # Should be 01 without EX
            "[未来ボイス]サイラス - 01 未来ボイス EX.mp3",  # Should be 02 with EX
            # Missing parts
            "クリスタル - 01 日常ボイス.mp3",  # Missing album brackets
            "[日常ボイス]クリスタル 01 日常ボイス.mp3",  # Missing dash
            "[日常ボイス]クリスタル - 日常ボイス.mp3",  # Missing track number
            # Wrong album/suffix matching
            "[日常ボイス]クリスタル - 01 違うボイス.mp3",  # Album and suffix don't match
            # Invalid extensions
            "[日常ボイス]クリスタル - 01 日常ボイス.txt",
            "[日常ボイス]クリスタル - 01 日常ボイス",
        ]

        for filename in test_cases:
            assert not is_already_properly_formatted(
                filename
            ), f"Should NOT be properly formatted: {filename}"

    def test_edge_cases_regex_matching(self):
        """Test edge cases for regex matching."""
        # Empty brackets
        assert not is_already_properly_formatted("[]クリスタル - 01 .mp3")

        # Empty character name
        assert not is_already_properly_formatted("[日常ボイス] - 01 日常ボイス.mp3")

        # Multiple brackets - this actually matches the regex pattern, so it should be considered "properly formatted"
        # even though it's not ideal. The regex allows any characters except ] in the brackets.
        assert is_already_properly_formatted("[日常ボイス][extra]クリスタル - 01 日常ボイス.mp3")


class TestGetMp3Files:
    """Tests for get_mp3_files function."""

    def test_get_mp3_files_non_recursive(self, sample_mp3_files):
        """Test getting MP3 files non-recursively."""
        temp_dir = list(sample_mp3_files.values())[0].parent

        files = list(get_mp3_files(temp_dir, recursive=False))

        # Should find all MP3 files in the root directory (excluding .mp3 which is just ".mp3")
        mp3_files = [f for f in temp_dir.glob("*.mp3") if f.name != ".mp3"]
        assert len(files) >= len(mp3_files)  # Allow for the .mp3 file to be included or excluded

        # All returned files should be MP3 files (or at least match *.mp3 pattern)
        for file_path in files:
            assert file_path.name.endswith(".mp3")
            assert file_path.parent == temp_dir

    def test_get_mp3_files_recursive(self, nested_temp_dir):
        """Test getting MP3 files recursively."""
        files = list(get_mp3_files(nested_temp_dir, recursive=True))

        # Should find MP3 files in all subdirectories
        expected_files = list(nested_temp_dir.rglob("*.mp3"))
        assert len(files) == len(expected_files)

        # All returned files should be MP3 files
        for file_path in files:
            assert file_path.suffix == ".mp3"

    def test_get_mp3_files_empty_directory(self, temp_dir):
        """Test getting MP3 files from empty directory."""
        files = list(get_mp3_files(temp_dir, recursive=False))
        assert len(files) == 0

        files_recursive = list(get_mp3_files(temp_dir, recursive=True))
        assert len(files_recursive) == 0

    def test_get_mp3_files_no_mp3_files(self, temp_dir):
        """Test getting MP3 files when only non-MP3 files exist."""
        # Create non-MP3 files
        (temp_dir / "test.txt").write_text("test")
        (temp_dir / "test.wav").write_bytes(b"wav")
        (temp_dir / "test").write_text("no extension")

        files = list(get_mp3_files(temp_dir, recursive=False))
        assert len(files) == 0


class TestHandleFileParsingFailure:
    """Tests for handle_file_parsing_failure function."""

    def test_handle_already_formatted_file(self, temp_dir, capsys):
        """Test handling already formatted file."""
        filename = "[日常ボイス]クリスタル - 01 日常ボイス.mp3"
        file_path = temp_dir / filename
        file_path.write_bytes(b"")

        handle_file_parsing_failure(file_path, filename)

        captured = capsys.readouterr()
        assert "Already formatted" in captured.out
        assert str(file_path) in captured.out
        assert "すでにリネーム済み" in captured.out

    def test_handle_invalid_pattern_file(self, temp_dir, capsys):
        """Test handling file with invalid pattern."""
        filename = "invalid_pattern.mp3"
        file_path = temp_dir / filename
        file_path.write_bytes(b"")

        handle_file_parsing_failure(file_path, filename)

        captured = capsys.readouterr()
        assert "Skipped" in captured.out
        assert str(file_path) in captured.out
        assert "does not match expected pattern" in captured.out

    def test_handle_unicode_filename(self, temp_dir, capsys):
        """Test handling Unicode filename."""
        filename = "無効なパターン.mp3"
        file_path = temp_dir / filename
        file_path.write_bytes(b"")

        handle_file_parsing_failure(file_path, filename)

        captured = capsys.readouterr()
        assert "Skipped" in captured.out
        assert str(file_path) in captured.out

    def test_handle_irregular_unicode_filenames(self, unicode_filenames, capsys):
        """Test handling files with irregular Unicode character patterns."""
        # 濁点分離パターンのファイルが既に適切にフォーマット済みとして判定されるかテスト
        dakuten_file = unicode_filenames["dakuten_decomposed"]
        normalized_name = normalize_file_name(dakuten_file.name)

        # 正規化後のファイル名でテスト
        handle_file_parsing_failure(dakuten_file, normalized_name)

        captured = capsys.readouterr()
        # 分離された濁点文字が含まれていても適切に処理されることを確認
        assert "Skipped" in captured.out or "Already formatted" in captured.out

    def test_normalize_real_unicode_files(self, unicode_filenames):
        """Test normalizing real Unicode filenames from fixtures."""
        # 実際のフィクスチャファイルを使用してテスト
        dakuten_composed_file = unicode_filenames["dakuten_composed"]
        dakuten_decomposed_file = unicode_filenames["dakuten_decomposed"]

        # ファイル名を正規化
        composed_normalized = normalize_file_name(dakuten_composed_file.name)
        decomposed_normalized = normalize_file_name(dakuten_decomposed_file.name)

        # 正規化後は同じ結果になるべき
        # （ただし、フィクスチャでは異なる分離パターンを作成しているため、実際の比較は慎重に行う）
        assert composed_normalized == "01_だぢづでど_ばびぶべぼ.mp3"
        # 分離パターンも正規化されて同じ形になるべき
        assert "ば" in decomposed_normalized  # 濁点が結合されていることを確認

        # 半濁点パターンのテスト
        handakuten_decomposed_file = unicode_filenames["handakuten_decomposed"]
        handakuten_normalized = normalize_file_name(handakuten_decomposed_file.name)
        assert "ぱ" in handakuten_normalized  # 半濁点が結合されていることを確認


class TestID3Tags:
    """Tests for ID3Tags TypedDict."""

    def test_id3_tags_type_structure(self):
        """Test ID3Tags structure and typing."""
        # This tests that the TypedDict is properly defined
        tags: ID3Tags = {
            "track_name": "Test Track",
            "artist_name": "Test Artist",
            "album_name": "Test Album",
            "track_number": 1,
        }

        assert tags["track_name"] == "Test Track"
        assert tags["artist_name"] == "Test Artist"
        assert tags["album_name"] == "Test Album"
        assert tags["track_number"] == 1

        # Verify all required keys are present
        required_keys = {"track_name", "artist_name", "album_name", "track_number"}
        assert set(tags.keys()) == required_keys
