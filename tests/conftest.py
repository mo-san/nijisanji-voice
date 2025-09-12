"""
pytest configuration and fixtures for nijisanji-voice tests.
"""

import tempfile
from pathlib import Path
from typing import Generator

import pytest


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        yield Path(tmp_dir)


@pytest.fixture
def sample_mp3_files(temp_dir: Path) -> dict[str, Path]:
    """Create sample MP3 files for testing."""
    files = {
        # EX patterns
        "ex_standard": "EX_鏑木ろこ_内緒話ボイス.mp3",
        # Standard patterns
        "standard": "倉持めると_内緒話ボイス.mp3",
        # Numbered patterns (01-06)
        "numbered_01": "01_蝸堂みかる_内緒話ボイス.mp3",
        "numbered_02": "02_鏑木ろこ_内緒話ボイス.mp3",
        "numbered_03": "03_倉持めると_内緒話ボイス.mp3",
        "numbered_04": "04_獅子堂あかり_内緒話ボイス.mp3",
        "numbered_05": "05_蝸堂みかる_内緒話ボイス.mp3",
        "numbered_06": "06_鏑木ろこ_内緒話ボイス.mp3",
        # Welcome Voice patterns
        "welcome_02": "02_Welcome Voice_猫屋敷美紅.mp3",
        "welcome_03": "03_Welcome Voice_十河ののは.mp3",
        # Already formatted files
        "already_formatted_01": "[内緒話ボイス]倉持めると - 01 内緒話ボイス.mp3",
        "already_formatted_02": "[内緒話ボイス]鏑木ろこ - 02 内緒話ボイス EX.mp3",
        # Invalid patterns
        "invalid_no_extension": "invalid_file_name",
        "invalid_wrong_parts": "truly_invalid_pattern.mp3",
        "invalid_empty": ".mp3",
    }

    created_files = {}
    for key, filename in files.items():
        file_path = temp_dir / filename
        # Create empty MP3 files for testing
        file_path.write_bytes(b"")
        created_files[key] = file_path

    return created_files


@pytest.fixture
def nested_temp_dir(temp_dir: Path) -> Path:
    """Create nested directory structure for recursive testing."""
    # Create subdirectories
    sub1 = temp_dir / "sub1"
    sub2 = temp_dir / "sub2"
    sub1.mkdir()
    sub2.mkdir()

    # Create files in different locations
    (temp_dir / "01_root_内緒話ボイス.mp3").write_bytes(b"")
    (sub1 / "02_sub1_内緒話ボイス.mp3").write_bytes(b"")
    (sub2 / "03_sub2_内緒話ボイス.mp3").write_bytes(b"")

    return temp_dir


@pytest.fixture
def unicode_filenames(temp_dir: Path) -> dict[str, Path]:
    """Create files with various Unicode characters for normalization testing."""
    import unicodedata
    
    files = {
        # NFKC normalization test cases
        "nfc": "01_テスト_ボイス.mp3",  # Already NFC
        "nfd": "01_テスト_ボイス.mp3",  # NFD form (decomposed) - will be created below
        "fullwidth": "０１_テスト_ボイス.mp3",  # Full-width characters
        
        # 濁点・半濁点分離パターンのテストケース
        "dakuten_composed": "01_ばびぶべぼ_だぢづでど.mp3",  # 1文字の濁点文字
        "dakuten_decomposed": "01_ばびぶべぼ_だぢづでど.mp3",  # 分離形式 - will be created below
        "handakuten_composed": "01_ぱぴぷぺぽ_ボイス.mp3",  # 1文字の半濁点文字
        "handakuten_decomposed": "01_ぱぴぷぺぽ_ボイス.mp3",  # 分離形式 - will be created below
        
        # カタカナでも同様のパターン
        "katakana_dakuten_composed": "EX_バビブベボ_ダヂヅデド.mp3",
        "katakana_dakuten_decomposed": "EX_バビブベボ_ダヂヅデド.mp3",  # 分離形式 - will be created below
        
        # 混合パターン（一部が分離、一部が結合）
        "mixed_normalization": "02_ばパ_だヂ.mp3",  # 混合形式 - will be created below
        
        # 長音記号の異形
        "long_vowel_tilde": "01_ボイス～_テスト.mp3",  # 波ダッシュ
        "long_vowel_dash": "01_ボイスー_テスト.mp3",  # 長音記号
    }

    created_files = {}
    for key, filename in files.items():
        # 特定のキーに対して、分離形式（NFD）を作成
        if key == "nfd":
            # NFD形式で作成
            nfd_filename = unicodedata.normalize("NFD", filename)
            file_path = temp_dir / nfd_filename
        elif key == "dakuten_decomposed":
            # 濁点を分離した形式で作成: は+濁点, は+濁点, など
            # "ば" -> "は" + "◌゙" (U+3099)
            # 各文字に濁点を追加
            decomposed_with_dakuten = "01_は\u3099ひ\u3099ふ\u3099へ\u3099ほ\u3099_た\u3099ち\u3099つ\u3099て\u3099と\u3099.mp3"
            file_path = temp_dir / decomposed_with_dakuten
        elif key == "handakuten_decomposed":
            # 半濁点を分離した形式で作成: は+半濁点
            # "ぱ" -> "は" + "◌゚" (U+309A)
            decomposed_with_handakuten = "01_は\u309Aひ\u309Aふ\u309Aへ\u309Aほ\u309A_ボイス.mp3"
            file_path = temp_dir / decomposed_with_handakuten
        elif key == "katakana_dakuten_decomposed":
            # カタカナの濁点を分離
            decomposed_katakana = "EX_ハ\u3099ヒ\u3099フ\u3099ヘ\u3099ホ\u3099_タ\u3099チ\u3099ツ\u3099テ\u3099ト\u3099.mp3"
            file_path = temp_dir / decomposed_katakana
        elif key == "mixed_normalization":
            # 一部分離、一部結合の混合パターン
            mixed = "02_は\u3099パ_た\u3099ヂ.mp3"  # は+濁点、パ（結合済み）、た+濁点、ヂ（結合済み）
            file_path = temp_dir / mixed
        else:
            file_path = temp_dir / filename
        
        # Create empty files for testing
        file_path.write_bytes(b"")
        created_files[key] = file_path

    return created_files
