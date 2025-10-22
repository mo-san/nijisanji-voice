# トラック名重複削除修正

## 問題の詳細

`extract_track_info`関数で、トラック名とアルバム名が同じ場合に重複表示される問題。

### 実際の不具合例

```
ファイル名: [Trick or Treat Voice]倉持めると - 02 Trick or Treat Voice EX.mp3

修正前: "Trick or Treat Voice [Trick or Treat Voice] 倉持めると EX"
修正後: "[Trick or Treat Voice] 倉持めると EX"
```

## 修正内容

`nijisanji_voice/commands/write_tags.py`の`extract_track_info`関数に、トラック名とアルバム名が同じ場合の重複削除ロジックを追加。

### 実装ロジック

```python
# EXの処理
elif base_track_name.endswith(" EX"):
    prefix = base_track_name[:-3]  # " EX" を除去
    # トラック名とアルバム名が同じ場合は、トラック名を省略
    if prefix == album_name:
        track_name = f"[{album_name}] {artist_name} EX"
    else:
        track_name = f"{prefix} [{album_name}] {artist_name} EX"
```

### 動作パターン

| トラック名 | アルバム名 | 出力 |
|-----------|-----------|------|
| `Trick or Treat Voice` | `Trick or Treat Voice` | `[Trick or Treat Voice] 倉持めると EX` |
| `サンプルボイス` | `ABCボイス` | `サンプルボイス [ABCボイス] アルテミス EX` |

**ルール**: 
- トラック名 == アルバム名 → `[アルバム] アーティスト EX`
- トラック名 != アルバム名 → `トラック名 [アルバム] アーティスト EX`

## テスト修正

以下のテストケースを修正：
- `test_extract_track_info_ex_another_track`: 重複削除の期待値に修正
- `test_parse_ex_another_filename`: 重複削除の期待値に修正
- `test_parse_filename_without_date_prefix_advanced`: 重複削除の期待値に修正

全88テスト通過。
