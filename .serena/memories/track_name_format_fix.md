# トラック名フォーマット修正

## 問題

`extract_track_info`関数で、EXでない通常のトラックのタイトルが正しく生成されていなかった。

### 修正前の動作

```python
# ファイル名: [Trick or Treat Voice]倉持めると - 01 Trick or Treat Voice.mp3
# 期待: タイトル = "[Trick or Treat Voice] 倉持めると"
# 実際: タイトル = "[Trick or Treat Voice] 倉持めると"  # 正しい

# ファイル名: [Trick or Treat Voice]倉持めると - 02 Trick or Treat Voice EX.mp3
# 期待: タイトル = "[Trick or Treat Voice] 倉持めると EX"
# 実際: タイトル = "Trick or Treat Voice [Trick or Treat Voice] 倉持めると EX"  # 間違い
```

元のトラック名部分（`Trick or Treat Voice`）が重複して表示されていた。

## 修正内容

`nijisanji_voice/commands/write_tags.py:43-86`の`extract_track_info`関数を修正：

### 変更点

1. **EX(Another)の処理**: 元のトラック名のプレフィックスを削除
   - 変更前: `track_name = track_name[:-12] + f" [{album_name}] {artist_name} EX(Another)"`
   - 変更後: `prefix = base_track_name[:-12]; track_name = f"{prefix} [{album_name}] {artist_name} EX(Another)"`

2. **EXの処理**: 元のトラック名のプレフィックスを削除  
   - 変更前: `track_name = track_name[:-3] + f" [{album_name}] {artist_name} EX"`
   - 変更後: `prefix = base_track_name[:-3]; track_name = f"{prefix} [{album_name}] {artist_name} EX"`

3. **通常トラックの処理**: アルバム名とアーティスト名のみ
   - 変更前・後: `track_name = f"[{album_name}] {artist_name}"`（変更なし）

### 正しい動作

| ファイル名 | タイトル |
|-----------|---------|
| `[XXXボイス]アーティスト - 01 XXXボイス.mp3` | `[XXXボイス] アーティスト` |
| `[XXXボイス]アーティスト - 02 XXXボイス EX.mp3` | `[XXXボイス] アーティスト EX` |
| `[XXXボイス]アーティスト - 03 XXXボイス EX(Another).mp3` | `[XXXボイス] アーティスト EX(Another)` |

※ 通常トラックでは元のトラック名（`XXXボイス`）は含まれない
※ EX/EX(Another)では元のトラック名が削除され、`[アルバム] アーティスト EX`の形式になる

## テスト

全88テストが通過し、既存機能に影響なし。
