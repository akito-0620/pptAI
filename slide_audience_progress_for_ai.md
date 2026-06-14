# SlideAudienceAddIn 進捗メモ for 生成AI

更新日: 2026-05-04

## 概要

- プロジェクト: `C:\Users\akito\codex_proj\pptAI\SlideAudienceAddIn`
- 種別: PowerPoint VSTO Add-in
- 目的: PowerPoint スライドショー中に現在スライドを解析し、短い観客コメントを生成して透明 WPF overlay に表示する
- 主な表示モード: `Panel` / `Flow` / `Bubble`

## 現在できていること

- PowerPoint VSTO Add-in として動作
- `SlideShowBegin` / `SlideShowNextSlide` / `SlideShowEnd` を検知
- `Slide.Export` による PNG 保存
- `SlideTextExtractor` によるスライド内テキスト抽出
- Gemini API によるコメント生成
- `EnableApi=false` / API失敗 / APIキーなし時の dummy fallback
- コメントは最大10件程度生成
- `AudienceComment` に `Type` / `Persona` / `Text` / `Confidence`
- 10〜15文字程度、最大18文字以内、句点なしに正規化
- WPF 透明 overlay 表示
- overlay は専用 STA Dispatcher thread 上で更新
- クリック透過は `WS_EX_TRANSPARENT`
- `Panel` / `Flow` / `Bubble` の3モード
- 余白優先配置
- JSONL logging
- PowerPoint Ribbon UI から実験条件・表示条件を変更
- UI変更は `_settings.SaveLocal()` で `appsettings.local.json` に保存
- APIキー本文は UI / 設定 / JSONL に出さない

## 主要ファイル

- `SlideAudienceAddIn\ThisAddIn.cs`
- `SlideAudienceAddIn\Ribbon\SlideAudienceRibbon.cs`
- `SlideAudienceAddIn\Models\AppSettings.cs`
- `SlideAudienceAddIn\Models\AudienceComment.cs`
- `SlideAudienceAddIn\Models\CommentGenerationResult.cs`
- `SlideAudienceAddIn\Models\OverlayDisplayMode.cs`
- `SlideAudienceAddIn\Services\CommentGenerationService.cs`
- `SlideAudienceAddIn\Services\ExperimentLogger.cs`
- `SlideAudienceAddIn\Services\SlideExporter.cs`
- `SlideAudienceAddIn\Services\SlideTextExtractor.cs`
- `SlideAudienceAddIn\Services\SlideWhitespaceAnalyzer.cs`
- `SlideAudienceAddIn\Overlay\OverlayController.cs`
- `SlideAudienceAddIn\Overlay\OverlayWindow.xaml`
- `SlideAudienceAddIn\Overlay\OverlayWindow.xaml.cs`
- `SlideAudienceAddIn\Config\appsettings.example.json`
- `SlideAudienceAddIn\Config\appsettings.local.json`
- `SlideAudienceAddIn\SlideAudienceAddIn.csproj`

## Ribbon UI

実装ファイル:

- `SlideAudienceAddIn\Ribbon\SlideAudienceRibbon.cs`
- `SlideAudienceAddIn\ThisAddIn.cs`

Ribbon:

- タブ名: `SlideAudience`
- `ThisAddIn.CreateRibbonExtensibilityObject()` で `new SlideAudienceRibbon(this)` を返す
- `SlideAudienceRibbon.cs` は `Office.IRibbonExtensibility`

現在の UI group:

- `Run`
  - Enable
  - Test Current Slide
  - Clear Overlay
  - Open Log Folder
- `Condition`
  - Preset: `None` / `Panel` / `Flow` / `Bubble` / `Debug`
  - Display Mode: `Panel` / `Flow` / `Bubble`
- `Generation`
  - Gemini API toggle
  - Max Comments: `3` / `5` / `10`
- `Placement`
  - Whitespace toggle
  - Max Visible: `1` / `2` / `3`
- `Timing`
  - Interval Min: `0.5` / `1.0` / `1.5` / `2.0` / `2.5` / `3.0` / `4.0` / `5.0`
  - Interval Max: 同上
- `Appearance`
  - Font Size: `18` / `22` / `26` / `28` / `32` / `36` / `42` / `48`
  - Text Color: White / Black / Yellow / Cyan / Pink / Green
  - Background: Transparent / Light Black / Black / Dark Black / White / Yellow
- `Motion`
  - Whitespace Prob: `0.0` / `0.25` / `0.5` / `0.7` / `0.85` / `1.0`
  - Flow Speed: `40` / `60` / `80` / `100` / `120` / `160` / `200`
  - Bubble In: `0.2` / `0.5` / `0.8` / `1.0` / `1.5` / `2.0`
  - Bubble Out: 同上

UI 操作時の基本処理:

- `_settings` を更新
- 必要に応じて `_overlayController.ApplySettings(_settings.Overlay)` を呼ぶ
- `_settings.SaveLocal()` で `appsettings.local.json` に保存
- `ExperimentLogger.LogSettingsChanged(...)` を記録
- `Debug.WriteLine` に変更内容を出す

## Overlay 設定

`OverlaySettings` に現在ある主なプロパティ:

```csharp
public OverlayDisplayMode DisplayMode { get; set; } = OverlayDisplayMode.Panel;
public double CommentLifetimeSeconds { get; set; } = 7;
public double FlowSpeedPixelsPerSecond { get; set; } = 80;
public double BubbleLifetimeSeconds { get; set; } = 9;
public double BubbleFadeSeconds { get; set; } = 1.0;
public double BubbleFadeInSeconds { get; set; } = 1.0;
public double BubbleFadeOutSeconds { get; set; } = 1.0;
public double CommentDisplayIntervalSeconds { get; set; } = 1.2;
public double CommentDisplayIntervalMinSeconds { get; set; } = 1.0;
public double CommentDisplayIntervalMaxSeconds { get; set; } = 2.5;
public double CommentFontSize { get; set; } = 28;
public string CommentTextColor { get; set; } = "#FFFFFFFF";
public string CommentBackgroundColor { get; set; } = "#99000000";
public bool UseWhitespaceAwarePlacement { get; set; } = true;
public double WhitespacePlacementProbability { get; set; } = 0.7;
public int WhitespaceGridColumns { get; set; } = 12;
public int WhitespaceGridRows { get; set; } = 8;
public int MaxSimultaneousComments { get; set; } = 3;
```

後方互換:

- `CommentDisplayIntervalSeconds` は残している
- `BubbleFadeSeconds` は残している
- 実処理では新しい `CommentDisplayIntervalMinSeconds` / `MaxSeconds`、`BubbleFadeInSeconds` / `BubbleFadeOutSeconds` を優先

## Flow / Bubble キュー制御

実装ファイル:

- `SlideAudienceAddIn\Overlay\OverlayWindow.xaml.cs`

仕様:

- 生成コメント数は最大10件程度
- `MaxSimultaneousComments` は「同時表示の上限」
- 初回に3件同時投入しない
- まず1件だけ即時表示
- 2件目以降は `CommentDisplayIntervalMinSeconds` 〜 `CommentDisplayIntervalMaxSeconds` のランダム待ち時間後に1件ずつ表示
- active count が `MaxSimultaneousComments` 以上なら dequeue せず queue に残して待つ
- 表示中コメントを途中 prune しない
- スライド切替時・終了時・Clear Overlay 時は queue / timer / animation / Canvas children を消す
- 古い timer callback は `_displaySessionId` で無視

Flow:

- `DoubleAnimation` で `Canvas.LeftProperty` を移動
- `startX = windowWidth + margin`
- `endX = -estimatedCommentWidth - margin`
- `durationSeconds = distance / FlowSpeedPixelsPerSecond`
- コメント全体が左外へ抜けてから remove
- 空いている lane を選ぶ
- 直前 lane と違う lane を優先

Bubble:

- `Storyboard` + `DoubleAnimationUsingKeyFrames` で opacity を制御
- fadeIn: `BubbleFadeInSeconds`
- hold: `BubbleLifetimeSeconds`
- fadeOut: `BubbleFadeOutSeconds`
- fadeOut 完了後に Canvas から remove
- 空いている slot を選ぶ

## スライド切替時の完全リセット

直近で追加した重要変更。

実装ファイル:

- `SlideAudienceAddIn\Overlay\OverlayWindow.xaml.cs`
- `SlideAudienceAddIn\Overlay\OverlayController.cs`
- `SlideAudienceAddIn\ThisAddIn.cs`

`OverlayWindow.xaml.cs` に追加:

- `ResetForNewSlide(string reason = "slideChanged")`
- `ClearAllComments(string reason)`
- `StopAllTimers()`
- `StopAllAnimations(string reason)`

リセット時に行うこと:

- `_displaySessionId++`
- panel hide timer 停止
- display timer 停止
- lifecycle timer 停止
- Flow / Bubble animation 停止
- `_commentQueue.Clear()`
- `_activeComments.Clear()`
- recent placement 情報 clear
- `AnimationCanvas.Children.Clear()`
- `CommentPanel.Children.Clear()`
- `PanelBorder.Visibility = Collapsed`
- reset time / removed count / active count / canvas children count を Debug log

`OverlayController.cs` に追加:

- `StartNewSlideSession()`
- `IsCurrentSession(int sessionId)`
- `ResetForNewSlide(string reason = "slideChanged")`
- `ShowComments(..., int slideSessionId)` overload

`ThisAddIn.cs` の順序:

1. `SlideShowBegin` / `SlideShowNextSlide` 検知
2. slide id / index を取得
3. `_overlayController.StartNewSlideSession()`
4. 既存コメント生成 cancellation を cancel
5. `_overlayController.ResetForNewSlide("slideChanged")`
6. slide export
7. whitespace analyze
8. text extract
9. comments generate
10. `IsCurrentSession(slideSessionId)` を確認
11. current session の場合だけ `ShowComments(..., slideSessionId)`

これにより、スライドAの Gemini / dummy fallback 結果がスライドBに後から表示されない。

## 次スライド1枚 preload

直近で追加した重要変更。

実装ファイル:

- `SlideAudienceAddIn\Models\SlideSnapshot.cs`
- `SlideAudienceAddIn\Models\SlidePreloadResult.cs`
- `SlideAudienceAddIn\Services\SlidePreloadService.cs`
- `SlideAudienceAddIn\ThisAddIn.cs`
- `SlideAudienceAddIn\Services\ExperimentLogger.cs`
- `SlideAudienceAddIn\Ribbon\SlideAudienceRibbon.cs`
- `SlideAudienceAddIn\Models\AppSettings.cs`
- `SlideAudienceAddIn\Config\appsettings.example.json`
- `SlideAudienceAddIn\Config\appsettings.local.json`
- `SlideAudienceAddIn\SlideAudienceAddIn.csproj`

方針:

- 案B: 次スライド1枚だけ preload
- 全スライド事前生成はしない
- `PreloadSlideCount` は設定として持つが、現状実装は1枚先読みのみ
- PowerPoint COM object は background task に渡さない
- background task に渡すのは `SlideSnapshot` の pure data のみ

`SlideSnapshot`:

- VSTO / PowerPoint thread 側で作る
- `Slide.Export`
- `SlideTextExtractor.ExtractText`
- slide id / slide index / presentation hash / cache key
- image path / extracted text / export time / text extract time

`SlidePreloadService`:

- `SlideSnapshot` を受け取る
- background task で以下だけ実行
  - 保存済み PNG の余白解析
  - Gemini API / dummy fallback によるコメント生成
  - cache 保存
- `PowerPoint.Slide` / `Shape` には触らない
- `Dictionary` + `lock` で thread-safe cache
- `_inflight` で同じ cache key の重複 preload を避ける
- cancellation 対応
- cache trim 対応

現在スライド表示時の流れ:

1. slide id / slide index を取得
2. `_overlayController.StartNewSlideSession()`
3. 現在生成 cancellation を cancel
4. `_overlayController.ResetForNewSlide("slideChanged")`
5. cache key を作成
6. preload cache hit なら即 `ShowComments`
7. cache miss なら従来通り export / text extract / whitespace analyze / comments generate
8. その後、次スライドの `SlideSnapshot` を VSTO thread で作成
9. background task で次スライド preload 開始

cache key:

- presentation hash
- slide id
- slide index
- comments settings hash
- whitespace placement settings hash
- `EnableApi`
- Gemini model

ログには raw cache key を出さない。`cacheKeyHash` のみ保存。

Preload 設定:

```json
"Preload": {
  "EnablePreloadNextSlide": true,
  "PreloadSlideCount": 1,
  "PreloadCacheMaxSlides": 5
}
```

Ribbon `Preload` group:

- `Preload Next` toggle
- `Clear Cache` button
- `Preload Count` dropdown: `1` / `2` / `3`

注意:

- `Preload Count` は UI と設定としては存在するが、現状は1枚だけ preload
- Enable OFF / SlideShowEnd / Clear Cache / preload に影響する設定変更で preload cache を clear
- preload 中に slideshow end しても cancellation される
- cache hit でも `OverlayController.IsCurrentSession(...)` を確認してから表示

## 余白解析

実装ファイル:

- `SlideAudienceAddIn\Services\SlideWhitespaceAnalyzer.cs`

モデル:

```csharp
public class WhitespaceRegion
{
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public double Score { get; set; }
}

public class SlideWhitespaceAnalysisResult
{
    public List<WhitespaceRegion> Regions { get; set; }
    public bool Succeeded { get; set; }
    public string FallbackReason { get; set; }
}
```

解析内容:

- PNG を `System.Drawing.Bitmap` で読み込み
- 解析用に最大幅 640px へ縮小
- デフォルト 12列 x 8行のグリッドで評価
- 各セルで以下を score 化
  - 白に近いピクセル割合
  - 明度分散の小ささ
  - エッジ量の少なさ
  - 暗い文字っぽいピクセルの少なさ
  - 外周 bonus
- OpenCV など外部依存なし
- 失敗時は従来配置へ fallback

Debug log:

- `Whitespace analysis started`
- original image width / height
- analysis bitmap width / height
- grid columns / rows
- detected region count
- each region x/y/width/height/score
- whitespace analysis time

## 余白優先配置

実装ファイル:

- `SlideAudienceAddIn\Overlay\OverlayWindow.xaml.cs`

Flow:

- 余白 region から lane top を選ぶ
- `WhitespacePlacementProbability` に従う
- region がない / 小さい / probability fallback の場合は固定 lane fallback
- 近い Y が続きすぎないよう `_recentFlowTops` で避ける

Bubble:

- 上位 whitespace region から重み付きランダム
- 最近使った region はペナルティ
- jitter: x ±40px, y ±30px
- 画面外へ出ないよう clamp
- 小さすぎる region は使わない
- fallback anchor:
  - `rightTop`
  - `rightMiddle`
  - `rightBottom`
  - `centerTop`
  - `centerBottom`
- fallback anchor も直前と違うものを優先
- 現在表示中 Bubble と重なりにくい位置を最大10回 retry

## コメント生成

実装ファイル:

- `SlideAudienceAddIn\Services\CommentGenerationService.cs`
- `SlideAudienceAddIn\Models\AudienceComment.cs`
- `SlideAudienceAddIn\Models\CommentGenerationResult.cs`

仕様:

- Gemini API で最大10件程度生成
- `Comments.MaxCommentsPerSlide` で制御
- 各コメントに `type`, `persona`, `text`
- dummy fallback も10件程度
- 最大18文字以内に正規化
- 句点「。」は削除
- APIキーは環境変数から読み、ログ・UI・設定に本文を出さない

代表 persona:

- `beginner`
- `expert`
- `skeptic`
- `curious`
- `practical`
- `research_evaluator`
- `designer`
- `experienced_speaker`
- `tsukkomi`
- `empathetic`

## JSONL logging

実装ファイル:

- `SlideAudienceAddIn\Services\ExperimentLogger.cs`

保存先:

```text
Documents/SlideAudience/logs/
```

ファイル名:

```text
session_yyyyMMdd_HHmmss.jsonl
```

イベント:

- `session_started`
- `session_ended`
- `slide_changed`
- `slide_analyzed`
- `comments_generated`
- `comments_shown`
- `comments_hidden`
- `overlay_mode_changed`
- `condition_preset_changed`
- `settings_changed`
- `error`

共通項目:

- `timestampUtc`
- `eventType`
- `sessionId`
- `presentationHash`
- `slideId`
- `slideIndex`
- `displayMode`
- `enableApi`
- `usedApi`
- `latencyMs`

プライバシー:

- プレゼンファイルパスの生文字列は保存しない
- `presentationHash` は SHA-256
- APIキーは保存しない
- `SaveSlideTextToLog=false` ではスライド本文テキストを保存しない
- ログ保存失敗でも PowerPoint を落とさない

追加済みの主な項目:

- `persona`
- `commentIndex`
- `displayOrder`
- `activeCountAtDisplay`
- `laneIndex`
- `slotIndex`
- `displayStartTime`
- `displayEndTime`
- `removedReason`
- `placementMode`
- `selectedRegionIndex`
- `fallbackAnchor`
- `whitespaceRegionCount`
- `selectedPlacementRegion`
- `commentDisplayIntervalMinSeconds`
- `commentDisplayIntervalMaxSeconds`
- `whitespacePlacementProbability`
- `commentFontSize`
- `commentTextColor`
- `commentBackgroundColor`
- `flowSpeedPixelsPerSecond`
- `bubbleFadeInSeconds`
- `bubbleFadeOutSeconds`
- `slideExportTimeMs`
- `whitespaceAnalysisTimeMs`

注意:

- `displayStartTime` / `displayEndTime` は現状では推定値寄り
- 実際の WPF animation completed を逐次 JSONL に戻す仕組みは未実装

## 現在の設定例

`SlideAudienceAddIn\Config\appsettings.example.json` と `appsettings.local.json` に以下を含む。

```json
{
  "Comments": {
    "Mode": "Mixed",
    "MaxCommentsPerSlide": 10,
    "MinCharacters": 10,
    "MaxCharacters": 18,
    "Language": "ja-JP"
  },
  "Overlay": {
    "DisplayMode": "Panel",
    "CommentLifetimeSeconds": 7,
    "FlowSpeedPixelsPerSecond": 80,
    "BubbleLifetimeSeconds": 9,
    "BubbleFadeSeconds": 1.0,
    "BubbleFadeInSeconds": 1.0,
    "BubbleFadeOutSeconds": 1.0,
    "CommentDisplayIntervalSeconds": 1.2,
    "CommentDisplayIntervalMinSeconds": 1.0,
    "CommentDisplayIntervalMaxSeconds": 2.5,
    "CommentFontSize": 28,
    "CommentTextColor": "#FFFFFFFF",
    "CommentBackgroundColor": "#99000000",
    "UseWhitespaceAwarePlacement": true,
    "WhitespacePlacementProbability": 0.7,
    "WhitespaceGridColumns": 12,
    "WhitespaceGridRows": 8,
    "MaxSimultaneousComments": 3
  },
  "Experiment": {
    "EnableLogging": true,
    "SaveSlideTextToLog": false,
    "SaveImagePathToLog": true
  }
}
```

## ビルド確認状況

Codex 環境で使用した MSBuild:

```text
C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
```

確認結果:

- C# コンパイルは通過
- `SlideAudienceAddIn\bin\Debug\SlideAudienceAddIn.dll` は生成される
- 最後の VSTO / ClickOnce 署名ターゲットで停止

停止エラー例:

```text
Cannot build because the ClickOnce manifest signing option is not selected
```

または:

```text
ResolveKeySource / X509Store.Open / Access denied
```

これは Codex sandbox からユーザー証明書ストアや ClickOnce 署名設定へアクセスできないことが原因。コードの C# コンパイルエラーではない。

## 次に確認すべきこと

1. Visual Studio で `SlideAudienceAddIn.sln` を開く
2. ClickOnce manifest signing 設定を確認
3. F5 で PowerPoint 起動
4. Ribbon に `SlideAudience` タブが出るか確認
5. `Panel` / `Flow` / `Bubble` を切り替え
6. Interval Min / Max を変更し、Flow / Bubble の出現間隔がランダムになるか確認
7. Font Size / Text Color / Background が3モードに反映されるか確認
8. Flow Speed が Flow に効くか確認
9. Bubble In / Bubble Out が fade に効くか確認
10. スライド切替直後に古いコメントが完全に消えるか確認
11. スライドA生成中にスライドBへ移動しても、AのコメントがBに出ないか確認
12. Clear Overlay で queue / timer / animation / Canvas children が消えるか確認
13. 長時間スライドショーで Canvas children count が増え続けないか確認
14. `Documents/SlideAudience/logs` に JSONL が生成されるか確認
15. APIキー本文が UI / 設定 / JSONL に出ていないか確認

## 注意点

- APIキー本文は絶対に保存・表示・ログ出力しない
- `SaveSlideTextToLog=false` ではスライド本文テキストを JSONL に保存しない
- `EnableApi=false` の dummy fallback を壊さない
- Gemini API 接続を壊さない
- Panel / Flow / Bubble を壊さない
- STA Dispatcher 設計を維持
- クリック透過を維持
- JSONL logging を維持
- appsettings.local.json が壊れていても PowerPoint を落とさない
