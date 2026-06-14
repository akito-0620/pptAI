# SlideAudience Implementation Status

作成日: 2026-05-02

## 概要

`ppt_ai_audience_comments_design.md` に基づき、Windows 版 PowerPoint 用の C# VSTO Add-in プロトタイプ `SlideAudience` を実装中です。

目的は、PowerPoint のスライドショー中に現在スライドをもとに「観客が思いそうな短いコメント」を生成し、PowerPoint 本体や `.pptx` を編集せず、透明な WPF オーバーレイとしてスライド上に表示することです。

現在は MVP の土台として、スライドショーイベント検知、スライド PNG 書き出し、スライド内テキスト抽出、ダミー/Gemini コメント生成、WPF オーバーレイ、キャッシュ、JSONL ログ、設定画面、リボン操作まで実装済みです。

## リポジトリ構成

```text
SlideAudience/
  SlideAudience.sln
  README.md
  IMPLEMENTATION_STATUS.md
  .gitignore
  SlideAudienceAddIn/
    SlideAudienceAddIn.csproj
    ThisAddIn.cs
    ThisAddIn.Designer.cs
    ThisAddIn.Designer.xml
    Config/
      appsettings.example.json
    Models/
      AppSettings.cs
      AudienceComment.cs
      CommentGenerationResult.cs
    Overlay/
      OverlayController.cs
      OverlayWindow.xaml
      OverlayWindow.xaml.cs
      CommentViewModel.cs
    Ribbon/
      SlideAudienceRibbon.cs
    Services/
      SlideShowEventService.cs
      SlideExporter.cs
      SlideTextExtractor.cs
      CommentGenerationService.cs
      CommentCache.cs
      ExperimentLogger.cs
    Settings/
      SettingsWindow.xaml
      SettingsWindow.xaml.cs
    Utils/
      JsonHelper.cs
      Win32WindowHelper.cs
    Properties/
      AssemblyInfo.cs
  tools/
    python/
      requirements.txt
      .env.example
      README.md
      scripts/
        test_gemini_image.py
        pregenerate_comments.py
        analyze_logs.py
```

## 実装済み機能

### 1. PowerPoint Add-in 基本構成

- `SlideAudience.sln` と `SlideAudienceAddIn.csproj` を作成。
- PowerPoint VSTO Add-in 想定の C# プロジェクト。
- `.NET Framework 4.8` をターゲットに設定。
- `ThisAddIn.cs` に起動・終了処理を実装。
- `ThisAddIn.Designer.cs` / `.xml` を追加し、VSTO の `ThisAddIn` 接着部を用意。

### 2. スライドショーイベント処理

実装ファイル:

- `SlideAudienceAddIn/Services/SlideShowEventService.cs`

実装内容:

- `Application.SlideShowBegin`
- `Application.SlideShowNextSlide`
- `Application.SlideShowEnd`

を購読し、スライドショー開始・スライド切替・終了を検知します。

スライド切替時の処理:

1. 現在の `Slide` を取得。
2. `SlideID` / `SlideIndex` を取得。
3. オーバーレイをスライドショーウィンドウ位置に合わせる。
4. `SlideID` キャッシュがあれば即表示。
5. キャッシュがなければローディング表示。
6. スライドを PNG に書き出し。
7. スライド内テキストを抽出。
8. コメント生成。
9. 結果をキャッシュ。
10. オーバーレイに表示。
11. JSONL ログ保存。

高速スライド切替対策として `CancellationTokenSource` を使い、古い生成結果が後から表示されないようにしています。

### 3. スライド PNG 書き出し

実装ファイル:

- `SlideAudienceAddIn/Services/SlideExporter.cs`

内容:

- `Slide.Export(filePath, "PNG", 1280, 720)` を使用。
- 出力先は `%TEMP%/SlideAudience/exports/`。
- ファイル名は `slide_{SlideID}_{timestamp}.png`。

### 4. スライド内テキスト抽出

実装ファイル:

- `SlideAudienceAddIn/Services/SlideTextExtractor.cs`

内容:

- Shape の TextFrame からテキスト抽出。
- GroupShape の再帰処理。
- Table セル内テキストの抽出。
- 一部 Shape が例外を投げても全体処理を止めない設計。

### 5. コメント生成

実装ファイル:

- `SlideAudienceAddIn/Services/CommentGenerationService.cs`

内容:

- Gemini API が無効、または `GEMINI_API_KEY` が未設定の場合はローカルのダミーコメントを返す。
- Gemini API が有効な場合は、スライド PNG と抽出テキストを送信して JSON 形式のコメントを取得する。
- API キーはソースに直書きせず、環境変数から読む。
- Gemini API キーは `x-goog-api-key` ヘッダーで送信。
- Gemini レスポンスの JSON fenced code block を除去する処理あり。
- JSON パース失敗や API エラー時は PowerPoint を落とさず、ダミーコメントへフォールバック。

対応コメントモード:

- `Mixed`
- `UnderstandingOnly`
- `InterestOnly`
- `CriticalOnly`

コメント種別:

- `understanding`
- `interest`
- `question`

### 6. コメントキャッシュ

実装ファイル:

- `SlideAudienceAddIn/Services/CommentCache.cs`

内容:

- `SlideID` をキーにコメントリストをキャッシュ。
- 同じスライドに戻った場合は再生成せず即表示。
- リボンからキャッシュ削除可能。

### 7. WPF オーバーレイ表示

実装ファイル:

- `SlideAudienceAddIn/Overlay/OverlayWindow.xaml`
- `SlideAudienceAddIn/Overlay/OverlayWindow.xaml.cs`
- `SlideAudienceAddIn/Overlay/OverlayController.cs`

内容:

- 透明・枠なし・タスクバー非表示・Topmost の WPF ウィンドウ。
- PowerPoint スライドショーウィンドウの HWND から位置とサイズを取得。
- 取得できない場合はプライマリ作業領域に表示。
- 右下・左下・右上・左上の表示位置を設定可能。
- `WS_EX_TRANSPARENT` を付与し、クリック透過に対応。
- コメントは半透明のパネルで表示。
- コメント種別ごとにアクセント色を変更。

### 8. リボン UI

実装ファイル:

- `SlideAudienceAddIn/Ribbon/SlideAudienceRibbon.cs`

リボンに追加済みの操作:

- `Enable SlideAudience`
- `Disable SlideAudience`
- `Generate Comments for Current Slide`
- `Pregenerate All Slides`
- `Clear Cache`
- `Open Settings`

`Pregenerate All Slides` は、発表前にアクティブなプレゼンテーション全スライドのコメントを生成し、`SlideID` キャッシュに入れます。

### 9. 設定画面

実装ファイル:

- `SlideAudienceAddIn/Settings/SettingsWindow.xaml`
- `SlideAudienceAddIn/Settings/SettingsWindow.xaml.cs`

設定可能な項目:

- Add-in 有効/無効
- Gemini API 有効/無効
- Gemini モデル名
- Gemini タイムアウト秒数
- コメントモード
- 最大コメント数
- オーバーレイ文字サイズ
- オーバーレイ表示位置
- JSONL ログ保存
- 抽出スライドテキストをログに含めるか

保存先:

```text
SlideAudienceAddIn/Config/appsettings.local.json
```

### 10. 実験ログ

実装ファイル:

- `SlideAudienceAddIn/Services/ExperimentLogger.cs`

保存先:

```text
Documents/SlideAudience/logs/
```

JSONL で記録するイベント:

- `slide_changed`
- `slide_analyzed`
- `comments_generated`
- `comments_shown`
- `error`

プライバシー配慮:

- デフォルトではスライド本文テキストをログ保存しない。
- `SaveSlideTextToLog` が true の場合のみ抽出テキストをログに含める。
- プレゼンファイルパスはそのまま保存せず、SHA-256 ハッシュ化する。

### 11. Python 補助ツール

実装場所:

```text
tools/python/
```

内容:

- `requirements.txt`
- `.env.example`
- `scripts/test_gemini_image.py`
- `scripts/pregenerate_comments.py`
- `scripts/analyze_logs.py`

用途:

- Gemini 画像理解の単体検証。
- PNG 群からの事前コメント生成。
- JSONL ログの簡易集計。

PowerPoint Add-in 本体は Python に依存しません。

## 設定ファイル

サンプル:

```text
SlideAudienceAddIn/Config/appsettings.example.json
```

初期値:

```json
{
  "Gemini": {
    "Model": "gemini-2.5-flash",
    "ApiKeyEnvironmentVariable": "GEMINI_API_KEY",
    "TimeoutSeconds": 20,
    "EnableApi": false
  },
  "Comments": {
    "Mode": "Mixed",
    "MaxCommentsPerSlide": 3,
    "MinCharacters": 15,
    "MaxCharacters": 35,
    "Language": "ja-JP"
  },
  "Overlay": {
    "Position": "BottomRight",
    "Width": 420,
    "MarginRight": 48,
    "MarginBottom": 48,
    "FontSize": 24,
    "UseNiconicoFlow": false
  },
  "Experiment": {
    "EnableLogging": true,
    "SaveSlideTextToLog": false
  }
}
```

Gemini API を使う場合:

1. `appsettings.local.json` で `"EnableApi": true` にする。
2. 環境変数 `GEMINI_API_KEY` を設定する。
3. PowerPoint を再起動する。

## 現在の確認状況

### 通った確認

- C# 中核ロジックのコンパイル確認は通過。
  - 対象: Models / Utils / Services の主要ロジック。
  - VSTO/WPF/XAML 全体ビルドではなく、コンパイラで確認できる範囲。
- Python 補助スクリプトは `py_compile` で構文確認済み。

### まだできていない確認

- Visual Studio での VSTO Add-in フルビルド。
- PowerPoint 実機でのロード確認。
- スライドショー上でのオーバーレイ表示確認。
- Gemini API 実通信確認。
- `.vsto` 配置・登録・発行確認。

## 現在の環境制約

この作業環境では、以下の理由でフルビルドできていません。

1. コマンドライン上から VSTO MSBuild targets が見つからない。
2. `.NET Framework 4.8 targeting pack` がコマンドライン上から確認できない。
3. MSBuild 実行時に以下で止まる。

```text
error MSB4184:
ToolLocationHelper.GetPlatformSDKLocation(Windows, 7.0) を評価できません。
パス 'C:\Users\akito\AppData\Local\Microsoft SDKs' へのアクセスが拒否されました。
```

そのため、次のステップとして Visual Studio Installer で以下を入れたうえで、Visual Studio から `SlideAudience.sln` を開いて確認する必要があります。

- Office/SharePoint development workload
- .NET Framework 4.8 targeting pack

## 重要な未解決リスク

### 1. VSTO プロジェクトの完全性

`ThisAddIn.Designer.cs` と `.xml` は作成済みですが、実際の Visual Studio VSTO テンプレート生成物と完全一致している保証はまだありません。

Visual Studio でロード時に問題が出る場合は、以下の方針が現実的です。

1. Visual Studio で新規 PowerPoint VSTO Add-in プロジェクトを作成。
2. 生成された `.csproj` / `ThisAddIn.Designer.cs` / `.vsto` 関連設定を正とする。
3. 現在の `Services`, `Overlay`, `Models`, `Settings`, `Ribbon`, `Utils` を移植。

### 2. XAML/WPF 全体ビルド未確認

WPF XAML を含むフルビルドは未確認です。XAML の名前解決や VSTO プロジェクト上での `Page` ビルド設定は Visual Studio で確認が必要です。

### 3. PowerPoint HWND 取得

`SlideShowWindow.HWND` からスライドショーウィンドウ位置を取得しています。環境や PowerPoint バージョンにより挙動確認が必要です。

### 4. オーバーレイのクリック透過

`WS_EX_TRANSPARENT`, `WS_EX_TOOLWINDOW`, `WS_EX_NOACTIVATE` を付与しています。スライドショー操作を邪魔しないか実機確認が必要です。

### 5. Gemini レスポンス形式

Gemini の REST API 形式は公式ドキュメントに合わせていますが、実通信は未確認です。`responseMimeType = application/json` とモデル名 `gemini-2.5-flash` の組み合わせで正常に JSON が返るか確認が必要です。

## 次に議論したいこと

生成AIに相談したい論点:

1. VSTO プロジェクトをこのまま完成させるべきか、Visual Studio で生成した正式テンプレートへ移植するべきか。
2. MVP の最短実機確認手順。
3. PowerPoint スライドショー上の WPF オーバーレイ実装で注意すべき点。
4. `Slide.Export()` をスライドショー中に毎回呼ぶ設計で問題ないか。
5. Gemini API の呼び出し形式、JSON 出力指定、画像送信形式に改善点があるか。
6. 発表中の遅延を避けるため、事前生成を MVP の中心にするべきか。
7. ログ項目として研究評価に足すべきものは何か。
8. 実験参加者へのプライバシー説明として不足している観点はあるか。
9. WPF ではなく WinForms / Direct2D / WebView2 overlay にするメリットがあるか。
10. 将来的に Office.js Add-in へ移行する可能性を考えるべきか。

## 次の具体タスク候補

優先順の候補:

1. Visual Studio で新規 PowerPoint VSTO Add-in を作成して、この実装を正式テンプレートへ移植する。
2. ダミーコメントのみで PowerPoint 実機起動テスト。
3. スライドショー中のオーバーレイ表示位置・クリック透過・終了時クローズ確認。
4. `Pregenerate All Slides` の動作確認。
5. Gemini API を ON にして単一スライド生成テスト。
6. ログ JSONL の実データ確認。
7. 設定画面の保存先が Add-in 実行環境で適切か確認。
8. エラー時 UX を調整。
9. 発表中に邪魔にならないオーバーレイデザインへ微調整。

## 参考リンク

- Gemini API reference: https://ai.google.dev/api
- Visual Studio VSTO Add-ins: https://learn.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins?view=vs-2022
- PowerPoint SlideShowNextSlide event: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshownextslide
- PowerPoint Slide.Export method: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export
