# PowerPoint AI Audience Comments Add-in 設計書

## 0. この設計書の目的

このファイルは、Codex に渡して「PowerPoint スライドショー中に、スライドごとの AI 生成コメントを表示する専用プロトタイプ」を実装してもらうための設計書である。

既に画面監視型の汎用アプリは存在する前提とし、本プロジェクトでは PowerPoint 専用の連携型実装を行う。

---

## 1. プロジェクト概要

### 1.1 システム名

仮称: **SlideAudience**

### 1.2 コンセプト

PowerPoint のスライドショー中に、現在表示されているスライドをもとに AI が「観客が思いそうなコメント」を生成し、スライド上にオーバーレイ表示する。

コメントは発表者の説明を代弁するものではなく、擬似的な観客の内心として提示する。

例:

- 「つまり、発表中に観客の反応を可視化するってこと？」
- 「これは面白そう。でも邪魔にならないのかな？」
- 「発表者ではなく観客目線なのがポイントか」
- 「たしかに、聴衆の疑問を先回りできそう」

### 1.3 研究上の位置づけ

本システムは、プレゼンテーション視聴中の聴衆に対して、AI が生成した擬似聴衆コメントを提示することで、以下に与える影響を調べるための HCI プロトタイプである。

- 聴衆の興味
- 内容理解
- 疑問形成
- プレゼンへの親近感
- 注意散漫さ
- 発表者への印象

---

## 2. 実装方針

### 2.1 採用する方式

**Windows 版 PowerPoint + C# VSTO Add-in + WPF 透明オーバーレイ**で実装する。

理由:

- PowerPoint のスライドショーイベントを直接取得できる
- 画面差分ではなく、現在のスライド番号・SlideID を正確に取得できる
- `Slide.Export()` により現在スライドを画像として保存できる
- PowerPoint ファイル本体を編集せず、透明ウィンドウでコメントを重ねられる
- 研究実験用に「コメントあり/なし」「コメント種類」「コメント数」を切り替えやすい

### 2.2 MVP のゴール

MVP では以下を実現する。

1. PowerPoint スライドショー開始中に Add-in が動作する
2. スライドが切り替わったらイベントを検知する
3. 現在スライドを PNG として一時保存する
4. Gemini API にスライド画像とプロンプトを送信する
5. Gemini から JSON 形式でコメントを受け取る
6. 透明オーバーレイウィンドウ上にコメントを表示する
7. スライドが変わったら前のコメントを消し、新しいコメントを表示する
8. スライドショー終了時にオーバーレイを閉じる

### 2.3 MVP ではやらないこと

最初のバージョンでは以下は対象外とする。

- PowerPoint ファイルに直接コメント Shape を挿入すること
- 複数人のリアルタイム反応取得
- 音声認識による発話内容の解析
- Web 配信連携
- Mac 版 PowerPoint 対応
- PowerPoint Online 対応

---

## 3. 開発環境

### 3.1 対象環境

- OS: Windows 10 / Windows 11
- PowerPoint: Microsoft PowerPoint デスクトップ版
- IDE: Visual Studio 2022
- 言語: C#
- プロジェクト種別: PowerPoint VSTO Add-in
- .NET: .NET Framework 4.8 推奨

### 3.2 必要な外部 API

- Gemini API
- API キーは環境変数 `GEMINI_API_KEY` から読む
- モデル名は設定ファイルから変更可能にする

例:

```text
GEMINI_MODEL=gemini-2.5-flash
```

モデル名は固定せず、`appsettings.json` または簡易設定クラスから変更できるようにする。


### 3.3 Python を使う場合の仮想環境ルール

本体は **C# VSTO Add-in** として実装するため、通常は Python は必須ではない。
ただし、Codex が以下の用途で Python を使う場合は、必ずプロジェクト内に仮想環境を作成してから実装する。

Python を使ってよい用途:

- Gemini API 接続の検証用スクリプト
- スライド画像からコメント生成を試す CLI ツール
- 事前生成モード用のバッチ処理
- 実験ログ JSONL / CSV の分析・整形
- 将来的に C# Add-in から呼び出すローカル補助サーバー

Python を使う場合の必須ルール:

1. グローバル環境に直接インストールしない
2. `tools/python/` 以下に Python 関連コードを置く
3. `tools/python/.venv/` を作成する
4. 依存関係は `tools/python/requirements.txt` に固定する
5. API キーは `.env` または環境変数で管理し、Git に含めない
6. `.venv/`, `.env`, `__pycache__/` は `.gitignore` に入れる

Windows PowerShell でのセットアップ例:

```powershell
cd tools/python
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

`Activate.ps1` が実行できない場合は、PowerShell を管理者ではなく通常モードで開き、以下を実行する。

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

`tools/python/requirements.txt` の初期案:

```text
google-genai
python-dotenv
pillow
pydantic
```

`tools/python/.env.example` の初期案:

```text
GEMINI_API_KEY=your_api_key_here
GEMINI_MODEL=gemini-2.5-flash
```

Python 検証スクリプトの想定:

```text
tools/python/
  .venv/                    # Git管理しない
  requirements.txt
  .env.example
  README.md
  scripts/
    test_gemini_image.py     # 画像1枚からコメント生成を試す
    pregenerate_comments.py  # スライド画像群から事前生成
    analyze_logs.py          # JSONLログ分析
```

C# 側から Python を直接呼び出す実装は MVP では避ける。
MVP では C# から Gemini API を直接呼ぶ。Python は検証・事前生成・ログ分析の補助用途に限定する。

---

## 4. 全体アーキテクチャ

```text
PowerPoint VSTO Add-in
  |
  |-- SlideShowEventService
  |     |-- SlideShowNextSlide を検知
  |     |-- 現在 Slide を取得
  |
  |-- SlideExporter
  |     |-- Slide.Export() で PNG 保存
  |
  |-- SlideTextExtractor
  |     |-- スライド内テキストを抽出
  |
  |-- CommentGenerationService
  |     |-- Gemini API へ画像 + テキストを送信
  |     |-- JSON をパース
  |
  |-- CommentCache
  |     |-- SlideID ごとにコメントをキャッシュ
  |
  |-- OverlayWindow / OverlayController
  |     |-- PowerPoint スライドショー画面上に透明 WPF ウィンドウ表示
  |     |-- コメントを表示
  |
  |-- ExperimentLogger
        |-- スライドID、時刻、生成コメント、表示条件をCSV/JSONL保存
```

---

## 5. 主要コンポーネント仕様

## 5.1 ThisAddIn

### 役割

VSTO Add-in のエントリーポイント。PowerPoint のイベント登録と終了処理を行う。

### 実装内容

- `Application.SlideShowBegin` を購読
- `Application.SlideShowNextSlide` を購読
- `Application.SlideShowEnd` を購読
- 必要なサービスを初期化

### 疑似コード

```csharp
private SlideShowEventService _slideShowEventService;
private OverlayController _overlayController;
private CommentGenerationService _commentGenerationService;

private void ThisAddIn_Startup(object sender, EventArgs e)
{
    _overlayController = new OverlayController();
    _commentGenerationService = new CommentGenerationService();
    _slideShowEventService = new SlideShowEventService(
        this.Application,
        _overlayController,
        _commentGenerationService
    );

    _slideShowEventService.RegisterEvents();
}

private void ThisAddIn_Shutdown(object sender, EventArgs e)
{
    _slideShowEventService?.UnregisterEvents();
    _overlayController?.Close();
}
```

---

## 5.2 SlideShowEventService

### 役割

PowerPoint のスライドショーイベントを受け取り、スライド変更時の処理を開始する。

### 重要イベント

- `SlideShowBegin`
- `SlideShowNextSlide`
- `SlideShowEnd`

### スライド切り替え時の処理

1. `Wn.View.Slide` から現在スライドを取得
2. `SlideID` と `SlideIndex` を取得
3. オーバーレイ位置をスライドショーウィンドウに合わせる
4. キャッシュがあれば即表示
5. キャッシュがなければローディング表示
6. スライド画像を書き出す
7. Gemini API でコメント生成
8. 結果を表示
9. ログ保存

### 注意点

- イベントが連続して呼ばれる可能性があるため、`CancellationTokenSource` を使う
- スライドがすぐ次へ進んだ場合、古いスライドのコメントを表示しない
- SlideID を基準に、現在表示中のスライドと生成結果のスライドが一致するか確認する

### 疑似コード

```csharp
private CancellationTokenSource _currentCts;
private int _currentSlideId;

private async void OnSlideShowNextSlide(PowerPoint.SlideShowWindow wn)
{
    _currentCts?.Cancel();
    _currentCts = new CancellationTokenSource();
    var token = _currentCts.Token;

    var slide = wn.View.Slide;
    int slideId = slide.SlideID;
    int slideIndex = slide.SlideIndex;
    _currentSlideId = slideId;

    _overlayController.AttachToSlideShowWindow(wn);
    _overlayController.ShowLoading(slideIndex);

    try
    {
        var comments = await _commentGenerationService.GenerateForSlideAsync(slide, slideIndex, token);

        if (token.IsCancellationRequested) return;
        if (_currentSlideId != slideId) return;

        _overlayController.ShowComments(comments);
        _logger.Log(slideId, slideIndex, comments);
    }
    catch (Exception ex)
    {
        _overlayController.ShowError("コメント生成に失敗しました");
        _logger.LogError(slideId, slideIndex, ex);
    }
}
```

---

## 5.3 SlideExporter

### 役割

現在スライドを PNG 画像として一時フォルダへ保存する。

### 仕様

- `Slide.Export(filePath, "PNG", width, height)` を使う
- 推奨解像度: 1280 x 720 または 1920 x 1080
- 保存先: `%TEMP%/SlideAudience/exports/`
- ファイル名: `slide_{SlideID}_{timestamp}.png`

### 実装例

```csharp
public string ExportSlideAsPng(PowerPoint.Slide slide, int width = 1280, int height = 720)
{
    string dir = Path.Combine(Path.GetTempPath(), "SlideAudience", "exports");
    Directory.CreateDirectory(dir);

    string filePath = Path.Combine(
        dir,
        $"slide_{slide.SlideID}_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png"
    );

    slide.Export(filePath, "PNG", width, height);
    return filePath;
}
```

---

## 5.4 SlideTextExtractor

### 役割

スライド内のテキストを抽出し、Gemini への補助情報として送る。

### 理由

画像だけでもコメント生成は可能だが、スライド画像内の小さい文字や日本語の認識が不安定な場合がある。PowerPoint の Shape から直接テキストを抽出することで精度を上げる。

### 抽出対象

- タイトル
- 箇条書き
- 図中のラベル
- テキストボックス
- 表内テキスト

### 実装方針

Shape を再帰的に走査する。

```csharp
public string ExtractText(PowerPoint.Slide slide)
{
    var sb = new StringBuilder();

    foreach (PowerPoint.Shape shape in slide.Shapes)
    {
        ExtractShapeText(shape, sb);
    }

    return sb.ToString();
}

private void ExtractShapeText(PowerPoint.Shape shape, StringBuilder sb)
{
    try
    {
        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
            shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
        {
            var text = shape.TextFrame.TextRange.Text;
            if (!string.IsNullOrWhiteSpace(text))
            {
                sb.AppendLine(text.Trim());
            }
        }

        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
        {
            foreach (PowerPoint.Shape child in shape.GroupItems)
            {
                ExtractShapeText(child, sb);
            }
        }
    }
    catch
    {
        // 一部Shapeで例外が出ても全体処理を止めない
    }
}
```

---

## 5.5 CommentGenerationService

### 役割

スライド画像とテキストを Gemini API に送信し、観客コメントを生成する。

### 入力

- slideIndex
- slideId
- slideImagePath
- extractedText
- previousSlideSummary optional
- mode
  - `AudienceLike`
  - `UnderstandingOnly`
  - `CriticalOnly`
  - `Mixed`

### 出力

```csharp
public class AudienceComment
{
    public string Type { get; set; }
    public string Text { get; set; }
    public double? Confidence { get; set; }
}
```

### JSON 出力形式

Gemini には以下の JSON 形式で返すように指示する。

```json
{
  "comments": [
    {"type": "understanding", "text": "つまり、観客の内心を見せる仕組み？"},
    {"type": "interest", "text": "これは発表が少し配信っぽくなるね"},
    {"type": "question", "text": "でもコメントが多いと邪魔にならない？"}
  ]
}
```

### プロンプト

```text
あなたはプレゼンテーションを見ている観客です。
入力されたスライド画像とスライド内テキストを見て、観客が自然に思いそうな短いコメントを生成してください。

目的:
- 発表者の説明を繰り返すのではなく、観客の内心を代弁する
- 聴衆の理解、興味、疑問形成を助ける
- プレゼンの空気を少し柔らかくする

コメントの種類:
1. understanding: 「つまり〇〇ということ？」のような理解確認
2. interest: 「〇〇が面白そう」のような興味・感想
3. question: 「でも〇〇が難しくない？」のような疑問・批判

制約:
- 各コメントは15〜35文字程度
- 日本語で書く
- 攻撃的すぎる表現は避ける
- 発表者の代弁ではなく、観客の反応にする
- 1スライドあたり最大3コメント
- 事実を断定しすぎない
- スライドにない内容を勝手に広げすぎない

出力は必ずJSONのみ:
{
  "comments": [
    {"type": "understanding", "text": "..."},
    {"type": "interest", "text": "..."},
    {"type": "question", "text": "..."}
  ]
}
```

### API キー管理

- `GEMINI_API_KEY` を環境変数から取得する
- API キーをソースコードに直接書かない
- `.gitignore` に設定ファイルやログを入れる

---

## 5.6 CommentCache

### 役割

同じスライドに戻ったとき、再度 API を呼ばずにコメントを表示する。

### キー

- `Presentation.FullName + SlideID + Mode + PromptVersion`

### 仕様

- メモリキャッシュを MVP で実装
- 発展版では JSON ファイルに保存
- スライドが編集された場合の検出は MVP では不要

---

## 5.7 OverlayWindow

### 役割

PowerPoint スライドショー画面の上に透明ウィンドウを表示し、AI コメントを重ねる。

### 技術

- WPF Window
- `WindowStyle=None`
- `AllowsTransparency=True`
- `Background=Transparent`
- `Topmost=True`
- `ShowInTaskbar=False`

### 表示位置

MVP では右下に縦並びで表示する。

```text
+------------------------------------------------+
|                                                |
|                 PowerPoint Slide               |
|                                                |
|                         [コメント1]             |
|                         [コメント2]             |
|                         [コメント3]             |
+------------------------------------------------+
```

### コメントデザイン

- 半透明の黒背景
- 白文字
- 角丸吹き出し
- 1コメント最大1行または2行
- フェードイン
- 文字サイズは 20〜28 px 程度

### XAML イメージ

```xml
<Window x:Class="SlideAudience.OverlayWindow"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="False">
    <Grid IsHitTestVisible="False">
        <StackPanel x:Name="CommentPanel"
                    HorizontalAlignment="Right"
                    VerticalAlignment="Bottom"
                    Margin="0,0,48,48"
                    Width="420">
        </StackPanel>
    </Grid>
</Window>
```

### クリック透過

MVP では `IsHitTestVisible=False` で十分。必要なら Win32 API でクリック透過を追加する。

---

## 5.8 OverlayController

### 役割

オーバーレイウィンドウの生成、位置調整、表示更新を管理する。

### 主なメソッド

```csharp
public void AttachToSlideShowWindow(PowerPoint.SlideShowWindow wn);
public void ShowLoading(int slideIndex);
public void ShowComments(IEnumerable<AudienceComment> comments);
public void ShowError(string message);
public void Hide();
public void Close();
```

### スライドショーウィンドウへの追従

PowerPoint の SlideShowWindow の HWND からウィンドウ矩形を取得し、OverlayWindow の位置とサイズを合わせる。

```csharp
[DllImport("user32.dll")]
private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
```

注意:

- PowerPoint の Interop で `wn.HWND` が使えるか確認する
- 使えない場合は、MVP ではプライマリモニタ全体に表示してもよい
- 研究用途では「発表モニタを指定できる設定」を追加するとよい

---

## 5.9 ExperimentLogger

### 役割

研究実験用に表示ログを保存する。

### 保存先

`Documents/SlideAudience/logs/`

### ログ形式

JSONL 推奨。

1行1イベントで保存する。

```json
{"timestamp":"2026-05-02T12:00:00.123+09:00","event":"slide_changed","slideIndex":3,"slideId":256}
{"timestamp":"2026-05-02T12:00:02.456+09:00","event":"comments_shown","slideIndex":3,"slideId":256,"comments":[{"type":"understanding","text":"つまり、観客の反応を作る？"}]}
```

### 記録する内容

- timestamp
- event type
- presentation file path hash
- slideIndex
- slideId
- mode
- generated comments
- API latency
- error message

プライバシー保護のため、デフォルトではスライド本文全文はログに保存しない。

---

## 6. UI 仕様

## 6.1 Ribbon UI

PowerPoint のリボンに簡単な操作パネルを追加する。

### ボタン

- Enable SlideAudience
- Disable SlideAudience
- Generate Comments for Current Slide
- Clear Cache
- Open Settings

### 設定項目

- コメント表示 ON/OFF
- API 使用 ON/OFF
- コメント数: 1 / 2 / 3
- コメントモード
  - Mixed
  - Understanding only
  - Interest only
  - Critical question only
- 表示位置
  - 右下
  - 左下
  - 上部
  - ニコニコ風に横流し
- 生成方式
  - リアルタイム生成
  - 事前生成 / キャッシュ優先
  - ダミーコメント

---

## 7. 実装ステップ

### Step 1: VSTO Add-in のひな形

- PowerPoint VSTO Add-in プロジェクトを作る
- 起動時にリボンボタンを表示
- スライドショーイベントをログ出力する

完了条件:

- PowerPoint 起動時に Add-in が読み込まれる
- スライドショー中、スライドが変わるたびに slideIndex / slideId がログに出る

---

### Step 2: スライド画像のエクスポート

- SlideExporter を実装
- スライド遷移時に PNG を保存

完了条件:

- スライドを切り替えるたびに `%TEMP%/SlideAudience/exports/` に PNG が生成される

---

### Step 3: オーバーレイ表示

- WPF OverlayWindow を作る
- スライドショー中に透明ウィンドウを表示
- ダミーコメントを表示

完了条件:

- スライドショー画面上にコメントが重なる
- PowerPoint のスライド本体は編集されない
- スライドショー終了時にオーバーレイが閉じる

---

### Step 4: Gemini API 接続

- CommentGenerationService を実装
- API キーを環境変数から読む
- 画像 + プロンプトを送信
- JSON をパースしてコメント表示

完了条件:

- スライド画像から観客コメントが生成される
- 生成失敗時はエラー表示になり、アプリは落ちない

---

### Step 5: キャッシュとキャンセル処理

- SlideID ごとのキャッシュを実装
- スライド高速切り替え時に古いコメントが表示されないようにする
- `CancellationToken` を導入

完了条件:

- 同じスライドに戻ったとき即表示される
- 1つ前のスライドの生成結果が遅れて表示されない

---

### Step 6: 実験用ログ

- JSONL ログを保存
- API latency を測定
- 表示コメントを記録

完了条件:

- 実験後に「どのスライドで何が表示されたか」を確認できる

---

## 8. ディレクトリ構成案

```text
SlideAudience/
  SlideAudience.sln
  SlideAudienceAddIn/
    ThisAddIn.cs
    Ribbon/
      SlideAudienceRibbon.cs
      SlideAudienceRibbon.xml
    Services/
      SlideShowEventService.cs
      SlideExporter.cs
      SlideTextExtractor.cs
      CommentGenerationService.cs
      CommentCache.cs
      ExperimentLogger.cs
    Overlay/
      OverlayWindow.xaml
      OverlayWindow.xaml.cs
      OverlayController.cs
      CommentViewModel.cs
    Models/
      AudienceComment.cs
      CommentGenerationResult.cs
      AppSettings.cs
    Utils/
      Win32WindowHelper.cs
      JsonHelper.cs
    Config/
      appsettings.example.json
    README.md
    .gitignore
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

---

## 9. 設定ファイル例

`appsettings.example.json`

```json
{
  "Gemini": {
    "Model": "gemini-2.5-flash",
    "ApiKeyEnvironmentVariable": "GEMINI_API_KEY",
    "TimeoutSeconds": 20
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

---

## 9.1 .gitignore の必須項目

`.gitignore` には最低限以下を含める。

```gitignore
# Visual Studio / build outputs
.vs/
bin/
obj/
*.user
*.suo

# VSTO / Office temporary files
*.vsto
*.manifest

# Logs and temporary exports
logs/
exports/
*.jsonl
*.log

# Python virtual environment and secrets
tools/python/.venv/
tools/python/.env
**/__pycache__/
*.pyc

# Generic secrets
.env
appsettings.local.json
```

---

## 10. エラーハンドリング

### 想定エラー

- API キー未設定
- Gemini API への通信失敗
- JSON パース失敗
- Slide.Export 失敗
- PowerPoint スライドショーウィンドウが取得できない
- オーバーレイ表示位置の取得失敗

### 方針

- PowerPoint 本体を落とさない
- 例外はすべてログに残す
- 画面には短いエラーだけ表示
- API 失敗時はダミーコメントまたはコメントなしにフォールバック

---

## 11. セキュリティ・プライバシー

- スライド画像を外部 API に送信するため、機密資料では使わない前提にする
- 研究参加者には、スライド画像が AI API に送信されることを説明する
- API キーは環境変数で管理し、Git に含めない
- ログにスライド本文を保存しない設定をデフォルトにする
- 生成コメント、時刻、スライド番号、遅延時間のみ保存する

---

## 12. テスト項目

### 12.1 機能テスト

- Add-in が PowerPoint 起動時に読み込まれる
- スライドショー開始を検知できる
- スライド遷移を検知できる
- スライド終了を検知できる
- PNG エクスポートができる
- Gemini API からコメントが返る
- JSON パースができる
- オーバーレイが表示される
- スライドショー終了時にオーバーレイが閉じる

### 12.2 異常系テスト

- API キーがない
- ネット接続がない
- JSON が壊れている
- スライドを高速で切り替える
- スライドショーを途中で終了する
- 複数モニタ環境で動かす

### 12.3 研究用テスト

- コメントなし条件
- ダミー固定コメント条件
- AI 生成コメント条件
- コメント1個条件
- コメント3個条件
- 疑問コメントのみ条件
- 共感コメントのみ条件

---

## 13. Codex への実装指示文

以下を Codex にそのまま渡す。

```text
Windows版PowerPoint用のC# VSTO Add-inを実装してください。
目的は、PowerPointのスライドショー中にスライド遷移を検知し、現在スライドからAI生成の観客コメントを作り、透明WPFオーバーレイで表示することです。

このリポジトリに以下の構成で実装してください。

1. PowerPoint VSTO Add-in プロジェクトを作成する
2. ThisAddIn.cs で SlideShowBegin / SlideShowNextSlide / SlideShowEnd を購読する
3. SlideShowNextSlide で現在の Slide, SlideID, SlideIndex を取得する
4. SlideExporter.cs を作り、Slide.Export(path, "PNG", 1280, 720) で現在スライドをPNG保存する
5. SlideTextExtractor.cs を作り、スライド内のShapeからテキストを抽出する
6. CommentGenerationService.cs を作り、Gemini APIへ画像とテキストを送り、JSON形式でコメントを受け取る
7. Gemini APIキーは環境変数 GEMINI_API_KEY から読む。ソースコードに直書きしない
8. AudienceComment.cs と CommentGenerationResult.cs を定義する
9. OverlayWindow.xaml を作り、透明・最前面・タスクバー非表示のWPFウィンドウとしてコメントを右下に表示する
10. OverlayController.cs を作り、PowerPointのスライドショーウィンドウ上にOverlayWindowを表示・更新・終了できるようにする
11. スライド高速切り替え対策として CancellationToken を使い、古いAPI結果を表示しない
12. CommentCache.cs を作り、SlideIDごとに生成済みコメントをキャッシュする
13. ExperimentLogger.cs を作り、slide_changed / comments_generated / comments_shown / error をJSONLで保存する
14. PowerPoint本体や.pptxファイルを編集しない。コメントは必ずオーバーレイ表示にする
15. API失敗時もPowerPointが落ちないように、例外処理とログ保存を入れる
16. README.md にセットアップ手順、環境変数設定、Visual Studioでの実行方法を書く
17. Python を使う補助ツールを作る場合は、必ず tools/python/.venv を作成し、requirements.txt と .env.example を用意する。グローバル環境に pip install しない
18. Python 補助ツールは MVP では必須ではない。使う場合も、Gemini API 検証、事前生成、ログ分析などの補助用途に限定する

まずMVPとして、スライド遷移検知、PNGエクスポート、ダミーコメント表示、スライドショー終了時のオーバーレイ終了までを実装してください。
その後、Gemini API接続を実装してください。
Python を使う場合は、最初に tools/python/.venv を作成してから作業してください。
```

---

## 14. 実装優先順位

最初から Gemini API まで一気に作らず、以下の順番で進める。

1. スライド遷移検知
2. ダミーコメントのオーバーレイ表示
3. Slide.Export による画像保存
4. スライド内テキスト抽出
5. Gemini API 接続
6. キャッシュ
7. 実験ログ
8. リボンUI
9. 表示スタイル調整
10. Python 仮想環境つき補助ツールの整備 必要な場合のみ
11. 事前生成モード

---

## 15. 将来拡張

### 15.1 事前生成モード

発表前に全スライドを解析し、コメントを事前生成する。

メリット:

- 発表中の遅延がない
- 実験条件を固定しやすい
- API失敗リスクを減らせる

### 15.2 ニコニコ風流しコメント

コメントを右から左に流す表示モードを追加する。

ただし、研究評価では注意散漫になりやすいため、最初は右下固定表示を推奨する。

### 15.3 コメント制限

1スライドあたりのコメント数を制限する。

候補:

- 1コメント: 最小限の注意誘導
- 2コメント: 共感 + 疑問
- 3コメント: 理解確認 + 興味 + 疑問

### 15.4 発表者ノート活用

スライド本文だけでなく発表者ノートを読み、より文脈に合ったコメントを生成する。

注意:

- 発表者ノートには外部送信したくない情報が含まれる可能性がある
- 設定でON/OFFできるようにする

---

## 16. 最初の成功条件

最初の成功条件は以下。

```text
PowerPointでスライドショーを開始する
↓
スライドを次に進める
↓
右下に「つまり、このスライドのポイントは〇〇？」のようなダミーコメントが表示される
↓
次のスライドに進むとコメントが切り替わる
↓
スライドショーを終了するとコメントウィンドウも閉じる
```

ここまでできれば、PowerPoint専用プロトタイプとして成立する。
その後、Gemini API でコメント生成を動的化する。

---

## 17. 参考ドキュメント

- Codex Help: https://help.openai.com/en/articles/11369540-using-codex-with-your-chatgpt-plan
- PowerPoint SlideShowNextSlide event: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshownextslide
- PowerPoint Slide.Export method: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export
- Gemini Generate Content API: https://ai.google.dev/api/generate-content
- Gemini Image Understanding: https://ai.google.dev/gemini-api/docs/image-understanding
```
