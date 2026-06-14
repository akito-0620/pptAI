# pptAI — SlideAudience

PowerPoint スライドショー中に、現在のスライドをもとに **「観客が思いそうな短いコメント」を AI でリアルタイム生成**し、透明なオーバーレイとして表示する Windows PowerPoint VSTO Add-in です。

![概念図](docs/concept.png)

---

## 概要

プレゼン発表者が聴衆の反応を想像しながら話せるよう、スライド切替のたびに Gemini API がコメントを生成してスライドショー画面の右下に表示します。PowerPoint ファイル自体は一切変更しません。

| 機能 | 説明 |
|---|---|
| スライドショーイベント検知 | 開始 / スライド切替 / 終了を自動検知 |
| AI コメント生成 | Google Gemini API でスライド画像＋テキストを解析 |
| WPF 透明オーバーレイ | クリック透過・Topmost・位置設定可能 |
| スライドキャッシュ | SlideID ごとにコメントをキャッシュし再生成を省略 |
| 事前生成 | 発表前に全スライド分を一括生成してキャッシュ |
| 実験ログ | JSONL 形式でスライド切替・生成・表示を記録 |
| オフラインフォールバック | API キー未設定時はダミーコメントを表示 |

---

## スクリーンショット

> *(実機確認後に追加予定)*

---

## アーキテクチャ

```
pptAI/
├── SlideAudience/              # ソリューション全体
│   ├── SlideAudience.sln
│   └── SlideAudienceAddIn/     # VSTO Add-in 本体 (C# / .NET Framework 4.8)
│       ├── Config/
│       │   └── appsettings.example.json
│       ├── Models/             # データモデル
│       ├── Overlay/            # WPF 透明オーバーレイ
│       ├── Ribbon/             # PowerPoint リボン UI
│       ├── Services/           # イベント・生成・キャッシュ・ログ
│       ├── Settings/           # 設定ウィンドウ (WPF)
│       └── Utils/
└── tools/python/               # 補助スクリプト (Gemini 検証・ログ集計)
```

---

## 必要環境

- Windows 10 / 11
- Microsoft PowerPoint (デスクトップ版)
- Visual Studio 2022
  - ワークロード: **Office/SharePoint 開発**
  - 個別コンポーネント: **.NET Framework 4.8 ターゲティングパック**
- Google Gemini API キー（省略可・ダミーコメントで動作確認可能）

---

## セットアップ

### 1. 設定ファイルのコピー

```text
SlideAudienceAddIn/Config/appsettings.example.json
```

を以下にコピーして編集します。

```text
SlideAudienceAddIn/Config/appsettings.local.json
```

> `appsettings.local.json` は `.gitignore` により除外されています。

### 2. Gemini API キーの設定（任意）

```powershell
setx GEMINI_API_KEY "your_api_key_here"
```

`appsettings.local.json` で `"EnableApi": true` に変更し、PowerPoint を再起動します。

API キーを設定しない場合は `"EnableApi": false` のままにしてください。ダミーコメントが表示されます。

### 3. ビルドと起動

1. `SlideAudience/SlideAudience.sln` を Visual Studio 2022 で開く
2. Office デバッグホストとして PowerPoint を選択
3. F5 でデバッグ起動
4. PowerPoint でファイルを開き、スライドショーを開始

---

## 設定一覧

`appsettings.example.json` の初期値:

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
    "FontSize": 24,
    "UseNiconicoFlow": false
  },
  "Experiment": {
    "EnableLogging": true,
    "SaveSlideTextToLog": false
  }
}
```

| キー | 説明 |
|---|---|
| `Comments.Mode` | `Mixed` / `UnderstandingOnly` / `InterestOnly` / `CriticalOnly` |
| `Overlay.Position` | `BottomRight` / `BottomLeft` / `TopRight` / `TopLeft` |
| `Experiment.SaveSlideTextToLog` | `true` にするとスライド本文をログに記録（デフォルト: false） |

---

## リボンから操作できること

PowerPoint のリボンに **SlideAudience** タブが追加されます。

| ボタン | 説明 |
|---|---|
| Enable / Disable | Add-in の有効・無効切替 |
| Generate Comments for Current Slide | 現在のスライドのコメントを今すぐ生成 |
| Pregenerate All Slides | 発表前に全スライド分を一括生成してキャッシュ |
| Clear Cache | キャッシュを削除して再生成を強制 |
| Open Settings | 設定ウィンドウを開く |

---

## ログとエクスポート

| 出力先 | 内容 |
|---|---|
| `Documents/SlideAudience/logs/` | JSONL 実験ログ |
| `%TEMP%/SlideAudience/exports/` | スライド PNG（一時ファイル） |

プレゼンファイルのパスは SHA-256 ハッシュとして記録されます。

---

## Python 補助ツール

```
tools/python/scripts/
├── test_gemini_image.py      # Gemini 画像理解の単体検証
├── pregenerate_comments.py   # PNG 群から事前コメント生成
└── analyze_logs.py           # JSONL ログの簡易集計
```

```bash
cd tools/python
pip install -r requirements.txt
python scripts/test_gemini_image.py
```

---

## 動作確認手順（MVP スモークテスト）

API なしで基本動作を確認する最短手順:

1. `"EnableApi": false` のままにする
2. Visual Studio からデバッグ起動
3. PowerPoint でスライドショーを開始
4. 次のスライドに進む
5. 右下に半透明のコメントパネルが表示されることを確認
6. スライドショーを終了してオーバーレイが消えることを確認

---

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| プロジェクトが VSTO として読み込まれない | Visual Studio Installer で `Office/SharePoint 開発` ワークロードを追加 |
| `.NET Framework 4.8` ターゲットエラー | Visual Studio Installer の個別コンポーネントから `.NET Framework 4.8 ターゲティングパック` を追加 |
| オーバーレイがクリックを吸収する | `OverlayWindow.xaml.cs` に `WS_EX_TRANSPARENT` の設定があるか確認 |
| Gemini がエラーを返す | `"EnableApi": false` に戻してダミーモードで継続確認 |

---

## ライセンス

MIT

---

## 参考リンク

- [Gemini API Reference](https://ai.google.dev/api)
- [Visual Studio VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins?view=vs-2022)
- [PowerPoint SlideShowNextSlide イベント](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshownextslide)
- [PowerPoint Slide.Export メソッド](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export)
