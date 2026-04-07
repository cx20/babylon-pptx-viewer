# PPTX Viewer - Babylon.js PowerPoint Simulator

3Dモニター上にPowerPointファイル(.pptx)をドラッグ&ドロップで読み込み、Babylon.js GUIでレンダリングするビューアです。

## デモ

Babylon.js Playground v2 で動作します。Playground v2 は複数ファイル・ESモジュールに対応しているため、各 `.js` ファイルをタブに追加してそのまま実行できます。

## ファイル構成

```
pptx-viewer/
├── README.md              ← このファイル
└── src/
    ├── index.js           ← エントリポイント（createScene, D&D, キーボード操作）
    ├── constants.js       ← 定数（EMU, キャンバスサイズ, 名前空間, デフォルトスライド）
    ├── color-utils.js     ← 色解決（テーマ色, schemeClr/srgbClr, shade/tint修飾子）
    ├── zip-helpers.js     ← ZIP操作（.rels解析, 画像データURL変換）
    ├── background.js      ← 背景抽出（画像/単色, duotone/アート効果検出）
    ├── text-parser.js     ← テキスト解析（段落, フォント, アウトライン, ジオメトリ）
    ├── style-inheritance.js ← スタイル継承（レイアウト/マスターからの継承チェーン）
    ├── shape-parsers.js   ← シェイプ解析（sp, pic, cxnSp, grpSp, graphicFrame, table）
    ├── slide-parser.js    ← スライドXML→要素配列変換
    ├── pptx-parser.js     ← メインオーケストレーター（ZIP展開→全スライド解析）
    ├── scene-setup.js     ← 3Dシーン構築（カメラ, ライト, PCモデル）
    ├── gui-frame.js       ← PowerPoint UIフレーム（タイトルバー, リボン, パネル）
    └── slide-renderer.js  ← スライド描画（メインキャンバス + サムネイル）
```

## アーキテクチャ

### データフロー

```
.pptx file (ZIP)
    │
    ▼
┌─────────────┐    ┌──────────────┐    ┌───────────────────┐
│ zip-helpers  │───▶│ color-utils  │───▶│ style-inheritance │
│ (ZIP/rels)   │    │ (theme)      │    │ (layout/master)   │
└─────────────┘    └──────────────┘    └───────────────────┘
    │                                          │
    ▼                                          ▼
┌─────────────┐    ┌──────────────┐    ┌───────────────────┐
│ background   │    │ text-parser  │───▶│ shape-parsers     │
│ (bg/effects) │    │ (paragraphs) │    │ (sp/pic/grp/gf)   │
└─────────────┘    └──────────────┘    └───────────────────┘
    │                                          │
    ▼                                          ▼
┌─────────────────────────────────────────────────────────┐
│                    pptx-parser.js                        │
│         (orchestrates all parsing, builds slides[])      │
└─────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────┐
│                   slide-renderer.js                      │
│         (renders slides to Babylon.js GUI canvas)        │
└─────────────────────────────────────────────────────────┘
```

### 共有状態 (`app` オブジェクト)

`index.js`で作成され、レンダリング系モジュールに渡されます：

```javascript
var app = {
    scene: scene,           // Babylon.jsシーン
    gui: { ... },           // GUIコントロール参照
    slides: [],             // パース済みスライドデータ
    currentSlide: 0,        // 現在表示中のスライドインデックス
    thumbRects: []          // サムネイル矩形（選択状態更新用）
};
```

### スライドデータモデル

各スライドは以下の構造を持ちます：

```javascript
{
    bg: "#FFFFFF",              // 背景色
    bgImage: "data:image/...", // 背景画像（dataURL or null）
    bgTint: {                   // 背景効果（null or object）
        type: "artEffect",      // "duotone" | "artEffect" | "tint" | "alpha"
        color: "#0E5580"        // ティント色
    },
    elements: [                 // スライド上の要素配列
        { type: "text", text: "...", x: 0.1, y: 0.2, w: 0.8, fontSize: 24, color: "#FFF", ... },
        { type: "shape", shape: "rect", x: 0.0, y: 0.0, w: 0.5, h: 0.5, fillColor: "#ACD433", ... },
        { type: "image", dataUrl: "data:...", x: 0.5, y: 0.0, w: 0.5, h: 1.0, ... },
    ],
    notes: "Slide 1"
}
```

座標系: `x`, `y`, `w`, `h` はすべて0.0〜1.0の小数（スライドサイズに対する割合）。

### 要素スキーマ (Issue 04)

すべての要素は `normalizeElement()` で正規化され、以下のスキーマに準拠します：

#### 共通プロパティ（すべての要素）

| プロパティ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `type` | string | "shape" | "text" \| "shape" \| "image" \| "table" |
| `x` | number | 0 | 左端座標（0.0-1.0） |
| `y` | number | 0 | 上端座標（0.0-1.0） |
| `rotation` | number | 0 | 回転角度（度） |

#### テキスト要素 (`type: "text"`)

| プロパティ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `text` | string | "" | テキスト内容 |
| `w` | number | 1 | 幅（0.0-1.0） |
| `fontSize` | number | 12 | ポイント |
| `color` | string | "#000000" | 16進カラーコード |
| `fontWeight` | string | "normal" | "normal" \| "bold" |
| `fontStyle` | string | "normal" | "normal" \| "italic" |
| `fontFamily` | string | "Calibri" | フォント名 |
| `align` | string | "left" | "left" \| "center" \| "right" |

#### 図形要素 (`type: "shape"`)

| プロパティ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `shape` | string | "rect" | "rect" \| "ellipse" \| "line" \| "circle" 等 |
| `w` | number | 1 | 幅（0.0-1.0） |
| `h` | number | 1 | 高さ（0.0-1.0） |
| `fillColor` | string | "#FFFFFF" | 塗りつぶし色 |
| `strokeColor` | string | "#000000" | 枠線色 |
| `thickness` | number | 1 | 線の太さ（ピクセル） |
| `x1`, `y1`, `x2`, `y2` | number | — | 線用：始点・終点座標（line shapeのみ） |

#### 画像要素 (`type: "image"`)

| プロパティ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `dataUrl` | string | null | 画像dataURL |
| `w` | number | 1 | 幅（0.0-1.0） |
| `h` | number | 1 | 高さ（0.0-1.0） |
| `crop` | object | — | `{l, t, r, b}` クロップ値（0.0-1.0） |

#### テーブル要素 (`type: "table"`)

| プロパティ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `rows` | number | 0 | 行数 |
| `cols` | number | 0 | 列数 |
| `tableData` | array | [] | セルデータ配列 |

**正規化の利点**:
- すべての要素が保証されたプロパティセットを持つため、レンダリング時に undefined チェックが不要
- パーサーはこの関数を使用して一貫性を確保
- 新しい要素タイプやプロパティの追加時も DEFAULT_ELEMENT_SCHEMA を更新するだけで完結

## 対応状況

### OOXML要素

| 要素 | 状態 | 備考 |
|------|------|------|
| `p:sp` (シェイプ) | ✅ | rect, ellipse, roundRect等 |
| `p:pic` (画像) | ✅ | srcRectクロップ対応 |
| `p:cxnSp` (コネクタ) | ✅ | flipH/flipV対応 |
| `p:grpSp` (グループ) | ✅ | 再帰的ネスト対応 |
| `p:graphicFrame` (表) | ⚠️ | 基本的なテーブルのみ |
| `p:graphicFrame` (チャート) | ⚠️ | プレースホルダー表示 |
| `p:graphicFrame` (ダイアグラム) | ⚠️ | プレースホルダー表示 |
| SmartArt | ❌ | 未対応 |
| アニメーション | ❌ | 未対応 |

### テキスト

| 機能 | 状態 | 備考 |
|------|------|------|
| フォントサイズ | ✅ | 0.75倍スケーリング |
| 太字/斜体 | ✅ | |
| 色（solidFill, schemeClr） | ✅ | shade/tint/lumMod修飾子対応 |
| アライメント | ✅ | left/center/right |
| 箇条書き（buChar） | ✅ | |
| 大文字変換（cap="all"） | ✅ | レイアウト継承対応 |
| CJK文字折り返し | ✅ | ゼロ幅スペース挿入 |
| 行間 | ⚠️ | 近似値 |

### 背景

| 機能 | 状態 | 備考 |
|------|------|------|
| 単色背景 | ✅ | |
| 画像背景 | ✅ | |
| グラデーション | ⚠️ | 最初の色で近似 |
| デュオトーン | ✅ | グレースケール検出→dk2ティント |
| アート効果 | ⚠️ | dk2色でオーバーレイ近似 |
| 背景継承（slide→layout→master） | ✅ | |

### スタイル継承

| レイヤー | 状態 | 備考 |
|----------|------|------|
| スライド自身のスタイル | ✅ | |
| レイアウトプレースホルダー | ✅ | cap, anchor, fontSize, color, fontRef |
| マスターtxStyles | ✅ | titleStyle, bodyStyle（非bgImageスライドのみ） |
| マスタープレースホルダーfontRef | ✅ | |
| テーマ色（dk1〜accent6） | ✅ | |

## 開発ガイド

### モジュールの修正

各ファイルは単一責務を持ちます。修正する際の指針：

- **新しいシェイプタイプを追加**: `shape-parsers.js` の `parseShapeTree` に分岐追加
- **テキスト書式を追加**: `text-parser.js` の `parseParagraphs` に属性読み取り追加
- **新しいスタイル継承**: `style-inheritance.js` に読み取りロジック追加
- **チャート/SmartArt対応**: `shape-parsers.js` の `parseGraphicFrame` を拡張
- **描画改善**: `slide-renderer.js` の `renderSlide` を修正
- **UI変更**: `gui-frame.js` を修正

### Playground での使用

Babylon.js Playground v2 は以下の機能をサポートしています：

- **複数ファイル** — VS Code スタイルのタブで各モジュールを別ファイルとして追加可能
- **ES Modules** — `import` / `export` をそのままブラウザ実行可能なモジュールに変換
- **NPM モジュール統合** — `jszip` などの外部ライブラリを直接インポート可能
- **TypeScript IntelliSense** — 自動型取得による補完・定義ジャンプ
- **Chrome DevTools** — ブレークポイント・ステップ実行などの高度なデバッグ
- **シェーダーサポート** — `.wgsl` / `.glsl` を別ファイルで管理

各 `.js` ファイルを Playground のタブに追加し、エントリポイントとして `index.js` を実行してください。ローカル開発は Live Server で `index.html` を開いて確認できます。

### デバッグ

コンソールログのプレフィックスで問題箇所を特定：

| プレフィックス | ファイル | 内容 |
|---------------|----------|------|
| `[PPTX]` | pptx-parser.js | 全体オーケストレーション |
| `[BG]` | background.js | 背景抽出 |
| `[BLIP]` | background.js | デュオトーン/アート効果 |
| `[TREE]` | shape-parsers.js | シェイプツリー走査 |
| `[SP]` | shape-parsers.js | 個別シェイプ解析 |
| `[PIC]` | shape-parsers.js | 画像要素 |
| `[GF]` | shape-parsers.js | graphicFrame |
| `[LAYOUT]` | style-inheritance.js | レイアウトスタイル |
| `[MASTER]` | style-inheritance.js | マスタースタイル |
| `[RENDER]` | slide-renderer.js | 描画処理 |

## 既知の制限事項

1. **背景アート効果**: PowerPointのアート効果（ぼかし、エッチング等）はピクセルレベル処理のためWebGLでの完全再現は困難。dk2色のオーバーレイで近似。
2. **チャート**: OOXMLチャートの描画は未実装。プレースホルダーボックスを表示。
3. **SmartArt**: ダイアグラム描画XML(`dsp:drawing`)の解析は未実装。
4. **アニメーション**: スライドアニメーション/トランジションは未対応。
5. **フォントサイズ**: Babylon.js GUIの制約により、元のサイズの75%でレンダリング。
6. **埋め込みフォント**: カスタムフォントは未対応。Segoe UI/Calibriにフォールバック。

## トラブルシューティング

### Live Server で起動失敗

#### ❌ "renderCanvas not found" / キャンバスが見つからない

**原因**: `index.html` の `<canvas id="renderCanvas">` が読み込み前に参照されている。

**対処**:
1. ページを完全リロード（Ctrl+Shift+R / Cmd+Shift+R）
2. Live Server の設定でディレイ読み込みを確認
3. `index.html` の `<script type="module" src="index.js"></script>` が `<body>` の最後にあるか確認

#### ❌ "Failed to load file library" / JSZip ロード失敗

**原因**: `libs/jszip.min.js` が見つからない、またはファイルが破損している。

**対処**:
1. `libs/jszip.min.js` が存在するか確認
2. ブラウザコンソール（F12）で `libs/jszip.min.js` の読み込みエラーを確認
3. ファイルが破損している場合は [jszip 3.10.1](https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js) を再ダウンロードして `libs/` に配置

#### ❌ "Failed to initialize graphics engine" / グラフィックスエンジン初期化失敗

**原因**: WebGL が無効、またはGPU/ドライバの非対応。

**対処**:
1. `about:gpu` (Chrome) または `about:preferences#advanced` (Firefox) でハードウェアアクセラレーション有効化
2. グラフィックスドライバを最新版に更新
3. 別のブラウザで試す（Chrome, Firefox, Edge）
4. シークレット/プライベートウィンドウで試す（拡張機能との競合を排除）

### PPTX 読み込み失敗

#### ❌ "PPTX parse error" / ファイル解析エラー

**原因**: ファイルが破損している、またはサポート外の形式。

**対処**:
1. ファイルが有効な .pptx ファイルか確認（ZIP形式で圧縮されている）
2. PowerPoint で一度開いて上書き保存し、修復
3. 別の .pptx ファイルで試す
4. データURL生成時のメモリ不足は大容量ファイルの場合あり

### コンソールで段階別エラー情報を確認

ブラウザコンソール（F12 → Console タブ）で以下の形式のログを確認できます：

```
[INIT] boot sequence start
[INIT/ENGINE] starting
[INIT/ENGINE] done
[INIT/SCENE] starting
[INIT/SCENE] done
[INIT/UI] starting
[INIT/UI] done
[INIT/INPUT] starting
[INIT/INPUT] done
[INIT] boot sequence complete
```

エラーが発生した場合：
- `[INIT] failed` と表示される
- `error code:` でエラー分類を確認
- `dev message:` で詳細情報を確認

### 既知の環境別問題

| 環境 | 症状 | 原因 | 対処 |
|------|------|------|------|
| Windows Safari | グラフィックスが表示されない | WebGL2未対応 | Chrome/Edge を使用 |
| iPad/iPhone | タッチで図形選択不可 | モバイルタッチイベント未対応 | テスト用にはPCブラウザを使用 |
| VPN/プロキシ | Babylon.js CDN読み込み失敗 | ホワイトリスト未設定 | IT部門に `cdn.babylonjs.com` をホワイトリスト登録依頼（JSZipはローカル動作） |
