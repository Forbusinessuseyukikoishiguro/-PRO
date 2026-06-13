# ラベンダー研修スライド・ジェネレーター（アーティファクト版）
# 要件定義・詳細設計書 — 完全再現用

| 項目 | 内容 |
|---|---|
| 製品名 | ラベンダー研修スライド・ジェネレーター（アーティファクト版） |
| ブランド | うさうさ研修工房 🐰🪻 |
| 版 | **v1.0.0 PRO** |
| 日付 | **2026-06-13** |
| 配布形態 | 単一HTMLファイル（claude.ai アーティファクト想定／内蔵Claudeで生成） |
| 想定ファイル名 | `lavender_slide_studio_app_v1_0_0_pro_20260613.html` |
| 本書の目的 | 本書のみを参照して、同一仕様のツールをゼロから完全再現できること |

> 本書は「事実（仕様・実装値）」を主体とし、設計判断の背景は【所感】として明示します。関係者の個人名は記載しません（必要時は役職名）。

---

## 1. 概要と目的

### 1.1 プロダクト概要
研修テーマと条件を入力すると、新卒エンジニア研修ワークショップの **スライド構成案**（表紙〜まとめ、講師ノート・演習・所要分つき）を生成し、`.pptx` / PDF / Markdown / 講師メモ の4形式に書き出せる単一HTMLツール。

### 1.2 目的
講師がゼロから構成を考える負荷を減らし、「直す前提の下書きを一瞬で用意する」。完成品生成ではなく **たたき台生成** が目的。

### 1.3 設計方針（ハード制約）
- 単一の自己完結HTMLファイル。**外部CDN/フォント/ライブラリへの依存ゼロ**。AI呼び出し以外はオフライン動作可。
- **APIキー設定不要**（claude.ai 内蔵のClaudeで生成。`x-api-key` を送らない）。
- ES5互換：**オプショナルチェーン `?.`／Null合体 `??`／後読み正規表現 `(?<…)` を使用しない**。
- すべての `innerHTML` 経路で **エスケープ関数 `esc()` を通す**。
- CSSに `[hidden]{display:none!important}` を含める。
- 出力ファイル名は日付スタンプ付き。本体ファイル名はASCIIのみ。
- バージョン刻印を5か所で一致させる（§14）。
- 公開物に関係者実名を載せない。

---

## 2. 用語定義

| 用語 | 定義 |
|---|---|
| デック（deck） | スライド配列。`slide` オブジェクトの配列 |
| スライド（slide） | 1枚分のデータ（§6.1のスキーマ） |
| OOXML | Office Open XML（ECMA-376）。`.pptx` の中身のXML群 |
| EMU | English Metric Unit。OOXMLの座標単位（914400 EMU = 1 inch） |
| アーティファクト版 | claude.ai 内で動作し内蔵Claudeを使う版（本書の対象） |

---

## 3. 機能要件（FR）

| ID | 機能 | 概要 |
|---|---|---|
| FR-01 | スライド生成 | フォーム入力からプロンプトを組み立て、内蔵Claudeへ送信し、JSONを解析してデック化・描画 |
| FR-02 | 見本表示 | キー/通信なしで内蔵の見本デック（6枚）を描画。書き出しも可能 |
| FR-03 | PPTX書き出し | 純JSでOOXMLを手組みし `.pptx` を生成・ダウンロード |
| FR-04 | PDF書き出し | 印刷用DOMを構築し `window.print()`。ブラウザの「PDFに保存」で出力 |
| FR-05 | Markdown | 全文Markdownのコピー／`.md` 保存 |
| FR-06 | 講師メモ | 「要点／トーク／ワーク」中心のトークスクリプトを `.md` 保存 |
| FR-07 | 使い方「黒板」 | チョーク調の使い方（3ステップ＋ヒント）。既定で開 |
| FR-08 | ログ | 動作・エラーを時刻つきで記録。コピー／クリア。エラー時インジケータ点灯 |
| FR-09 | 作り直す | 現フォーム条件で再生成 |

---

## 4. 非機能要件（NFR）

| ID | 区分 | 要件 |
|---|---|---|
| NFR-01 | 依存性 | 外部CDN/フォント/ライブラリ・localStorage/sessionStorage を使用しない |
| NFR-02 | セキュリティ | 全 `innerHTML` 経路をエスケープ。APIキーは保持しない（artifact版はキー入力自体が無い） |
| NFR-03 | 互換性 | ES5互換（§1.3）。`Array.isArray`/`padStart`不使用（手実装）。`String.prototype.indexOf` で部分一致 |
| NFR-04 | 堅牢性 | 異常データ（型違い・非配列・巨大文字列・特殊文字）でクラッシュしない |
| NFR-05 | アクセシビリティ | `prefers-reduced-motion` でアニメ停止。`label`/`for` 関連付け。フォーカスリング |
| NFR-06 | 可搬性 | 生成 `.pptx` が PowerPoint 互換構造で開ける（独立ツールで検証） |
| NFR-07 | 保守性 | バージョン刻印5か所一致。CHANGELOGをHTML冒頭コメントに保持 |

---

## 5. デザイン仕様

### 5.1 カラートークン（CSS変数・`:root`）
| 変数 | 値 | 用途 |
|---|---|---|
| `--bg` | `#F6F3FB` | 背景（淡ラベンダー白） |
| `--ink` | `#382A52` | 本文 |
| `--ink-soft` | `#6B5A8A` | 補助文字 |
| `--lav-50` | `#EFEAF8` | ノート背景など |
| `--lav-100` | `#E6DEF5` | 番号バッジ・選択チップ |
| `--lav-200` | `#D2C5EC` | 枠線（淡） |
| `--lav-400` | `#A98FD8` | 既定アクセント・スライド上辺 |
| `--lav-500` | `#8E6FCB` | プライマリ（既定の種別色） |
| `--lav-600` | `#7558B0` | ボタン・リンク・講師ノート見出し |
| `--plum` | `#4A3370` | 見出し・タイトル・表紙系種別 |
| `--sage` | `#8AA06E` | ラベンダーの茎（ログ正常ドット） |
| `--sage-deep` | `#6E8456` | 導入/事例の種別色 |
| `--amber` | `#C28A24` | クイズの種別色 |
| `--rose` | `#BE708F` | 演習の種別色・演習ボックス |
| `--line` | `#E0D6F0` | 枠線 |
| `--card` | `#FFFFFF` | カード地 |

> 制約：CSS変数名はASCIIのみ（静的解析で検査）。

### 5.2 フォント（システムスタック＝外部依存なし）
- 見出し `--display`：`"Hiragino Mincho ProN","Yu Mincho","YuMincho",serif`
- 本文 `--body`：`"Hiragino Kaku Gothic ProN","Yu Gothic","YuGothic","Meiryo",sans-serif`
- 影 `--shadow`：`0 10px 30px -12px rgba(74,51,112,.28)`

### 5.3 ラベンダー（花穂）SVGモチーフ
`<symbol id="sprig" viewBox="0 0 40 64">` として定義し、背景・ヘッダー・パネル見出し・ローディングで再利用。
- 茎：`<path d="M20 64 C20 50 19 42 20 28">`、stroke `#8AA06E`、stroke-width `1.6`、linecap round
- 葉：2枚（`#9DB081`）。下記2パス
  - `M20 50 C13 47 10 49 8 53 C13 53 17 53 20 50 Z`
  - `M20 44 C27 41 30 43 32 47 C27 47 23 47 20 44 Z`
- 花穂：楕円8個（下→上）。色は下から `#A98FD8`/`#9476C9`(左右)/`#A98FD8`/`#B49CDF`(左右)/`#8E6FCB`/`#7558B0`。各 `cx,cy,rx,ry` は実装の通り：
  `(20,26,4,5.4)(14.5,22,3.4,4.6)(25.5,22,3.4,4.6)(20,18,3.6,4.8)(15.5,13.5,3,4)(24.5,13.5,3,4)(20,9.5,3,4.2)(20,4,2.4,3.4)`

### 5.4 背景パターン
`<pattern id="lav" width="150" height="170" patternUnits="userSpaceOnUse" patternTransform="rotate(-4)">` 内に `#sprig` を4つ配置（位置・回転・opacityは実装値）。`#bgwrap` は `position:fixed; inset:0; z-index:-1`、矩形に `fill="url(#lav)" opacity=".22"`。

### 5.5 種別カラーマップ（`TYPE_COLORS`）
キー（部分一致）→色：
`表紙/タイトル=#4A3370`, `導入/アイスブレイク=#6E8456`, `本編/解説=#7558B0`, `演習/ハンズオン=#BE708F`, `クイズ=#C28A24`, `事例=#6E8456`, `つまずき=#BE708F`, `まとめ/振り返り/Q&A=#4A3370`、該当なし＝`#8E6FCB`。

---

## 6. データ構造

### 6.1 slide オブジェクト
```
{
  "no":       number,   // 通し番号（無ければ index+1 を使用）
  "type":     string,   // 表紙/導入/本編/演習/クイズ/事例/まとめ 等（文字列でなくても可・強制変換）
  "title":    string,
  "bullets":  string[], // 0〜5項目（非配列は空配列として扱う）
  "notes":    string,   // 講師ノート（口頭の要点）
  "minutes":  number,   // 所要分（非数値は0扱い）
  "activity": string    // 演習/クイズの具体指示。無ければ空文字
}
```
deck は上記の配列。`parseDeck` は **配列であること・1件以上** を不変条件として検証。

---

## 7. 画面設計（DOM構成）

```
header（眉ラベル/タイトル/リード/badges[verpill,badge-soft]/sprig装飾×2）
.wrap
 ├ details.board#board[open]      使い方「黒板」: 3ステップ + tips
 ├ .grid
 │   ├ .panel（研修の設計図）
 │   │   └ form#designer
 │   │       theme(text,必須) / audience(select) / level(select)
 │   │       duration(select,value="表示,枚数目安") / tone(select)
 │   │       goal(textarea,任意) / elements(checkbox×6) / ramen(checkbox)
 │   │       button#go「スライドを生成する」
 │   └ .stage#stage → .empty#empty（初期）/ ローディング / .toolbar+.slide… / .err
 └ details.logpanel#logPanel      ログ: tools(copy/clear) + pre#logBox + dot#logDot
footer（版・日付刻印）
#printArea（印刷専用・通常非表示）
.toast#toast
```

### 7.1 フォーム項目（既定値）
- audience：新卒エンジニア / 中途・若手エンジニア / 文系出身の未経験者 / 新卒・中途の混在
- level：入門（前提知識なし）/ 初級 / 中級 / 応用
- duration（`value="表示,枚数目安"`）：`30分,6〜8` / `60分,9〜12`(既定) / `90分,12〜14` / `半日(3〜4時間),12〜15(章立て中心)`
- tone：やさしく丁寧に / テンポよく明るく / がっつり技術重視
- elements（checkbox, 既定☑印）：導入アイスブレイク☑ / ハンズオン演習☑ / 理解度クイズ / 実務事例の紹介 / つまずきポイント解説 / 振り返りとQ&A☑
- ramen（checkbox, 既定OFF）：うさうさラーメン店のたとえ話を使う

### 7.2 スライドカード描画（`.slide`）
上辺色＝種別色。ヘッダに番号バッジ・タイトル・種別タグ・所要分。本文に箇条書き（`.bullets`、各項目に丸ビュレット）、`.activity`（あれば、ローズ左帯）、`.note`（あれば、ラベンダー左帯）。

---

## 8. 関数仕様（詳細設計）

> すべて IIFE/グローバル関数。`var`/関数宣言ベース。テスト用に末尾で `module.exports`（ブラウザでは未定義のため無害）。

### 8.1 共通ユーティリティ
| 関数 | 入出力 | 仕様 |
|---|---|---|
| `esc(s)` | string→string | `null/undefined`→`''`、`& < >` をHTMLエンティティ化。**全innerHTML経路で使用** |
| `tagColor(t)` | any→hex(`#......`) | `t` を `String()` 強制後、`TYPE_COLORS` を部分一致検索。無ければ `#8E6FCB` |
| `hex(t)` | any→hex(`......`) | `tagColor(t)` から `#` を除去 |
| `arr(x)` | any→array | 配列ならそのまま、でなければ `[]`（`Object.prototype.toString` 判定） |
| `toast(msg)` | string→void | 下部トースト表示、約1.9秒で消える |
| `todayStamp()` | →`YYYYMMDD` | ゼロ埋め手実装（`padStart`不使用） |

### 8.2 ログ
- `LOG`：`{t,level,msg}[]`。`level` は `info/ok/warn/err`。
- `log(level,msg)`：時刻整形→`LOG`追加→`#logBox` に色分け追記（`lerr/lwarn/lok`）→`err` で `#logDot` に `.err` 付与。msgは `esc` 経由。
- `window.addEventListener('error', …)` / `('unhandledrejection', …)` で捕捉してログ化。
- `copyLog()`：行整形して `copyText`。`clearLog()`：`LOG` を空にし表示・ドットをリセット。

### 8.3 生成系
- `buildPrompt()`：フォーム値を読み、§11のプロンプト文字列を返す（文字列連結。`ramen` ON時のみラーメン店指示を挿入）。
- `parseDeck(text)`：```` ```json ````/```` ``` ```` 除去→最初の `[` 〜最後の `]` を切り出し→`JSON.parse`→配列かつ非空を検証（違反は `throw`）。
- `generate()`：§9のフロー。Promiseを返す（テスト容易性のため）。
- `renderLoading()` / `renderError(msg)`：ステージ差し替え。`renderError` は `esc` 経由。
- `totalMins(d)`：`minutes` を `Number()`、非数値は0で総和。
- `renderDeck(deck)`：ツールバー＋カード群を生成。動的値はすべて `esc`。`bullets` は `arr()` でガード。

### 8.4 書き出し系
| 関数 | 仕様 |
|---|---|
| `deckToMarkdown()` | `# テーマ` → 各 `## n. title [type]（m分）` + 箇条書き + `> 🛠`/`> 🎙`。末尾に署名 |
| `deckToMemo()` | `# テーマ 講師メモ` + 合計分/枚数 → 各 `### [n] title（type・m分）` + 要点/トーク/ワーク |
| `buildPrintHtml()` | 表紙 `.p-cover` + 各 `.p-slide`（上辺=種別色、番号/タイトル/タグ/分、ul、`.p-act`/`.p-note`）。全て `esc` |
| `saveBlob(content,filename,mime)` | `Blob`→`URL.createObjectURL`→`<a download>` クリック→`revokeObjectURL` |
| `safeName()` | `currentTheme||'研修スライド'` から `\ / : * ? " < > |` を `_` 置換 |
| `copyText(text,okMsg)` | `navigator.clipboard` 優先、失敗時 `textarea + execCommand('copy')` フォールバック |
| `copyMarkdown/downloadMarkdown/downloadMemo` | それぞれ §FR-05/06。ファイル名は `safeName()+…+'_'+todayStamp()+'.md'` |
| `exportPdf()` | `#printArea` に `buildPrintHtml()` を流し込み、120ms後 `window.print()` |
| `exportPptx()` | `buildPptx(currentDeck)`→`saveBlob(…'_'+todayStamp()+'.pptx', PPTX MIME)`。try/catchでログ |

### 8.5 ZIP / PPTX 低レベル（§12, §13参照）
`crc32`, `makeZip`, `pEsc`, `pRun`, `pPara`, `pSlideXml`, `buildPptx`。

### 8.6 見本
`showSample()`：固定6枚（表紙/導入/本編/演習/クイズ/まとめ）を `currentDeck` に設定し `renderDeck`。

---

## 9. 生成処理フロー（FR-01 詳細）

```
1. theme = trim(#theme)。空ならトースト＋警告ログ＋return（fetchしない）
2. #go を disabled、文言「生成中…」、renderLoading()、currentTheme=theme、開始ログ
3. fetch POST https://api.anthropic.com/v1/messages
   headers: { "Content-Type":"application/json" }   ※x-api-key等は付けない
   body: { model:"claude-sonnet-4-6", max_tokens:4096, messages:[{role:"user",content:buildPrompt()}] }
4. res.ok でなければ：res.json() の error.message を付してErrorをthrow（取得失敗時はステータスのみ）
5. data.content から type==="text" を連結 → parseDeck() → currentDeck → renderDeck()
6. 成功/失敗をログ（所要msも記録）
7. finally相当：#go を復帰
8. catch：Failed to fetch/NetworkError は親切な日本語に変換し renderError + err ログ
```

---

## 10. AI連携仕様

| 項目 | 値 |
|---|---|
| エンドポイント | `POST https://api.anthropic.com/v1/messages` |
| ヘッダ | `Content-Type: application/json` のみ（**キー・versionヘッダは送らない**＝内蔵ブリッジが処理） |
| model | `claude-sonnet-4-6` 固定 |
| max_tokens | `4096`（デックJSONの切り詰め防止のため。※artifact既定の1000では不足し得る） |
| messages | `[{role:"user",content:<プロンプト>}]` |
| 応答解析 | `content[].type==="text"` を連結 → `parseDeck` |

---

## 11. プロンプト・テンプレート（完全版）

`buildPrompt()` が返す文字列（`${}` は対応するフォーム値、`ramen` ON時のみラーメン行を挿入）：

```
あなたは新卒エンジニア研修を数多く手がける一流の研修設計者です。以下の条件で、研修ワークショップ用のスライド構成案を作ってください。

# 条件
- 研修テーマ: ${theme}
- 対象者: ${audience}
- レベル: ${level}
- 所要時間: ${duration}（スライド枚数の目安: ${countHint}枚）
- トーン: ${tone}
- 学習ゴール: ${goal || 'テーマに沿って講師として適切に設定する'}
- 含める要素: ${els.join('、') || 'おまかせ'}
- 「うさうさラーメン店」というラーメン店のたとえ話を1〜2か所で使い、専門用語を初心者にやさしく解説すること。   ← ramen=true時のみ

# 出力ルール
- 必ず日本語で書く。
- 1枚目は表紙、最後は「まとめ／振り返り」で終える。
- bullets は各スライド最大5項目、短いフレーズで。
- notes には講師が口頭で話す要点を1〜2文で書く。
- 演習・クイズのスライドは activity に具体的な手順や問いを書く（それ以外は空文字）。
- type は「表紙／導入／本編／演習／クイズ／事例／まとめ」から選ぶ。
- minutes は所要時間の目安(分)を数値で。
- JSON配列のみを出力。前置き・後置き・マークダウンの囲みは一切書かない。

# 出力フォーマット
[{"no":1,"type":"表紙","title":"...","bullets":["..."],"notes":"...","minutes":3,"activity":""}]
```
`duration` と `countHint` は `select` の `value="表示,枚数目安"` を `,` 分割して得る。

---

## 12. ZIP書き出し仕様（純JS・無圧縮ストア）

- 圧縮方式：**ストア（method=0、無圧縮）**。
- フラグ：`0x0800`（ファイル名UTF-8）。version-needed=`20`。
- 文字列は `TextEncoder` でUTF-8バイト化。ファイル名はASCIIのみ。
- **CRC32**：多項式 `0xEDB88320` のテーブル方式。`crc32(bytes)` を各エントリで算出。
- レコード署名：ローカルヘッダ `0x04034b50` / 中央ディレクトリ `0x02014b50` / 終端 `0x06054b50`。
- 各エントリ：ローカルヘッダ→ファイル名→データ。最後に中央ディレクトリ群→終端レコード。
- `makeZip(files)`：`files=[{name, data:string|Uint8Array}]` を受け、`Uint8Array` を返す。

> 検証：本方式の出力は `python-pptx`/標準ZIPで展開でき、CRC健全（`testzip()===None`）を確認済み。

---

## 13. PPTX（OOXML）生成仕様

### 13.1 スライドサイズ（16:9）
`sldSz cx=12192000 cy=6858000`（EMU）。`notesSz cx=6858000 cy=12192000`。

### 13.2 同梱パート一覧（最小構成）
```
[Content_Types].xml                         ← Default(rels,xml) + Override(presentation/master/layout/theme/各slide)
_rels/.rels                                 ← officeDocument → ppt/presentation.xml
ppt/presentation.xml                        ← sldMasterIdLst, sldIdLst(256+i), sldSz, notesSz
ppt/_rels/presentation.xml.rels             ← master(rId1), theme(rIdTheme), 各slide(rId2..)
ppt/theme/theme1.xml                        ← clrScheme=Lavender(accent1..6に§5.5系の色)
ppt/slideMasters/slideMaster1.xml           ← 背景 #F6F3FB、clrMap、txStyles、sldLayoutId 2147483649
ppt/slideMasters/_rels/slideMaster1.xml.rels← layout(rId1), theme(rId2)
ppt/slideLayouts/slideLayout1.xml           ← type="blank"、masterClrMapping
ppt/slideLayouts/_rels/slideLayout1.xml.rels← master(rId1)
ppt/slides/slideN.xml                        ← §13.3
ppt/slides/_rels/slideN.xml.rels             ← layout(rId1)
```
- `sldMasterId id="2147483648"`、`sldLayoutId id="2147483649"`、`sldId id` は `256+index`。
- スライドの `r:id` は `rId(index+2)`（rId1はmaster）。

### 13.3 スライドXML（`pSlideXml`）図形構成
3つの `p:sp`（テキストは `ja-JP`）：
1. **Title**（id=2）：off `685800,457200` ext `10820400×1143000`、`anchor="b"`、文字サイズ `3600`・色 `4A3370`・太字。
2. **Accent**（id=3）：off `685800,1645920` ext `2286000×45720`、塗り＝種別色（`hex(type)`）。装飾の細帯。
3. **Body**（id=4）：off `685800,1965960` ext `10820400×4434840`、`normAutofit`。段落構成：
   - 箇条書き：`marL=285750 indent=-285750`、`buChar char="•"(&#8226;)`、サイズ `2000`・色 `382A52`。
   - 演習あり：見出し「🛠 演習・ワーク」(1400, `BE708F`, 太字) + 本文(1500, `5D3A4C`)。
   - ノートあり：見出し「🎙 講師ノート」(1400, `7558B0`, 太字) + 本文(1500, `6B5A8A`)。
- 各スライドに `clrMapOvr/overrideClrMapping`（bg1=lt1 … の標準マップ）を付与。
- XML特殊文字は `pEsc`（`& < >`）でエスケープ。

### 13.4 検証要件（必須）
生成 `.pptx` を独立ツール（`python-pptx`）で開き、(1)スライド枚数一致、(2)日本語テキスト読み戻し、(3)全XML整形式、(4)ZIP CRC健全、を満たすこと。

---

## 14. バージョン刻印仕様（5か所一致）

| # | 箇所 | 内容 |
|---|---|---|
| 1 | ファイル名 | `..._v1_0_0_pro_20260613.html` |
| 2 | HTML冒頭コメント | `Version: 1.0.0 PRO (2026-06-13)` |
| 3 | CHANGELOG（同コメント内） | `v1.0.0 PRO (2026-06-13)` 項 |
| 4 | verpillバッジ（ヘッダ） | `v1.0.0 PRO` |
| 5 | フッター | `PRO v1.0.0 ・ 2026-06-13` |

> 静的解析で5か所の整合を検査する（§15）。

---

## 15. 品質保証（テスト設計）

### 15.1 方針
商用グレードの jsdom テスト campaign を **版上げ前に必須**。**全項目グリーンを5回連続**で確認してから刻印。見つけた不具合は「製品バグ／テストコード由来」を正直に切り分け、製品バグは刻印前にゼロにする。

### 15.2 区分と主な観点（合計65項目・現行）
- **静的解析**：禁止構文（`?.`/`??`/lookbehind）不在、`[hidden]` ルール、CSS変数ASCII、div開閉一致、`esc()` 使用、`node --check` 構文、刻印5か所一致、キー入力欄/`x-api-key` 不在、内蔵モデル名。
- **単体**：`esc`/`tagColor`/`hex`/`totalMins`/`crc32`（既知値 `crc32("123456789")=0xCBF43926`）/`todayStamp`。
- **ホワイトボックス**：`buildPrompt`（ラーメンON/OFF・テーマ反映）、`parseDeck`（フェンス除去・非配列/空配列で例外）、`pSlideXml`（activity/notes分岐・buChar）。
- **結合**：見本6枚描画＋ツールバー、`safeName`、**XSSエスケープ**（タイトル/箇条書き）、`buildPptx`（PK署名・必須パート・slide数一致）。
- **総合（generate非同期・fetchモック）**：成功／HTTPエラー（error.message表示）／ネット断（親切文言）／テーマ空でfetch非呼出／壊れJSON。
- **ブラックボックス**：Markdown/講師メモ/印刷HTMLの内容、クリップボード連携。
- **モンキー**：null・型違い・非配列bullets・巨大文字列・特殊文字・改行/タブで無クラッシュ。

### 15.3 既知の修正事例（v1.0.0、刻印前）
モンキーテストが検出：`type` が非文字列だと `tagColor` が、`bullets` が非配列だと `.map` が落ちる → `tagColor` の文字列強制と `arr()` ガードで堅牢化。**製品側の実バグ**として刻印前に解消（版据え置き）。

---

## 16. 印刷（PDF）仕様

- `@media print`：`@page{size:A4 landscape;margin:14mm}`。`#bgwrap,header,.board,.grid,.logpanel,footer,.toast` を非表示、`#printArea` を表示。
- `.p-slide` は `page-break-after:always`（最終は auto）、上辺＝種別色。表紙 `.p-cover` は1枚目。
- 出力はブラウザの「PDFに保存」を利用（**外部依存なしを維持するため**ライブラリPDF化はしない）。

---

## 17. 既知の制限・非対応

- PDFはワンクリック直出力ではなく印刷ダイアログ経由（外部依存なし維持のため）。
- `.pptx` の講師ノートはスライド本文内に配置（PowerPointの「ノート」ペインへは分離しない）。
- 半日など枚数が多い設定では、`max_tokens` 上限により内容が簡潔化される場合がある。
- 生成にはネット接続が必要（見本・各書き出しはオフライン可）。

---

## 18. 再現・QAワークフロー（参考手順）

```
# 1. 本書 §5〜§13 に従い単一HTMLを実装（ファイル名は §14-1）
# 2. PPTXコードの妥当性検証（コンテナ等）
pip install python-pptx
node でbuildPptxを実行し .pptx を出力 → python-pptx で開封/整形式/CRC を確認
# 3. テスト campaign（jsdom）
npm install jsdom
node test_campaign.mjs          # 全65項目グリーンを確認
for i in 1..5; do node test_campaign.mjs; done   # 5回連続グリーン
# 4. 製品バグがあれば修正 → 再実行（ゼロになるまで）
# 5. §14 の5か所を一致させて刻印 → 納品（present_files）
```

---

## 19. 変更履歴（CHANGELOG）

- **v1.0.0 PRO (2026-06-13)**
  - 初の商用テスト済みリリース（アーティファクト版：APIキー設定不要・内蔵Claudeで生成）。
  - スライド生成（表紙〜まとめ／講師ノート／演習／所要分）。
  - 書き出し：`.pptx`（純JS・外部依存なしOOXML）／PDF（印刷）／Markdown／講師メモ。
  - 使い方「黒板」、🐛ログ、見本デック。
  - 開発版（v0.x：チャット内artifact／キー入力スタンドアロン）からの正式版。
  - jsdomテスト全65項目グリーン（5回連続）。モンキー検出の実バグを刻印前に修正。

---

### 付記（事実と所感・出典）
- 本書の仕様値（色・座標・サイズ・ヘッダ・件数等）は **v1.0.0 PRO 実装の実値** です。
- 「`.pptx`＝OOXML(ECMA-376)のZIP」は公開標準にもとづく一般的事実。検証手順・件数は今回ビルドで確認した実績のみ記載（推測値・効果指標は不記載）。
- 【所感】最小構成のOOXMLを手組みできると、`.pptx` が「テキストの束」に見えてきます。この透明性こそ、未経験の方に渡したい学びだと考えています。
- 本書は社内レビュー（査読）前の初稿。公開前にメイン講師・人事担当の確認を想定。関係者の個人名は不記載（必要時は役職名）。

> 面白きこともなき世を面白く 🐰🪻
