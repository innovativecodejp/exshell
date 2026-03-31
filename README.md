# Exshell

Excel を、Unix CLI 風に扱うための軽量 Shell です。  
PowerShell から Excel 上のテキストボックスを読み書きし、WSL の `diff` や `grep` などと組み合わせて使えます。

Exshell は、Excel を単なる表計算ツールではなく、**研究・検討・比較・メモ整理のためのテキスト入出力フロントエンド**として再利用するために作りました。

---

## 何ができるか

- Excel のテキストボックス内容を CLI から読む
- パイプ入力を Excel のテキストボックスへ書く
- 2つのテキストボックス内容を `diff` で比較する
- PowerShell + WSL + Excel をつないで、Unix 的なテキスト処理を Excel 上で行う

たとえば、次のような使い方を想定しています。

```powershell
eopen D:\work\sample.xlsx --sheet Main

ecat Main:src | unl

ecat Main:log | ugrep "ERROR"

ecat Main:left | usort | cate Main:sorted

ediff Main:left Main:right
```

---

## 想定ユースケース

- 研究メモや検討ログの比較
- 仕様草案やレビューコメントの差分確認
- Excel を使ったテキスト整理作業の効率化
- PowerShell / WSL を使った軽量なテキスト処理パイプライン
- 「Excel 上に置いたテキストを、CLI で扱いたい」場面全般

特に、**diff を頻繁に見る作業**、**テキストを一時的に並べて比較したい作業**、**Excel を見ながら思考を進める研究・設計作業**に向いています。

---

## 主なコマンド

| コマンド | 役割 |
|---|---|
| `eopen` | 対象の Excel ブックを開き、現在セッションに設定する |
| `els` | テキストボックス一覧を表示する |
| `ecat` | テキストボックス内容を標準出力へ出す |
| `cate` | 標準入力をテキストボックスへ書き込む |
| `ediff` | 2つのテキストボックス内容を比較する |
| `einfo` | 現在のセッション情報を表示する |

識別子は `SheetName:ShapeName` 形式です。

例:

```text
Main:txtInput
Main:txtOutput
Diff:left
Diff:right
```

---

## Exshell の思想

Exshell は、重い機能を抱え込むツールではありません。

方針は明確です。

- Excel 固有の操作だけを Exshell が担う
- テキスト加工や比較は WSL の既存コマンドへ委譲する
- 保存やクローズは Excel 側で判断する
- Exshell 自体は「橋渡し」に徹する

この割り切りにより、実装を肥大化させず、**Excel と Unix CLI の長所をそのまま接続する**ことを狙っています。

---

## 動作前提

- Windows
- Excel デスクトップ版
- PowerShell
- WSL
- .NET Framework 4.8

初期版では、Worksheet 上の Shape テキストボックスを対象としています。

---

## 制約

- Excel Web 版は対象外です
- ActiveX TextBox やフォームコントロールは初期版では対象外です
- Save / Close / Excel 終了は Exshell では行いません
- `ediff` は WSL の `diff` を利用します

---

## なぜ GitHub に公開するのか

Exshell は、一般的な Excel 自動化ツールというより、
**Excel を研究・設計・比較作業のためのテキスト UI として再定義する試み**です。

Excel を日常的に使いながら、PowerShell や Unix CLI の操作感も捨てたくない。  
その中間にある実務上の不便を、小さく埋めるためのツールとして公開します。

---

## 今後の拡張候補

- `eimport` / `eexport`
- `eclear`
- `epipe`
- `els --verbose`
- `ediff` へのオプション透過

---

## ステータス

初期実装・検証フェーズです。  
まずは、`eopen` / `ecat` / `cate` / `els` / `ediff` を中核機能として整備していきます。

---

## ライセンス

未定
