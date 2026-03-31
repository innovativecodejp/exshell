# Exshell 外部仕様書 v0.1

## 1. 文書情報

- 文書名: Exshell 外部仕様書
- 版数: v0.1
- 目的: Excel を入出力ウインドウとして利用する CLI ツール Exshell の外部仕様を定義する
- 想定読者: 設計者、実装者、利用者

---

## 2. 概要

Exshell は、Excel 上のテキストボックスを入出力対象として扱うコマンドラインツールである。  
PowerShell ターミナル上から操作し、Excel のテキストボックス内容を標準出力へ取り出したり、標準入力から書き戻したり、複数のテキストボックス内容を比較したりできる。

Exshell は単独でテキスト処理機能を抱え込まず、WSL 上の UNIX 系コマンドと連携して軽快な操作性を実現することを基本思想とする。

---

## 3. 基本方針

### 3.1 目的

- Excel をテキスト入出力のフロントエンドとして利用する
- UNIX CLI の軽快さ、透明感を Excel 利用環境へ持ち込む
- Excel、PowerShell、WSL を橋渡しする
- Excel テキストボックスをテキスト処理パイプラインへ接続する

### 3.2 実行環境前提

- OS は Windows とする
- Excel デスクトップ版がインストール済みであること
- WSL が使用可能であること
- PowerShell ターミナル上で利用すること
- Exshell は Excel COM Interop を用いて Excel を操作する

### 3.3 非対象

- Excel Web 版
- Excel 非インストール環境
- ActiveX TextBox の完全対応
- フォームコントロールの完全対応
- Excel ファイル保存管理
- Excel ブックのクローズ管理
- Excel アプリケーション終了管理

---

## 4. Exshell の責務

Exshell の責務は以下とする。

- 指定 Excel ブックを開く、または既に開いている対象ブックを操作対象にする
- 対象 Excel ブック内のテキストボックス一覧を取得する
- 指定テキストボックス内容を標準出力へ出力する
- 標準入力を指定テキストボックスへ書き込む
- 指定 2 テキストボックス内容を比較するため、一時ファイル化し WSL diff を実行する
- 必要に応じて Excel テキストを UNIX 系コマンドへ受け渡す橋渡しを行う

---

## 5. ライフサイクル管理方針

### 5.1 保存方針

Exshell は対象 Workbook に対して Save を実行しない。

- `cate` による書き込み後も Save は行わない
- 変更内容の保存可否は利用者が Excel 側で判断する
- 名前を付けて保存、通常保存は Excel 側の操作で行う

### 5.2 クローズ方針

Exshell は対象 Workbook に対して Close を実行しない。

- ブックを閉じる操作は Excel 側で行う
- Exshell は Excel.Application の Quit を実行しない
- 保存、破棄、閉鎖の最終判断は Excel 操作側の責任とする

---

## 6. 対象オブジェクト

### 6.1 対象

初期仕様で対象とするのは、Worksheet 上の Shape テキストボックスとする。

### 6.2 対応範囲

- `Shape.TextFrame2` または `Shape.TextFrame` から文字列取得可能な Shape
- テキストを持つ図形

### 6.3 非対応範囲

- ActiveX TextBox
- フォームコントロール
- 画像
- 文字列を持たない Shape
- グループ化図形の完全対応

---

## 7. 識別方式

### 7.1 識別子形式

テキストボックスの正式識別子は以下の形式とする。

`SheetName:ShapeName`

例:

```text
Main:txtInput
Main:txtOutput
Diff:left
Diff:right
```

### 7.2 シート省略時

シート名省略時は、既定シートを使用する。

### 7.3 既定シート

既定シートは `eopen` 実行時に指定可能とする。  
指定しない場合は、対象 Workbook のアクティブシートを既定シートとする。

---

## 8. セッション仕様

### 8.1 セッションの役割

`eopen` 後に `ecat`、`cate`、`ediff` を継続実行するため、Exshell は現在対象 Workbook および既定シートをセッションとして保持する。

### 8.2 セッション保持項目

- 対象 Workbook のフルパス
- 既定シート名

### 8.3 セッション保存場所

初期仕様ではローカルセッションファイルに保存する想定とする。  
保存先詳細は内部設計で定義する。

---

## 9. コマンド仕様

### 9.1 `eopen`

#### 形式

```powershell
eopen <excel file> [--sheet <sheet name>]
```

#### 機能

- 指定 Excel ファイルを開く
- 既に開いている場合はその Workbook を対象とする
- 対象 Workbook を現在セッションに設定する
- 必要に応じて既定シートを設定する

#### 備考

- Save は行わない
- Close は行わない
- Quit は行わない

#### 使用例

```powershell
eopen D:\work\sample.xlsx

eopen D:\work\sample.xlsx --sheet Main
```

---

### 9.2 `els`

#### 形式

```powershell
els [--sheet <sheet name>]
```

#### 機能

- 対象シート上のテキストボックス一覧を表示する
- `--sheet` 指定時は当該シートを対象とする
- 指定省略時は既定シートを対象とする

#### 出力例

```text
[Main]
txtInput
txtOutput
left
right
```

---

### 9.3 `ecat`

#### 形式

```powershell
ecat <textbox>
```

#### 機能

- 指定テキストボックス内容を標準出力へ出力する
- シート省略時は既定シートを使用する

#### 使用例

```powershell
ecat Main:txtOutput
ecat txtOutput
```

#### 想定利用例

```powershell
ecat Main:src | unl

ecat Main:src | ugrep "ERROR"
```

---

### 9.4 `cate`

#### 形式

```powershell
cate <textbox> [--append]
```

#### 機能

- 標準入力を指定テキストボックスへ書き込む
- デフォルトは上書き
- `--append` 指定時は追記する
- シート省略時は既定シートを使用する

#### 備考

- Save は行わない
- Close は行わない

#### 使用例

```powershell
type memo.txt | cate Main:txtInput

ecat Main:src | usort | cate Main:sorted

ecat Main:log | ugrep "ERROR" | cate Main:errorOnly --append
```

---

### 9.5 `ediff`

#### 形式

```powershell
ediff <textbox1> <textbox2>
```

#### 機能

- 指定 2 テキストボックスの内容を取得する
- 一時ファイル `temp1`、`temp2` を生成する
- WSL 上で `diff temp1 temp2` を実行する
- 差分結果を標準出力へ出力する
- 処理完了後、一時ファイルは削除する

#### 一時ファイル仕様

- 文字コード: UTF-8
- 改行コード: LF

#### 実行イメージ

```powershell
ediff Main:left Main:right
```

内部的には概念上、以下に相当する。

```powershell
ecat Main:left  > temp1
ecat Main:right > temp2
udiff temp1 temp2
```

#### 備考

- `diff` 実行は WSL を使用する
- 差分アルゴリズムは Exshell 自前実装ではなく、WSL `diff` を利用する

---

### 9.6 `einfo`

#### 形式

```powershell
einfo
```

#### 機能

- 現在セッションの対象 Workbook、既定シート等を表示する

#### 出力例

```text
Workbook : D:\work\sample.xlsx
Sheet    : Main
Excel    : Running
```

---

## 10. WSL 補助コマンド運用

Exshell 周辺の補助コマンドとして、PowerShell 上で以下の関数を定義する運用を前提とする。

```powershell
function uls   { wsl ls @Args }
function ucat  { wsl cat @Args }
function unl   { wsl nl @Args }
function udiff { wsl diff @Args }

function ugrep { wsl grep @Args }
function used  { wsl sed @Args }
function uawk  { wsl awk @Args }
function usort { wsl sort @Args }
function uuniq { wsl uniq @Args }
```

### 10.1 方針

- `u` は UNIX 系操作を意味する接頭辞とする
- PowerShell 標準の `ls` 等は上書きしない
- 利用者には UNIX 的操作感を提供しつつ、実装は WSL に委譲する

---

## 11. 利用例

### 11.1 行番号付き表示

```powershell
ecat Main:src | unl
```

### 11.2 grep 抽出

```powershell
ecat Main:log | ugrep "ERROR"
```

### 11.3 sort 結果を別テキストボックスへ書き戻し

```powershell
ecat Main:list | usort | cate Main:sorted
```

### 11.4 diff 表示

```powershell
ediff Main:left Main:right
```

### 11.5 追記

```powershell
ecat Main:todayLog | cate Main:history --append
```

---

## 12. エラー方針

### 12.1 基本方針

- 正常終了時は終了コード 0
- 異常終了時は非 0
- エラーメッセージは標準エラー出力へ出力する

### 12.2 想定エラー

- 引数不正
- セッション未確立
- Excel ファイル未検出
- 対象 Workbook 未検出
- 対象シート未検出
- 対象 Shape 未検出
- 文字列取得不能 Shape 指定
- WSL `diff` 実行失敗
- 一時ファイル作成失敗

### 12.3 代表的な終了コード例

- 1: 引数エラー
- 2: セッション未確立
- 3: Workbook 未検出
- 4: Worksheet 未検出
- 5: Shape 未検出
- 6: Excel 操作失敗
- 7: WSL 実行失敗

---

## 13. 制約事項

- Excel COM 操作に依存するため、Excel 非導入環境では動作しない
- Shape 種別差異により、すべての図形で安定動作を保証するものではない
- WSL 未導入環境では `ediff` および UNIX 補助コマンド運用を前提とした利用はできない
- 初期仕様では Shape 名の一意性確保を利用者運用に委ねる

---

## 14. 今後の拡張候補

- `eimport` : ファイル内容をテキストボックスへ反映
- `eexport` : テキストボックス内容をファイルへ出力
- `eclear` : テキストボックス内容をクリア
- `epipe` : テキストボックス → 外部コマンド → テキストボックスの一括処理
- `els --verbose` : Shape 種別、シート名、文字数表示
- `ediff` への diff オプション透過指定

---

## 15. 外部仕様まとめ

Exshell は Excel を UI とするテキスト処理ブリッジであり、以下を特徴とする。

- Excel テキストボックスを CLI 入出力対象とする
- PowerShell 上で操作する
- WSL の UNIX 系コマンドと組み合わせて利用する
- 保存、クローズ、終了は Excel 側操作に委ねる
- Exshell 自身は橋渡しに徹し、重いテキスト処理機能は WSL に委譲する

この方針により、Excel 上のテキストと UNIX 的テキスト処理環境を自然に接続することを目指す。
