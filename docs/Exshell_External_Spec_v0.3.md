# Exshell 外部仕様書 v0.3

## 1. 文書情報
- 文書名: Exshell 外部仕様書
- 版数: v0.3
- 対象: Exshell コマンドラインツール群
- 目的: Excel を入出力インフラとして扱い、PowerShell / WSL の双方から同等に利用できる CLI 仕様を定義する

---

## 2. Exshell の概要

Exshell は、Excel のテキストボックスを CLI の標準入出力に接続するためのブリッジツールである。  
Excel を GUI 側の入出力ウインドウとして用い、PowerShell および WSL から同等に操作できることを目的とする。

Exshell は以下の思想で設計する。

- Excel を人間向けの入出力 UI として使う
- CLI の軽快さ、透明感、合成可能性を維持する
- Unix コマンドは WSL 側の既存コマンドを活用する
- PowerShell 側でも単独利用しやすくする
- Exshell 自体は Excel と標準入出力を橋渡しすることに責務を絞る

---

## 3. 対象環境

### 3.1 OS
- Windows 11 を主対象とする

### 3.2 必須ソフトウェア
- Microsoft Excel デスクトップ版がインストール済みであること
- WSL が利用可能であること
- PowerShell が利用可能であること

### 3.3 開発・実装前提
- Exshell 本体は Windows 上の実行ファイルとして実装する
- Excel 操作は COM Interop により行う
- WSL からは Windows 実行ファイルを呼び出して使用する

---

## 4. 運用方針

### 4.1 両対応方針
Exshell は以下の 2 つの運用形態を正式にサポートする。

- PowerShell から直接 Exshell コマンドを実行する運用
- WSL から Exshell コマンドをラップして実行する運用

どちらを主運用とするかは利用者の好みと利用場面による。  
仕様上は両者を同等に扱う。

### 4.2 PowerShell 側の位置づけ
PowerShell では、Exshell をそのまま直接実行する。  
必要に応じて `uls`, `ucat`, `unl`, `udiff` などの補助関数を用いて WSL コマンドを簡単に呼び出してよい。

### 4.3 WSL 側の位置づけ
WSL では、`ls`, `grep`, `cat`, `nl`, `diff`, `sed`, `awk` などの Unix コマンドをそのまま用いる。  
Exshell コマンドは WSL の `alias` または `function` を介して Windows 実行ファイルを呼び出す。

### 4.4 保存・クローズ方針
Exshell は操作対象 Excel ブックに対して以下を実行しない。

- Save
- Close
- Excel.Application.Quit

保存、破棄、閉じる、終了は Excel 側で利用者が実行する。

---

## 5. Exshell の責務

Exshell の責務は次の通りとする。

- 指定 Excel ブックを開く、または既存の開いているブックを操作対象に設定する
- Excel シート上のテキストボックス一覧を取得する
- Excel テキストボックスの内容を標準出力へ送る
- 標準入力の内容を Excel テキストボックスへ書き込む
- Excel テキストボックス内容の比較を支援する
- PowerShell / WSL の双方から同等に利用できる CLI を提供する

次は Exshell の責務外とする。

- テキスト整形ロジックの大量実装
- grep, sort, sed, awk, diff などの既存 Unix コマンドの再実装
- Excel ブックの保存、閉鎖、終了管理
- 高度な Excel 編集機能全般

---

## 6. 操作対象

### 6.1 Excel ブック
- `.xlsx`
- `.xlsm`
- 必要に応じて `.xls` は将来拡張対象とする

### 6.2 Excel オブジェクト
初期版では、シート上の Shape 系テキストボックスを対象とする。

対象:
- Shape として配置されたテキストボックス
- TextFrame または TextFrame2 から文字列取得・設定可能な Shape

非対象:
- ActiveX TextBox
- フォームコントロール
- 画像、図、グループ化オブジェクトなど文字列取得不能な Shape

### 6.3 識別子
テキストボックス識別子は原則として次の形式とする。

- `SheetName:ShapeName`

既定シートが設定済みの場合は、シート名の省略を許可する。

例:
- `Main:txtInput`
- `Main:txtOutput`
- `txtInput`

---

## 7. セッション仕様

### 7.1 セッションの目的
`eopen` 実行後に、以後の `ecat`, `cate`, `els`, `ediff` が既定対象ブックを参照できるようにする。

### 7.2 セッション保持内容
最低限、以下を保持する。

- 対象ブックのフルパス
- 既定シート名
- 必要に応じてブック識別情報

### 7.3 セッションの永続化
セッション情報はローカルに保存してよい。  
保存場所は実装依存とするが、初期実装ではユーザープロファイル配下を推奨する。

---

## 8. パス仕様

### 8.1 PowerShell 側
PowerShell からは Windows パスをそのまま指定できる。

例:
```powershell
eopen D:\work\sample.xlsx
```

### 8.2 WSL 側
WSL からは Linux 形式パスを指定できる。

例:
```bash
eopen /mnt/d/work/sample.xlsx
```

### 8.3 パス解決方針
Exshell 本体は以下の双方を受け入れることを推奨する。

- Windows 形式パス
- WSL 形式パス

WSL 形式パスが渡された場合は、内部で Windows 形式パスへ変換する。

---

## 9. 文字コード・改行

### 9.1 標準入出力
- 標準入力・標準出力は UTF-8 を基本とする

### 9.2 一時ファイル
- 差分処理用一時ファイルは UTF-8 で作成する
- 改行は LF を推奨する

### 9.3 Excel 反映時
Excel テキストボックス反映時の改行は表示互換性を考慮し、内部で必要な変換を許容する。

---

## 10. コマンド一覧

初期版で定義する主要コマンドは次の通りとする。

- `eopen`
- `els`
- `ecat`
- `cate`
- `ediff`
- `einfo`

将来拡張候補:
- `eclear`
- `eexport`
- `eimport`
- `eselect`

---

## 11. コマンド仕様

### 11.1 eopen
#### 概要
指定 Excel ブックを開く、または既に開かれている場合はそれを操作対象に設定する。

#### 形式
```text
eopen <excel-file> [--sheet <sheet-name>]
```

#### 引数
- `<excel-file>` : 対象 Excel ファイル
- `--sheet <sheet-name>` : 既定シートを指定する

#### 動作
- 指定ファイルが未オープンなら Excel で開く
- 既に開かれているなら既存ブックを対象にする
- セッションへ対象ブックを登録する
- Save / Close は行わない

---

### 11.2 els
#### 概要
現在対象ブック、または指定シート内のテキストボックス一覧を表示する。

#### 形式
```text
els [--sheet <sheet-name>]
```

#### 出力
- シート名
- Shape 名
- 必要に応じて型情報や簡易情報

#### 例
```text
[Main]
txtInput
txtOutput
txtLog
```

---

### 11.3 ecat
#### 概要
指定テキストボックスの内容を標準出力へ出力する。

#### 形式
```text
ecat <textbox-id>
```

#### 動作
- 対象テキストを取得する
- 標準出力へ送る
- Save / Close は行わない

#### 利用例
PowerShell:
```powershell
ecat Main:txtInput
```

WSL:
```bash
ecat Main:txtInput | nl
ecat Main:txtInput | grep "ERROR"
```

---

### 11.4 cate
#### 概要
標準入力の内容を指定テキストボックスへ書き込む。

#### 形式
```text
cate <textbox-id> [--append]
```

#### 動作
- 標準入力をすべて読み込む
- デフォルトは上書き
- `--append` 指定時は追記
- Save / Close は行わない

#### 利用例
PowerShell:
```powershell
Get-Content .\memo.txt | cate Main:txtOutput
```

WSL:
```bash
cat memo.txt | cate Main:txtOutput
ecat Main:txtInput | sort | cate Main:txtSorted
```

---

### 11.5 ediff
#### 概要
2 つのテキストボックス内容を比較して差分を出力する。

#### 形式
```text
ediff <textbox-id-1> <textbox-id-2>
```

#### 動作
- 2 つのテキスト内容を取得する
- 一時ファイルへ出力する
- 差分を計算して標準出力へ出力する
- 実装上は WSL `diff` の利用を許容する
- 一時ファイルは処理後に削除する

#### 備考
WSL では次のような直接利用も想定する。

```bash
diff <(ecat Main:left) <(ecat Main:right)
```

このため `ediff` は便利コマンドとしての意味合いも持つ。

---

### 11.6 einfo
#### 概要
現在のセッション情報を表示する。

#### 形式
```text
einfo
```

#### 出力例
```text
Workbook : D:\work\sample.xlsx
Sheet    : Main
Excel    : Running
```

---

## 12. 利用例

### 12.1 PowerShell 単独利用
```powershell
eopen D:\work\sample.xlsx --sheet Main
els
ecat txtInput
Get-Content .\memo.txt | cate txtOutput --append
```

### 12.2 PowerShell + 補助関数
```powershell
ecat txtInput | unl
ecat txtInput | ugrep "TODO"
```

### 12.3 WSL 主体利用
```bash
eopen /mnt/d/work/sample.xlsx --sheet Main
els
ecat txtInput | nl
ecat txtInput | grep "ERROR"
ecat txtInput | sort | cate txtSorted
```

### 12.4 WSL での差分
```bash
ediff left right
diff <(ecat left) <(ecat right)
```

---

## 13. WSL ラッパー方針

WSL では `alias` または `function` で Exshell コマンドを公開してよい。  
実運用では `function` を推奨する。

例:
```bash
EXSHELL_DIR="/mnt/c/tools/exshell"

eopen() { "$EXSHELL_DIR/eopen.exe" "$@"; }
els()   { "$EXSHELL_DIR/els.exe" "$@"; }
ecat()  { "$EXSHELL_DIR/ecat.exe" "$@"; }
cate()  { "$EXSHELL_DIR/cate.exe" "$@"; }
ediff() { "$EXSHELL_DIR/ediff.exe" "$@"; }
einfo() { "$EXSHELL_DIR/einfo.exe" "$@"; }
```

---

## 14. PowerShell 補助関数方針

PowerShell では WSL コマンド補助関数を定義してよい。

例:
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

---

## 15. エラー仕様

### 15.1 基本方針
- 正常終了時は exit code 0
- エラー時は非 0
- エラーメッセージは stderr に出力する

### 15.2 想定エラー
- 引数不正
- セッション未確立
- 対象ブック未検出
- シート未検出
- テキストボックス未検出
- Excel 操作失敗
- 標準入力処理失敗
- 一時ファイル作成失敗
- 差分処理失敗

---

## 16. 将来拡張

- セルとの入出力連携
- 名前定義範囲との連携
- 複数ブック対応
- クリップボード連携
- パイプライン支援コマンド
- Excel 側メタ情報取得
- ログ出力、デバッグモード
- スクリプト実行補助

---

## 17. まとめ

Exshell v0.3 では、PowerShell / WSL の両対応を正式方針とする。  
Excel を GUI 入出力、Exshell をブリッジ、WSL を Unix コマンド実行基盤として位置づけ、どちらのシェルからも同等に使える CLI を目指す。
