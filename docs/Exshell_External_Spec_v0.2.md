# Exshell 外部仕様書 v0.2

## 1. 文書情報

- 文書名: Exshell 外部仕様書
- バージョン: v0.2
- 対象システム名: Exshell
- 作成目的: Exshell の外部仕様を定義し、利用者視点での動作・制約・利用方法を明確化する
- 本版の位置づけ: v0.1 をベースに、**WSL ターミナル主体運用**を前提とする改訂版

---

## 2. システム概要

Exshell は、Excel 上のテキストボックスを入出力バッファとして扱うための CLI ツール群である。  
Exshell 自体は Windows 上で動作し、Excel デスクトップアプリケーションを操作する。  
一方、利用者の主運用環境は **WSL ターミナル** とし、`ls`、`cat`、`grep`、`nl`、`diff` などの Unix コマンドをメインに使用する。

Exshell は、以下の橋渡しを担う。

- WSL / Unix CLI 世界
- Windows / Excel 世界

すなわち、Exshell は **Excel-CLI ブリッジ** として機能する。

---

## 3. 設計方針

### 3.1 基本方針

- Unix コマンドは WSL 側の標準コマンドをそのまま使用する
- Exshell は Excel 入出力に特化する
- Excel の保存・クローズ・終了は Exshell では行わない
- Exshell のコマンドは WSL ターミナルから呼び出して使用することを主運用とする
- PowerShell からの利用も可能だが、補助運用と位置づける

### 3.2 役割分担

#### Exshell の責務
- Excel ブックを開く
- Excel テキストボックスを読む
- Excel テキストボックスへ書く
- Excel テキストボックス一覧を取得する
- 必要に応じて一時ファイルを生成し、Excel 内容を WSL コマンドに受け渡す

#### WSL / Unix コマンドの責務
- ファイル一覧表示
- テキスト表示
- 行番号付与
- grep / sed / awk 等による加工
- diff による比較
- パイプ、リダイレクト、プロセス置換などのシェル機能

---

## 4. 想定実行環境

### 4.1 OS / 実行基盤
- Windows 11
- Excel デスクトップ版インストール済み
- WSL 利用可能
- 主運用ターミナル: WSL ターミナル
- 補助運用ターミナル: PowerShell

### 4.2 実装前提
- Exshell は Windows 実行ファイルとして実装する
- Excel 操作は COM Interop により行う
- Exshell コマンドは WSL から Windows 実行ファイルとして呼び出せることを前提とする

### 4.3 主運用形態
利用者は WSL ターミナル上で、以下のように Unix コマンドを直接使用する。

```bash
ls
grep "TODO" file.txt
ecat Main:src | nl
ecat Main:src | grep "ERROR"
diff <(ecat Main:left) <(ecat Main:right)
```

---

## 5. 対象データ

### 5.1 対象 Excel ファイル
- `.xlsx`
- `.xlsm`

### 5.2 対象オブジェクト
Exshell が操作対象とするのは、Excel ワークシート上の **Shape テキストボックス** とする。

### 5.3 非対象
以下は初期版では対象外とする。

- ActiveX テキストボックス
- フォームコントロール
- セルそのもの
- コメント / ノート
- グループ化された複雑図形内部の個別要素

---

## 6. テキストボックス識別仕様

### 6.1 正式識別子
テキストボックスの正式識別子は以下の形式とする。

```text
<SheetName>:<ShapeName>
```

例:
```text
Main:txtInput
Main:txtOutput
Diff:left
Diff:right
```

### 6.2 シート省略時
シート省略時は、現在の既定シートを対象とする。

例:
```bash
ecat txtInput
cate txtOutput
```

### 6.3 一意性
同一シート内では `ShapeName` は一意であることを前提とする。

---

## 7. セッション仕様

### 7.1 セッションの目的
`eopen` 実行後、対象ブックおよび既定シートを保持し、以後の `ecat` / `cate` / `els` / `ediff` に利用する。

### 7.2 セッション保持項目
- 対象ブックフルパス
- 既定シート名
- 必要に応じて Excel 接続情報

### 7.3 セッションの性質
- Exshell はセッションを保持する
- ただし、保存・クローズ・Excel 終了は行わない
- セッションは「操作対象の記録」であり、Excel の所有権管理ではない

---

## 8. 保存・クローズ方針

### 8.1 保存
- Exshell は `Save` を実行しない
- `cate` 等で内容が変更された場合も、自動保存しない
- 保存判断は Excel 側の操作で行う

### 8.2 クローズ
- Exshell は `Close` を実行しない
- Exshell は Excel アプリケーションの `Quit` を実行しない
- ブックを閉じるかどうかは Excel 側の操作で行う

### 8.3 責務分離
Exshell は Excel 内容の参照・更新を行うが、文書ライフサイクル管理は行わない。

---

## 9. コマンド一覧

初期版で定義するコマンドは以下とする。

- `eopen`
- `els`
- `einfo`
- `ecat`
- `cate`
- `ediff`

---

## 10. 各コマンド仕様

### 10.1 eopen

#### 概要
指定 Excel ファイルを開く、または既に開いていればそれを対象にする。  
あわせて Exshell の現在セッションを設定する。

#### 書式
```bash
eopen <excel-file> [--sheet <sheet-name>]
```

#### 引数
- `<excel-file>`: Excel ファイルパス
- `--sheet <sheet-name>`: 既定シート指定

#### 動作
- 指定ファイルが未オープンなら Excel で開く
- 既に開いている場合はそれを捕捉する
- セッションに対象ブックを設定する
- `--sheet` 指定時は既定シートも設定する

#### 備考
- 保存はしない
- クローズはしない

#### 例
```bash
eopen /mnt/d/work/sample.xlsx
eopen D:\work\sample.xlsx --sheet Main
```

---

### 10.2 els

#### 概要
対象シートのテキストボックス一覧を表示する。

#### 書式
```bash
els [--sheet <sheet-name>]
```

#### 動作
- 指定シート、または既定シートの Shape テキストボックス一覧を標準出力に出す

#### 出力例
```text
[Main]
txtInput
txtOutput
left
right
```

---

### 10.3 einfo

#### 概要
現在セッション情報を表示する。

#### 書式
```bash
einfo
```

#### 出力例
```text
Workbook : D:\work\sample.xlsx
Sheet    : Main
Excel    : Running
```

---

### 10.4 ecat

#### 概要
指定テキストボックス内容を標準出力へ出力する。

#### 書式
```bash
ecat <textbox-id>
```

#### 動作
- 指定テキストボックス内容を取得する
- 標準出力へ出力する

#### 用途
- Unix パイプラインへ流す
- `grep`, `nl`, `sed`, `awk`, `sort`, `diff` 等に接続する

#### 例
```bash
ecat Main:src
ecat Main:src | nl
ecat Main:src | grep "ERROR"
```

---

### 10.5 cate

#### 概要
標準入力を指定テキストボックスへ出力する。

#### 書式
```bash
cate <textbox-id> [--append]
```

#### 動作
- 標準入力を読み込む
- 指定テキストボックスへ反映する
- デフォルトは上書き
- `--append` 指定時は追記

#### 備考
- Save は行わない
- Close は行わない

#### 例
```bash
cat memo.txt | cate Main:txtInput
ecat Main:list | sort | cate Main:sorted
ecat Main:log | grep "ERROR" | cate Main:errorOnly
ecat Main:note | sed 's/foo/bar/g' | cate Main:note2 --append
```

---

### 10.6 ediff

#### 概要
2つのテキストボックス内容の差分を表示する。

#### 書式
```bash
ediff <textbox-id-1> <textbox-id-2>
```

#### 基本方針
`ediff` は簡便コマンドとして提供する。  
ただし、WSL 主体運用では以下の Unix 的な使用も推奨される。

```bash
diff <(ecat Main:left) <(ecat Main:right)
```

#### 動作
- 指定2テキストボックス内容を取得する
- 一時ファイルを生成する
- WSL `diff` を内部実行する
- 結果を標準出力に出す
- 一時ファイルは処理後に削除する

#### 一時ファイル仕様
- 文字コード: UTF-8
- 改行コード: LF

#### 例
```bash
ediff Main:left Main:right
diff <(ecat Main:left) <(ecat Main:right)
```

---

## 11. パス仕様

### 11.1 基本方針
WSL ターミナル主体運用のため、`eopen` では WSL パス入力を許容する。  
同時に、Windows パスも受け付ける。

### 11.2 許容形式
- Windows パス
- WSL パス

例:
```bash
eopen D:\work\sample.xlsx
eopen /mnt/d/work/sample.xlsx
```

### 11.3 内部変換
WSL パスが入力された場合、Exshell 内部で Windows パスへ変換して Excel に渡す。

---

## 12. 標準入出力仕様

### 12.1 標準出力
- `ecat`, `els`, `einfo`, `ediff` は標準出力を利用する

### 12.2 標準入力
- `cate` は標準入力を利用する

### 12.3 文字コード
- 標準入出力は UTF-8 を基本とする

### 12.4 改行
- WSL / Unix コマンド連携を優先し、LF ベースでの運用を基本とする
- Excel 表示時の改行保持は実装時に適切に変換する

---

## 13. エラー仕様

### 13.1 エラー出力
- エラーメッセージは標準エラー出力へ出す

### 13.2 終了コード
- 0: 正常終了
- 1: 引数エラー
- 2: セッション未設定
- 3: 対象ブック未検出
- 4: 対象シート未検出
- 5: 対象テキストボックス未検出
- 6: Excel 操作失敗
- 7: WSL / 外部コマンド実行失敗

---

## 14. 想定利用例

### 14.1 単純表示
```bash
ecat Main:src
```

### 14.2 行番号付き表示
```bash
ecat Main:src | nl
```

### 14.3 grep
```bash
ecat Main:src | grep "SELECT"
```

### 14.4 整形して別ボックスへ格納
```bash
ecat Main:list | sort | cate Main:sorted
```

### 14.5 差分表示
```bash
ediff Main:left Main:right
```

### 14.6 Unix 方式の差分表示
```bash
diff <(ecat Main:left) <(ecat Main:right)
```

---

## 15. PowerShell 補助運用

### 15.1 位置づけ
PowerShell からの利用は補助運用とする。

### 15.2 補助関数
PowerShell では、必要に応じて WSL コマンド呼び出しの補助関数を定義してよい。

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

### 15.3 補助運用の考え方
- 主運用は WSL ターミナル
- `uls` 系は PowerShell で同等操作感を得るための補助手段

---

## 16. 将来拡張候補

- `eexport`: テキストボックス内容をファイルへ出力
- `eimport`: ファイル内容をテキストボックスへ入力
- `eclear`: テキストボックス内容クリア
- `eselect`: 既定シート変更
- `epipe`: テキストボックス内容を外部コマンド処理し別ボックスへ反映
- セル入出力対応
- 追加オブジェクト種別対応

---

## 17. MVP 範囲

初期実装範囲は以下とする。

- `eopen`
- `els`
- `einfo`
- `ecat`
- `cate`
- `ediff`

かつ、以下の条件に限定する。

- Shape テキストボックスのみ対象
- 単一ブックセッション
- 既定シートまたは明示シート指定
- 保存・クローズは行わない
- WSL 主体運用を前提とする

---

## 18. まとめ

Exshell v0.2 は、Excel を Unix CLI ワークフローへ接続するためのブリッジツールとして定義する。  
主運用環境は WSL ターミナルであり、利用者は Unix コマンドを通常どおり使用する。  
Exshell は Excel の開閉や保存を管理せず、Excel テキストボックスと標準入出力の接続に特化する。

これにより、以下を両立する。

- Excel を UI / 入出力面として利用する
- Unix コマンドの軽快さと透明性を活用する
- PowerShell / WSL / Excel を一体的な研究・開発環境として統合する
