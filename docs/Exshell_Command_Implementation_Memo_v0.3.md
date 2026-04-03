# Exshell コマンド実装方針メモ v0.3

## 1. 文書情報
- 文書名: Exshell コマンド実装方針メモ
- 版数: v0.3
- 目的: 外部仕様書 v0.3 / 内部設計案 v0.3 に基づく実装着手方針をまとめる

---

## 2. 実装方針の基本

v0.3 では、PowerShell / WSL の両対応を正式方針とする。  
したがって、実装では次を優先する。

- Exshell 本体だけで主要機能が完結すること
- PowerShell と WSL の双方から違和感なく使えること
- Excel と標準入出力の橋渡しに責務を絞ること
- Unix コマンドの再実装を避けること

---

## 3. MVP 範囲

まず実装すべきコマンドは以下とする。

1. `eopen`
2. `els`
3. `ecat`
4. `cate`
5. `einfo`
6. `ediff`

この順でよい。

---

## 4. 推奨実装順序

### Step 1: 共通基盤
最初に以下を作る。

- 引数解析の枠組み
- `SessionInfo`
- `TextBoxRef`
- `PathResolver`
- `JsonSessionStore`

ここで WSL パス変換を先に入れておくことが重要。

#### 先に確認すべき点
- `/mnt/d/...` を正しく `D:\...` に変換できるか
- UTF-8 の入出力が安定するか

---

### Step 2: Excel 読み取り
次に以下を作る。

- `ExcelLocator`
- `ExcelBridge.OpenOrAttach`
- `ExcelBridge.ListTextBoxes`
- `ExcelBridge.ReadText`

この段階で `eopen`, `els`, `ecat` を通せるようにする。

#### ここで確認すること
- 既に開いているブックを捕まえられるか
- Shape 名の列挙が安定するか
- 日本語の Shape 名でも問題ないか
- TextFrame / TextFrame2 の差を吸収できるか

---

### Step 3: Excel 書き込み
次に以下を作る。

- `ExcelBridge.WriteText`
- `cate`

#### 確認項目
- 上書き動作
- `--append` 動作
- 複数行テキスト
- 改行の見え方
- Save を呼んでいないこと

---

### Step 4: セッション表示
- `einfo`

これは比較的容易なので途中でもよい。

---

### Step 5: 差分
- `TempFileService`
- `IDiffService`
- `ediff`

#### 初期方針
まずは一時ファイルを作るところまで実装する。  
その後、差分取得の方法を選ぶ。

候補:
1. 自前で簡易差分
2. 外部 diff 呼び出し

初期版は WSL `diff` 呼び出しでもよいが、将来差し替えできる構造にしておく。

---

## 5. PowerShell / WSL 両対応の考え方

### 5.1 本体で吸収する範囲
本体で吸収するべきもの:
- Windows パス / WSL パス
- stdin / stdout
- 文字コード
- 一時ファイル生成
- セッション管理

### 5.2 ラッパーで吸収しない方がよいもの
- ビジネスロジック
- Excel オブジェクト処理
- テキストボックス解決
- セッションファイル操作

WSL の `alias` / `function` はあくまで起動簡略化だけに留める。

---

## 6. PowerShell 側の利用方針

PowerShell では Exshell 本体を直接使えることを重視する。

例:
```powershell
eopen D:\work\sample.xlsx --sheet Main
els
ecat txtInput
Get-Content .\memo.txt | cate txtOutput --append
```

PowerShell で Unix コマンド風操作をしたい場合は、別途 profile に補助関数を用意する。

例:
```powershell
function uls   { wsl ls @Args }
function ucat  { wsl cat @Args }
function unl   { wsl nl @Args }
function udiff { wsl diff @Args }
function ugrep { wsl grep @Args }
```

---

## 7. WSL 側の利用方針

WSL では Unix コマンドをそのまま使うことを重視する。

例:
```bash
eopen /mnt/d/work/sample.xlsx --sheet Main
els
ecat txtInput | nl
ecat txtInput | grep "ERROR"
ecat txtInput | sort | cate txtSorted
```

ラッパー例:
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

## 8. コマンド別実装メモ

### 8.1 eopen
#### 実装優先度
最優先

#### ポイント
- パス正規化を最初に行う
- 既存 Workbook を探す
- なければ開く
- セッション保存

#### 注意
- Save/Close はしない
- 相対パス扱いを明確にする

---

### 8.2 els
#### 実装優先度
高

#### ポイント
- セッションから既定ブックを読む
- シート指定があれば優先
- `Shapes` から文字列可能 Shape のみ抽出

#### 注意
- 列挙順が変動しないか確認する
- 名前一覧をパイプ利用しやすい形式にする

---

### 8.3 ecat
#### 実装優先度
高

#### ポイント
- stdout に本文のみ出す
- 不要な飾りやログを混ぜない

#### 注意
- 末尾改行の扱いを早めに決める
- WSL パイプで文字化けしないことを確認する

---

### 8.4 cate
#### 実装優先度
高

#### ポイント
- stdin を全読込
- `--append` の有無で分岐
- 既存テキスト取得 + 結合 + 書込

#### 注意
- 空入力をどう扱うか
- 改行結合ルールを固定する
- Save をしないことをコードレビュー観点に入れる

---

### 8.5 einfo
#### 実装優先度
中

#### ポイント
- デバッグ確認にも使える
- セッション状態を見せるだけなので軽い

---

### 8.6 ediff
#### 実装優先度
中

#### ポイント
- 2 つのテキスト取得
- 一時ファイル作成
- UTF-8 LF で保存
- 差分結果出力
- 一時ファイル削除

#### 注意
- PowerShell / WSL で差分実装を差し替えたくなる可能性がある
- まずは内部インターフェースを切っておく

---

## 9. 実装時の具体注意

### 9.1 COM 解放
COM オブジェクト放置は Excel プロセス残留の原因になる。  
`Marshal.ReleaseComObject` 等を適切に行う。

### 9.2 Excel の所有権
Exshell は Excel の所有者ではない。  
そのため以下は禁止方針にする。

- Save 呼出し
- Close 呼出し
- Quit 呼出し

### 9.3 デバッグ出力
stdout に混ぜない。  
必要なら stderr またはログファイルへ出す。

### 9.4 文字コード
日本語を含むので UTF-8 を明示する。  
特に WSL パイプ接続で早めに実機確認する。

### 9.5 改行
差分やソートのため、テキスト出力は LF 系が扱いやすい。  
ただし Excel 表示時は必要に応じて調整が要る。

---

## 10. 最小動作確認シナリオ

### シナリオ1: PowerShell 基本
```powershell
eopen D:\work\sample.xlsx --sheet Main
els
ecat txtInput
```

### シナリオ2: PowerShell 書込み
```powershell
Get-Content .\memo.txt | cate txtOutput
Get-Content .\memo.txt | cate txtOutput --append
```

### シナリオ3: WSL 基本
```bash
eopen /mnt/d/work/sample.xlsx --sheet Main
els
ecat txtInput | nl
```

### シナリオ4: WSL パイプ
```bash
ecat txtInput | grep "ERROR"
ecat txtInput | sort | cate txtSorted
```

### シナリオ5: 差分
```bash
ediff left right
diff <(ecat left) <(ecat right)
```

---

## 11. ロードマップ案

### Phase 1
- PathResolver
- SessionStore
- eopen
- els
- ecat

### Phase 2
- cate
- einfo
- PowerShell 動作確認
- WSL 動作確認

### Phase 3
- ediff
- 一時ファイル整理
- 例外整理
- 終了コード整理

### Phase 4
- 補助ドキュメント整備
- WSL ラッパー例
- PowerShell profile 例
- 実運用チューニング

---

## 12. 実装判断の基準

迷った場合は以下の順で判断する。

1. PowerShell / WSL の両方から使えるか
2. stdout / stdin の合成可能性を壊していないか
3. Excel の保存・閉鎖責務を侵していないか
4. Unix コマンドの再実装になっていないか
5. 実装を単純に保てているか

---

## 13. まとめ

v0.3 の実装では、Exshell をシェル非依存の Excel ブリッジ CLI として構築する。  
PowerShell では直接利用、WSL では Unix コマンドと組み合わせて利用する。  
本体は両方の入口を等価に扱い、ラッパーは最小限に留める。
