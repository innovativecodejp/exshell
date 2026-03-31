# Exshell コマンド実装方針メモ

## 1. 文書の目的

本メモは、Exshell を実装する際の実装順序、実装戦略、注意点、暫定判断基準をまとめた開発用メモである。  
外部仕様書および内部設計案と矛盾しない範囲で、実装上の現実的な進め方を示す。

---

## 2. 実装方針の基本姿勢

Exshell は独自機能を増やしすぎず、以下の原則で実装する。

- Excel 固有処理のみを Exshell が担う
- テキスト加工や差分計算は既存 CLI に委譲する
- PowerShell / WSL / Excel の境界を曖昧にしない
- 初期版では「少機能だが壊れにくい」ことを優先する

---

## 3. 先に固定しておく実装前提

- 言語: C#
- ランタイム: .NET Framework 4.8
- IDE: Visual Studio 2022
- Excel: インストール済みデスクトップ版
- 外部ツール: WSL
- 利用シェル: PowerShell
- WSL補助コマンド系: `uls`, `ucat`, `unl`, `udiff` 文化を採用

---

## 4. 実装優先順位

### 4.1 最優先
以下を先に動かす。

1. セッション読込／保存
2. `eopen`
3. `einfo`
4. Shape 名でのテキスト取得
5. `ecat`
6. `cate`（上書き）
7. `cate --append`

理由:
- 最小の価値がここで成立するため
- Excel を CLI 入出力窓として使える状態になるため

### 4.2 次点
以下を次段階とする。

1. `els`
2. 一時ファイル出力
3. Windows → WSL パス変換
4. `wsl diff` 呼出
5. `ediff`

理由:
- `ediff` は Exshell の特徴的機能だが、依存要素がやや多い
- 先に `ecat` / `cate` を安定させた方が全体が進めやすい

### 4.3 後回しでよいもの
- `eimport`
- `eexport`
- `epipe`
- セル対応
- ActiveX対応
- 色付き出力
- 高度なログ
- 独自 diff 実装

---

## 5. コマンド別実装メモ

### 5.1 eopen
- 入力は Excel ファイルパス
- 絶対パス化して保持する
- 既に開いている Workbook があれば再利用する
- 開いていなければ Excel で開く
- セッションファイルへ保存する

注意点:
- `Workbooks.Open()` 時の引数は最小限にする
- 読み取り専用制御は初期版では深入りしない
- Save / Close を実行しない

### 5.2 ecat
- Shape 名を `Sheet:ShapeName` 形式で解析する
- シート省略時はセッション既定シートを使う
- Shape テキストを取得して標準出力へ流す

注意点:
- Shape が存在してもテキストを持てないことがある
- `TextFrame2` と `TextFrame` の両対応を入れる
- 出力終端改行の扱いを先に決めてぶらさない

### 5.3 cate
- 標準入力全文を読み込む
- デフォルトは上書き
- `--append` で末尾追記
- Save しない

注意点:
- パイプ入力が空のケース
- Excel 側既存改行と追記テキストの間の改行ルール
- 追記時の先頭改行有無を将来問題にしやすい

初期版推奨:
`--append` は単純連結とする。

```text
新値 = 既存値 + 入力値
```

### 5.4 els
- 対象シートの Shape を列挙
- テキスト取得可能 Shape のみ表示

注意点:
- Shape の種類差
- グループ図形混在
- 名前順で並べた方が使いやすい

初期版表示例:

```text
[Main]
txtInput
txtOutput
left
right
```

### 5.5 ediff
- 2つの Shape テキスト取得
- temp1, temp2 生成
- UTF-8 LF で書き込む
- `wsl diff temp1 temp2` 実行
- 標準出力へ返す
- temp 削除

注意点:
- WSL パス変換
- 日本語文字化け
- temp 後始末
- `diff` の終了コードは差分があると非0になり得る

重要:
`diff` は「差分がある」場合も 0 以外を返すことがあるため、  
**終了コードだけで異常判定しない** 設計が必要。

---

## 6. COM 実装方針メモ

### 6.1 Excel.Application の扱い
- 可能な限り短時間だけ参照を持つ
- コマンドごとに解決し直す方が安全
- グローバル長寿命保持は避ける

### 6.2 Workbook 解決
- パス一致で既存 Workbook を探す
- 見つからなければ Open する
- ファイル名一致ではなく、**フルパス一致** で判定する

### 6.3 Shape テキスト取得
推奨順:

1. `shape.TextFrame2.TextRange.Text`
2. `shape.TextFrame.Characters().Text`

### 6.4 Shape 判定
- まず Shape が存在するか
- 次にテキスト取得可能か
- 取得不能なら「非対応 Shape」として扱う

---

## 7. 文字コード・改行方針

### 7.1 標準方針
- .NET 内部文字列は通常の `string`
- temp ファイルは UTF-8
- WSL へ渡す temp は LF で保存
- Excel 表示は必要に応じて改行変換を受け入れる

### 7.2 注意
Excel から取った文字列に CRLF が含まれる場合、`ediff` 用 temp 書込時には LF 正規化した方が安定する。

推奨:

```text
CRLF → LF
CR   → LF
```

---

## 8. Path 変換実装メモ

Windows パス:

```text
C:\Users\user\AppData\Local\Temp\a.txt
```

WSL パス:

```text
/mnt/c/Users/user/AppData/Local/Temp/a.txt
```

実装方針:
- ドライブレターを小文字へ
- `\` を `/` に置換
- 先頭を `/mnt/<drive>/` に変換
- UNC パス対応は初期版では対象外でもよい

---

## 9. セッション実装メモ

### 9.1 保存先
初期版は `%APPDATA%\Exshell\session.json` 推奨

理由:
- カレントディレクトリに依存しない
- ユーザー単位で一貫する

### 9.2 実装方針
- アプリ起動時に毎回ロード
- `eopen` 成功時に保存
- セッションファイル不在時は「未確立」とする

---

## 10. エラー処理メモ

### 10.1 方針
- 想定可能エラーは明示メッセージで返す
- `catch (Exception)` にまとめすぎない
- stderr と exit code を必ず整合させる

### 10.2 メッセージ例
- セッション未確立: `No active Excel session. Run eopen first.`
- シート未存在: `Sheet not found: Main`
- Shape未存在: `Shape not found: txtInput`
- 非対応Shape: `Shape does not provide text content: xxx`

---

## 11. テスト実施順メモ

### 11.1 最初にやるべき確認
1. Excel が未起動でも `eopen` で開けるか
2. `eopen` 後に `einfo` が正しいか
3. `ecat` で日本語を読めるか
4. `cate` で日本語を上書きできるか
5. `cate --append` が正しく追記できるか

### 11.2 その次
1. `els` が必要 Shape を列挙できるか
2. `ediff` が temp を作れるか
3. `wsl diff` が実行できるか
4. 改行差だけの差分が意図どおりか

---

## 12. 実装順サンプル

### Step 1
- `SessionInfo`
- `JsonSessionStore`

### Step 2
- `ShapeReference`
- `ArgumentParser`

### Step 3
- `ExcelAppGateway`
- `WorkbookResolver`
- `WorksheetResolver`

### Step 4
- `ShapeResolver`
- `ShapeTextAccessor`

### Step 5
- `eopen`
- `einfo`

### Step 6
- `ecat`

### Step 7
- `cate`
- `cate --append`

### Step 8
- `els`

### Step 9
- `TempFileService`
- `PathConverter`
- `ProcessRunner`

### Step 10
- `ediff`

---

## 13. 実装時の割り切り

初期版では次を割り切ってよい。

- すべての Shape 種類に対応しない
- セルは扱わない
- 複数ブック同時管理はしない
- `diff` の詳細オプション透過は後回し
- PowerShell profile 自動設定はしない

この割り切りにより、早く価値を出せる。

---

## 14. 将来拡張メモ

- `ediff --unified`
- `ediff --ignore-space`
- `ecat --file`
- `cate --file`
- `epipe`
- `eimport`
- `eexport`
- `els --verbose`
- シート内の座標・種類表示
- 特定 prefix の Shape 一括操作

---

## 15. まとめ

Exshell 実装では、まず `eopen`, `ecat`, `cate`, `els`, `ediff` を最小構成で成立させることが重要である。  
特に以下を優先する。

- Excel COM 処理の安定化
- Shape テキスト取得／設定の吸収
- セッション明確化
- `ediff` の temp + `wsl diff` 化
- Save / Close を行わない責務分離

独自実装を増やしすぎず、WSL の既存資産を活かすことで、Exshell は軽量で透明な CLI ツールとして成立しやすくなる。
