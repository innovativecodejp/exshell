# Exshell 内部設計案 v0.3

## 1. 文書情報
- 文書名: Exshell 内部設計案
- 版数: v0.3
- 対象: Exshell CLI 実装
- 目的: 外部仕様書 v0.3 に対応する内部設計の骨子を定義する

---

## 2. 設計方針

Exshell は Windows 上の CLI 実行ファイルとして実装し、PowerShell / WSL の双方から同等に利用できる構造とする。  
内部設計では以下を重視する。

- CLI 層と Excel 操作層の分離
- PowerShell / WSL の差異を本体で極力吸収する
- Excel COM 操作を局所化する
- 標準入力・標準出力中心の単純なデータフローを維持する
- Save / Close を一切行わない責務分離
- テストしやすい構成にする

---

## 3. 想定実装構成

### 3.1 実装言語
- C#
- .NET Framework 4.8 を第一候補とする

### 3.2 主要技術
- Microsoft.Office.Interop.Excel
- 標準入出力処理
- ローカル JSON セッション保存
- 必要に応じて外部プロセス実行

---

## 4. 全体アーキテクチャ

以下の層構造を採用する。

1. CLI 層
2. アプリケーション層
3. Excel ドメイン操作層
4. インフラ層
5. 補助サービス層

### 4.1 CLI 層
責務:
- 引数解析
- サブコマンド判定
- 標準入力読み込み
- 標準出力 / 標準エラー出力
- 終了コード制御

### 4.2 アプリケーション層
責務:
- 各コマンドのユースケース実行
- セッション参照
- 引数と実処理の橋渡し

### 4.3 Excel ドメイン操作層
責務:
- ブック解決
- シート解決
- Shape 解決
- テキスト取得・更新

### 4.4 インフラ層
責務:
- COM Interop 呼び出し
- JSON 永続化
- 一時ファイル操作
- 外部プロセス呼び出し

### 4.5 補助サービス層
責務:
- パス変換
- テキストボックス識別子解析
- 文字コード正規化
- 改行正規化

---

## 5. 推奨プロジェクト構成

```text
Exshell/
  Exshell.Cli/
    Program.cs
    CommandDispatcher.cs
    Commands/
      EopenCommand.cs
      ElsCommand.cs
      EcatCommand.cs
      CateCommand.cs
      EdiffCommand.cs
      EinfoCommand.cs
  Exshell.Application/
    UseCases/
      EopenUseCase.cs
      ElsUseCase.cs
      EcatUseCase.cs
      CateUseCase.cs
      EdiffUseCase.cs
      EinfoUseCase.cs
    Dto/
  Exshell.Core/
    Models/
      SessionInfo.cs
      TextBoxRef.cs
      WorkbookRef.cs
    Services/
      IExcelBridge.cs
      ISessionStore.cs
      IPathResolver.cs
      ITempFileService.cs
      IDiffService.cs
  Exshell.Infrastructure/
    Excel/
      ExcelBridge.cs
      ExcelLocator.cs
      ShapeTextAccessor.cs
    Session/
      JsonSessionStore.cs
    Paths/
      PathResolver.cs
    Temp/
      TempFileService.cs
    Diff/
      DiffService.cs
  Exshell.Tests/
```

単一プロジェクトでも着手可能だが、責務は上記のように分離することを推奨する。

---

## 6. 主要モデル

### 6.1 SessionInfo
保持内容:
- WorkbookPath
- DefaultSheet
- LastUpdated

### 6.2 TextBoxRef
保持内容:
- SheetName
- ShapeName

### 6.3 WorkbookRef
保持内容:
- FullPath
- FileName

---

## 7. コアインターフェース

### 7.1 IExcelBridge
主責務:
- Workbook を開く
- Workbook を取得する
- TextBox 一覧を取得する
- TextBox 内容を取得する
- TextBox 内容を更新する

想定メソッド:
- `WorkbookRef OpenOrAttach(string path)`
- `IEnumerable<TextBoxRef> ListTextBoxes(string workbookPath, string? sheetName)`
- `string ReadText(string workbookPath, TextBoxRef target)`
- `void WriteText(string workbookPath, TextBoxRef target, string content, bool append)`

### 7.2 ISessionStore
主責務:
- セッション読込
- セッション保存
- セッションクリア

### 7.3 IPathResolver
主責務:
- Windows パス解決
- WSL パス解決
- 相対パス解決

### 7.4 ITempFileService
主責務:
- 一時ファイル生成
- 一時ファイル削除

### 7.5 IDiffService
主責務:
- 2 テキストの差分取得
- 必要に応じて外部 diff 連携

---

## 8. CLI 設計

### 8.1 コマンド形式
初期版は以下のどちらでも実装可能とする。

#### 方式 A: 単一 exe + サブコマンド
```text
exshell eopen ...
exshell ecat ...
```

#### 方式 B: 個別 exe
```text
eopen ...
ecat ...
cate ...
els ...
ediff ...
einfo ...
```

外部仕様との親和性から、運用上は方式 B を優先してよい。  
内部的には共通の `CommandDispatcher` を持つ。

---

## 9. パス解決設計

### 9.1 目的
PowerShell / WSL の双方から同じ Exshell 本体を呼び出せるようにする。

### 9.2 入力例
- `D:\work\sample.xlsx`
- `d:/work/sample.xlsx`
- `/mnt/d/work/sample.xlsx`

### 9.3 解決方針
`PathResolver` は次を行う。

- Windows パスはそのまま正規化
- `/mnt/<drive>/...` は `<Drive>:\...` へ変換
- 引用符除去
- 相対パスの絶対化

### 9.4 変換例
- `/mnt/d/work/sample.xlsx` → `D:\work\sample.xlsx`
- `/mnt/c/Users/me/file.xlsx` → `C:\Users\me\file.xlsx`

---

## 10. Excel ブリッジ設計

### 10.1 ExcelLocator
責務:
- 既存 Excel.Application を探索する
- 必要に応じて Excel.Application を起動する
- 対象 Workbook を特定する

### 10.2 Workbook 解決方針
`eopen` では以下の順に解決する。

1. 開いているブック群からフルパス一致を探す
2. 見つからなければ Excel で開く
3. セッションへ保存する

### 10.3 Shape 解決方針
対象シートの `Shapes` を列挙し、次を満たす Shape を候補とする。

- テキストを保持できる
- 名前が一致する

必要に応じて `TextFrame2` を優先、不可なら `TextFrame` を使用する。

### 10.4 Save / Close 非実行
`ExcelBridge` は以下を提供しない、または呼び出さない。

- Save
- Close
- Quit

---

## 11. セッション設計

### 11.1 保存形式
JSON を推奨する。

例:
```json
{
  "WorkbookPath": "D:\\work\\sample.xlsx",
  "DefaultSheet": "Main",
  "LastUpdated": "2026-04-03T13:00:00"
}
```

### 11.2 保存場所
推奨:
- `%APPDATA%\Exshell\session.json`

### 11.3 読込ルール
- `eopen` 実行時に保存
- 参照系コマンドはセッションから既定対象を取得
- 明示指定があればそちら優先

---

## 12. コマンド別内部処理

### 12.1 eopen
処理:
1. 引数からファイルパス取得
2. `PathResolver` で正規化
3. `ExcelBridge.OpenOrAttach`
4. 既定シート設定
5. `SessionStore.Save`

### 12.2 els
処理:
1. セッション読込
2. シート決定
3. `ExcelBridge.ListTextBoxes`
4. 標準出力へ整形出力

### 12.3 ecat
処理:
1. セッション読込
2. `TextBoxRef` 解決
3. `ExcelBridge.ReadText`
4. stdout へ出力

### 12.4 cate
処理:
1. stdin 全読込
2. セッション読込
3. `TextBoxRef` 解決
4. `ExcelBridge.WriteText(append)`
5. 終了

### 12.5 ediff
処理:
1. 2 つの `TextBoxRef` 解決
2. `ReadText` で 2 テキスト取得
3. 一時ファイル作成
4. UTF-8 LF で書込
5. `DiffService` 実行
6. 結果出力
7. 一時ファイル削除

### 12.6 einfo
処理:
1. セッション読込
2. 情報整形
3. 出力

---

## 13. DiffService 設計

### 13.1 役割
- 将来差し替え可能な差分サービスを提供する

### 13.2 初期方針
内部で 2 通りの実装余地を持つ。

- 実装 A: 自前の単純 diff
- 実装 B: 外部 `diff` 呼び出し

初期版では WSL `diff` の呼び出しを許容する。  
ただし PowerShell 単独運用も考慮し、将来は純 Windows の差分戦略も差し替え可能にする。

### 13.3 推奨
`IDiffService` の背後で実装を切替可能にする。

---

## 14. 標準入力・標準出力設計

### 14.1 stdin
- `cate` が使用する
- 入力はすべて読み切る
- UTF-8 前提で扱う

### 14.2 stdout
- `ecat`, `els`, `ediff`, `einfo` が使用する
- パイプ利用前提のため余計な装飾は避ける

### 14.3 stderr
- エラー情報専用

---

## 15. WSL 両対応設計

### 15.1 前提
WSL からは Windows exe を直接実行可能である。

### 15.2 本体側の吸収範囲
Exshell 本体で以下を吸収する。

- WSL パス入力
- UTF-8 入出力
- 標準入出力
- 一時ファイルの Windows / WSL 共存

### 15.3 WSL ラッパー
WSL 側の `function` は薄いラッパーに留める。  
ビジネスロジックは持たせない。

---

## 16. PowerShell 補助運用設計

PowerShell 側では `uls`, `ucat`, `unl`, `udiff` などの関数を別途 profile に定義してよい。  
ただし Exshell 本体はこれらに依存しない。

---

## 17. 例外処理・終了コード

### 17.1 想定例外
- ファイル不在
- Excel 起動失敗
- ブック未検出
- シート未検出
- Shape 未検出
- COM 例外
- セッション破損
- diff 実行失敗

### 17.2 終了コード案
- 0: 正常
- 1: 引数エラー
- 2: セッション未確立
- 3: ブック未検出
- 4: シート未検出
- 5: テキストボックス未検出
- 6: Excel 操作失敗
- 7: 標準入力処理失敗
- 8: 一時ファイル処理失敗
- 9: 差分処理失敗

---

## 18. テスト設計方針

### 18.1 単体テスト対象
- TextBoxRef 解析
- パス変換
- セッション JSON 読書き
- 引数解析
- 改行正規化

### 18.2 結合テスト対象
- Excel ブックオープン
- Shape 列挙
- テキスト読取
- テキスト更新
- WSL パス入力
- 一時ファイル diff

### 18.3 手動確認項目
- PowerShell からの利用
- WSL からの利用
- パイプ入力
- 追記
- 日本語文字列
- 複数行テキスト

---

## 19. 実装上の注意

- COM オブジェクトの解放を徹底する
- Excel プロセスを余分に残さない
- Save / Close を呼ばないことをコード上でも明確にする
- stdout にデバッグ出力を混ぜない
- WSL / PowerShell で文字化けしないことを早期に確認する

---

## 20. まとめ

Exshell v0.3 の内部設計は、PowerShell / WSL 両対応を前提に、  
本体でパス・標準入出力・Excel ブリッジを吸収し、Unix テキスト処理は外部コマンドへ委ねる構成とする。
