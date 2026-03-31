# Exshell 内部設計案 v0.1

## 1. 文書の目的

本書は、Exshell 外部仕様書 v0.1 に基づき、Exshell の初期実装に必要な内部構造、責務分割、主要クラス、データ構造、処理フロー、および実装上の注意事項を定義するものである。

Exshell は、Excel をテキスト入出力フロントエンドとして扱い、PowerShell + WSL 上の CLI 文化と接続するためのブリッジツールである。  
本バージョンでは、Excel デスクトップ版がインストール済みの Windows 環境を前提とし、C# + .NET Framework 4.8 + Excel COM Interop により実装する。

---

## 2. システム概要

### 2.1 役割
Exshell は以下の責務を持つ。

- Excel ブックを操作対象として確立する
- Excel ワークシート上の Shape テキストボックスを識別する
- テキストボックス内容を標準出力へ送る
- 標準入力またはファイル内容をテキストボックスへ反映する
- テキストボックス内容を一時ファイルへ出力し、WSL の `diff` 等の外部コマンドと接続する
- セッション情報を永続化し、継続操作を可能にする

### 2.2 非責務
Exshell は以下を責務としない。

- Workbook の Save 実行
- Workbook の Close 実行
- Excel.Application の Quit 実行
- Excel 文書全体の編集管理
- 差分アルゴリズムそのものの内製
- PowerShell / WSL 環境設定の代行

---

## 3. 前提技術

### 3.1 実行環境
- Windows 11
- Excel デスクトップ版インストール済み
- WSL 利用可能
- PowerShell Terminal 利用

### 3.2 開発技術
- Visual Studio 2022
- C#
- .NET Framework 4.8
- `Microsoft.Office.Interop.Excel`
- 必要に応じて `Microsoft.Office.Core`

### 3.3 外部連携
- `wsl diff`
- 将来的に `wsl nl`, `wsl grep`, `wsl sort` 等との連携拡張を想定する

---

## 4. アーキテクチャ方針

### 4.1 基本方針
Exshell は、以下の4層で構成する。

1. **CLI 層**  
   引数解析、コマンド振り分け、終了コード制御、標準入出力制御を担当する。

2. **アプリケーション層**  
   コマンド単位のユースケース処理を担当する。

3. **Excel 連携層**  
   Excel COM への接続、Workbook / Worksheet / Shape 操作を担当する。

4. **インフラ補助層**  
   セッション永続化、一時ファイル、外部プロセス起動、ログ出力を担当する。

---

## 5. 論理構成

### 5.1 推奨名前空間構成

```text
Exshell
 ├─ Program.cs
 ├─ Cli
 │   ├─ CommandDispatcher.cs
 │   ├─ CommandContext.cs
 │   ├─ CommandResult.cs
 │   └─ ArgumentParser.cs
 ├─ Application
 │   ├─ Commands
 │   │   ├─ EopenCommand.cs
 │   │   ├─ EcatCommand.cs
 │   │   ├─ CateCommand.cs
 │   │   ├─ EdiffCommand.cs
 │   │   ├─ ElsCommand.cs
 │   │   └─ EinfoCommand.cs
 │   ├─ Models
 │   │   ├─ SessionInfo.cs
 │   │   ├─ ShapeReference.cs
 │   │   └─ ExcelOpenResult.cs
 │   └─ Services
 │       ├─ SessionService.cs
 │       ├─ ExcelTextService.cs
 │       ├─ DiffService.cs
 │       └─ TempFileService.cs
 ├─ ExcelInterop
 │   ├─ ExcelAppGateway.cs
 │   ├─ WorkbookResolver.cs
 │   ├─ WorksheetResolver.cs
 │   ├─ ShapeResolver.cs
 │   └─ ShapeTextAccessor.cs
 └─ Infrastructure
     ├─ JsonSessionStore.cs
     ├─ ProcessRunner.cs
     ├─ PathConverter.cs
     └─ ConsoleWriter.cs
```

---

## 6. 主要データモデル

### 6.1 SessionInfo
現在のセッション状態を保持する。

```csharp
public class SessionInfo
{
    public string WorkbookPath { get; set; }
    public string DefaultSheetName { get; set; }
    public DateTime UpdatedAt { get; set; }
}
```

### 6.2 ShapeReference
コマンド引数から解決されたテキストボックス参照を表す。

```csharp
public class ShapeReference
{
    public string SheetName { get; set; }
    public string ShapeName { get; set; }
}
```

### 6.3 CommandResult
コマンド実行結果を表す。

```csharp
public class CommandResult
{
    public int ExitCode { get; set; }
    public string StandardOutput { get; set; }
    public string StandardError { get; set; }
}
```

---

## 7. クラス責務

### 7.1 CLI 層

#### Program
- エントリーポイント
- 例外の最終捕捉
- 終了コード返却

#### ArgumentParser
- コマンド名抽出
- オプション抽出
- 引数の基本構文チェック

#### CommandDispatcher
- コマンド名に応じてアプリケーション層の処理へ振り分ける

#### CommandContext
- 標準入力
- 標準出力
- 標準エラー
- カレントディレクトリ
- 実行オプション
などの実行コンテキストを保持する

### 7.2 アプリケーション層

#### EopenCommand
- 指定 Excel ファイルを開く、または既存の開いている対象を取得する
- セッションを書き込む

#### EcatCommand
- セッションから対象ブックを解決する
- ShapeReference を解析する
- テキストボックス内容を取得し、標準出力へ出す

#### CateCommand
- 標準入力を読み込む
- `--append` 指定有無に応じて上書き／追記を行う
- Save しない

#### EdiffCommand
- 2つのテキストボックス内容を取得する
- 一時ファイルへ UTF-8 LF で書き出す
- `wsl diff` を実行する
- 結果を標準出力へ返す
- 一時ファイルを削除する

#### ElsCommand
- 対象シート内の利用可能 Shape 一覧を出力する

#### EinfoCommand
- 現在の SessionInfo を表示する

### 7.3 Excel 連携層

#### ExcelAppGateway
- Excel.Application の取得または生成
- 表示制御
- COM オブジェクトの安全な解放補助

#### WorkbookResolver
- 指定パスに対応する Workbook を開く
- 既に開いている Workbook がある場合はそれを返す

#### WorksheetResolver
- Sheet 名から Worksheet を取得する
- 既定シート解決を行う

#### ShapeResolver
- Shape 名から対象 Shape を取得する
- テキストを持つ Shape か判定する

#### ShapeTextAccessor
- Shape から文字列を取得する
- Shape に文字列を設定する
- `TextFrame2` と `TextFrame` の差を吸収する

### 7.4 インフラ補助層

#### JsonSessionStore
- SessionInfo の JSON 永続化
- 読み込み／書き込み／削除

#### ProcessRunner
- `wsl` プロセス起動
- 引数エスケープ
- 標準出力・標準エラー取得
- 終了コード取得

#### PathConverter
- Windows パスを WSL パスへ変換する
- 例: `C:\Temp\a.txt` → `/mnt/c/Temp/a.txt`

#### TempFileService
- temp ファイル名の払い出し
- UTF-8 LF での書き込み
- 後始末

#### ConsoleWriter
- 標準出力／標準エラーへの書き分け
- 将来的な色付き表示対応の拡張点

---

## 8. コマンド別処理フロー

### 8.1 eopen
1. 引数解析
2. ファイル存在確認
3. Excel.Application 取得
4. 対象 Workbook 解決
5. 必要なら Workbook を Open
6. SessionInfo 更新
7. 実行結果出力

### 8.2 ecat
1. SessionInfo 読込
2. ShapeReference 解決
3. Workbook / Worksheet / Shape 解決
4. Shape テキスト取得
5. 標準出力へ書き込み

### 8.3 cate
1. SessionInfo 読込
2. ShapeReference 解決
3. 標準入力を全件読込
4. Shape 現在値取得
5. `--append` 指定時は連結
6. Shape に書込
7. 終了

### 8.4 ediff
1. 2つの ShapeReference 解決
2. それぞれのテキスト取得
3. temp1, temp2 を生成
4. UTF-8 LF で書込
5. Windows パスを WSL パスへ変換
6. `wsl diff temp1 temp2` 実行
7. 標準出力へ返却
8. temp ファイル削除

### 8.5 els
1. SessionInfo 読込
2. 対象 Worksheet 解決
3. Shapes 走査
4. テキスト取得可能 Shape のみ抽出
5. 一覧出力

---

## 9. Shape 取扱方針

### 9.1 対象
初期実装では以下を対象とする。

- ワークシート上の Shape
- 文字列取得／設定が可能なテキストボックス系 Shape

### 9.2 非対象
初期実装では以下を対象外とする。

- ActiveX TextBox
- フォームコントロール
- グループ化図形の内部個別要素
- SmartArt 等の特殊オブジェクト

### 9.3 テキスト取得優先順
1. `TextFrame2.TextRange.Text`
2. `TextFrame.Characters().Text`

---

## 10. セッション設計

### 10.1 保存先候補
初期実装では以下のいずれかを採用する。

- `%APPDATA%\Exshell\session.json`
- カレントディレクトリ配下 `.exshell-session.json`

### 10.2 推奨
ユーザー単位利用を考慮し、以下を推奨する。

```text
%APPDATA%\Exshell\session.json
```

### 10.3 セッション例
```json
{
  "WorkbookPath": "D:\\work\\sample.xlsx",
  "DefaultSheetName": "Main",
  "UpdatedAt": "2026-03-31T18:30:00"
}
```

---

## 11. WSL 連携設計

### 11.1 基本方針
- Exshell は WSL コマンドの結果を利用する
- 差分計算そのものは WSL `diff` に委譲する
- PowerShell profile 上では `uls` 系補助関数を別途定義する

### 11.2 `uls` 系は Exshell の内部責務ではない
以下は運用環境設定であり、Exshell 本体の内部実装責務ではない。

```powershell
function uls   { wsl ls @Args }
function ucat  { wsl cat @Args }
function unl   { wsl nl @Args }
function udiff { wsl diff @Args }
```

---

## 12. 例外・エラー処理設計

### 12.1 代表的エラー
- 引数不正
- セッション未確立
- Workbook 不存在
- Worksheet 不存在
- Shape 不存在
- Shape が文字列を持たない
- Excel COM 例外
- WSL 実行失敗
- temp ファイル操作失敗

### 12.2 終了コード案
- `0`: 成功
- `1`: 引数エラー
- `2`: セッション未確立
- `3`: Workbook 関連エラー
- `4`: Worksheet 関連エラー
- `5`: Shape 関連エラー
- `6`: Excel COM 例外
- `7`: WSL 実行エラー
- `8`: temp / ファイルエラー
- `9`: その他予期しないエラー

---

## 13. COM 運用上の注意

### 13.1 解放
Excel COM オブジェクトは極力短命に扱い、不要オブジェクトを明示的に解放する。

### 13.2 推奨
- Worksheet / Shapes / Shape を長く保持しない
- `Marshal.ReleaseComObject` を必要箇所で使用する
- `GC.Collect` 依存の設計にしない

### 13.3 注意
- COM 参照残留により Excel プロセスが残る可能性がある
- ただし本ツールは Excel 自体を終了させない前提なので、過剰な終了制御は行わない

---

## 14. ログ方針

初期版では詳細ログは必須としない。  
ただし、障害調査用として以下の最小限のフックを残す。

- デバッグビルド時の内部例外出力
- `--verbose` オプション将来拡張余地
- ProcessRunner のコマンド実行内容記録余地

---

## 15. テスト方針

### 15.1 単体テスト対象
- ShapeReference 解析
- SessionInfo 読込／保存
- Windows → WSL パス変換
- temp ファイル UTF-8 LF 書込
- 引数解析

### 15.2 手動結合テスト対象
- Excel 起動中／未起動時の eopen
- 既存 Workbook 再利用
- テキストボックス読込・書込
- `--append` 動作
- `ediff` と `wsl diff` 連携
- 日本語文字列 diff

### 15.3 テスト用 Excel
専用サンプルブックを準備し、少なくとも以下を持たせる。

- `Main:txtInput`
- `Main:txtOutput`
- `Main:left`
- `Main:right`

---

## 16. 今後の拡張候補

- `eimport`
- `eexport`
- `eclear`
- `eselect`
- `epipe`
- 複数ブック切替
- セル操作対応
- 色付き差分表示
- `wsl nl`, `wsl grep` 等との専用連携コマンド

---

## 17. 初期実装優先順位

### Phase 1
- SessionInfo
- eopen
- einfo
- ShapeReference
- ecat
- cate

### Phase 2
- els
- temp ファイル
- PathConverter
- ProcessRunner
- ediff

### Phase 3
- 例外整理
- 終了コード統一
- テストサンプル整備
- 運用ドキュメント整備

---

## 18. まとめ

Exshell の内部設計は、Excel COM 操作、CLI 入出力、WSL 連携を明確に分離し、「Excel テキストボックスを CLI で扱うための橋渡し」に責務を絞ることを基本方針とする。

初期版では、以下を特に重視する。

- Shape テキストボックスへの対象限定
- Save / Close 非実行
- セッションの明確化
- `ediff` における temp ファイル + `wsl diff` 連携
- PowerShell / WSL / Excel の役割分離

これにより、Exshell は過度に肥大化せず、UNIX 的な透明感を持つ Excel テキスト処理ツールとして成立する。
