# Exshell 内部設計案 v0.2

## 1. 文書概要

本書は、Exshell 外部仕様書 v0.2 に対応する内部設計案を示す。
Exshell は Excel デスクトップ版がインストールされた Windows 環境上で動作する CLI ツール群であり、主利用環境は WSL ターミナルとする。

Exshell 自体は Windows 実行ファイルとして実装し、WSL から呼び出されることを前提とする。
UNIX 系の一般的なテキスト処理は WSL 側の標準コマンドを利用し、Exshell は Excel と標準入出力の橋渡しに責務を限定する。

---

## 2. 設計方針

### 2.1 基本方針
- 実装言語は C# とする。
- 初期ターゲットは .NET Framework 4.8 を想定する。
- Excel 操作は COM Interop を使用する。
- 実行ファイルは Windows 側に配置し、WSL から `*.exe` として呼び出す。
- 主運用は WSL ターミナルで行う。
- `ls`, `grep`, `diff`, `nl`, `sed`, `awk` などの汎用処理は WSL 標準コマンドに委譲する。

### 2.2 責務分離
Exshell が担当する範囲は以下とする。
- Excel ブックのオープン
- セッション管理
- シート / Shape 解決
- テキストボックスの読み取り
- テキストボックスへの書き込み
- Excel テキストを stdout に流す処理
- stdin から Excel テキストボックスへ反映する処理
- 簡便差分コマンド `ediff` の提供

Exshell が担当しない範囲は以下とする。
- ブック保存
- ブッククローズ
- Excel アプリケーション終了
- 汎用 diff / grep / sort / uniq / sed / awk ロジック実装

### 2.3 v0.2 の重要変更点
v0.1 からの主な変更は以下。
- 主利用環境を PowerShell から WSL ターミナルへ変更
- WSL パス入力を正式対応対象に追加
- `ediff` の位置づけを「簡便コマンド」として整理
- PowerShell 補助コマンド群は参考運用へ移動

---

## 3. 想定アーキテクチャ

### 3.1 全体構成

```text
[WSL Terminal]
   ├─ ls / grep / diff / nl / sed / awk / sort / uniq
   ├─ pipe / redirect / process substitution
   └─ eopen.exe / ecat.exe / cate.exe / els.exe / ediff.exe
                │
                ▼
         [Windows CLI Process]
                │
                ▼
        [Excel COM Interop Layer]
                │
                ▼
           [Excel Workbook]
                │
                ▼
        [Worksheet / Shapes / Text]
```

### 3.2 実行モデル
- Exshell 各コマンドは独立プロセスとして起動する。
- 共有状態はセッションファイル経由で保持する。
- Excel インスタンスや Workbook 参照は毎回取得し直す。
- 永続 COM オブジェクト保持は行わない。

この方式により、WSL からの単発呼び出しと相性を良くし、プロセス寿命と COM オブジェクト寿命を分離する。

---

## 4. 論理レイヤ構成

### 4.1 層構造

#### 4.1.1 CLI 層
責務:
- コマンドライン引数解析
- オプション妥当性検証
- 標準入力読取
- 標準出力 / 標準エラー出力
- 終了コード制御

対象クラス例:
- `Program`
- `CommandLineParser`
- `CommandContext`
- `CommandExecutor`

#### 4.1.2 アプリケーション層
責務:
- `eopen`, `ecat`, `cate`, `els`, `ediff`, `einfo` などのユースケース実行
- セッション読込/保存
- 識別子解決
- Excel サービス呼び出し順制御

対象クラス例:
- `OpenWorkbookUseCase`
- `ReadTextboxUseCase`
- `WriteTextboxUseCase`
- `ListTextboxUseCase`
- `DiffTextboxUseCase`
- `SessionInfoUseCase`

#### 4.1.3 ドメイン/モデル層
責務:
- セッション情報保持
- 対象ブック、シート、Shape 識別情報保持
- 書込モード等の値オブジェクト化

対象クラス例:
- `ExshellSession`
- `WorkbookTarget`
- `SheetShapeReference`
- `WriteMode`
- `PathKind`

#### 4.1.4 インフラ層
責務:
- Excel COM 操作
- セッションファイル永続化
- 一時ファイル生成
- 外部プロセス呼出し
- パス変換

対象クラス例:
- `ExcelInteropGateway`
- `SessionStore`
- `TempFileService`
- `ProcessRunner`
- `PathNormalizer`
- `WslPathConverter`

---

## 5. コマンド別内部設計

## 5.1 eopen

### 5.1.1 役割
- 指定された Excel ファイルを開く、または既に開かれている同一ブックを取得する。
- セッション情報を更新する。

### 5.1.2 入力
- ファイルパス
- 既定シート名（任意）

### 5.1.3 処理概要
1. 入力パスを Windows パスへ正規化
2. ファイル存在確認
3. Excel.Application 取得または起動
4. 対象 Workbook を開く、または既存オープンブックから取得
5. 既定シートの妥当性確認
6. セッション保存
7. 結果表示

### 5.1.4 注意点
- WSL パス `/mnt/c/...` を Windows パスへ変換できること
- UNC などは初期版では必要に応じて制限してよい
- Save / Close / Quit は呼ばない

---

## 5.2 ecat

### 5.2.1 役割
- 指定テキストボックス内容を stdout に出力する

### 5.2.2 処理概要
1. セッション読込
2. 対象 `Sheet:ShapeName` 解決
3. Excel から文字列取得
4. stdout へ書込

### 5.2.3 注意点
- WSL パイプ利用を前提とし、stdout は余計な装飾を付けない
- 文字コードは UTF-8 を基本とする
- エラーメッセージは stderr に分離する

---

## 5.3 cate

### 5.3.1 役割
- stdin の内容を指定テキストボックスへ反映する

### 5.3.2 書込モード
- 上書き
- 追記

### 5.3.3 処理概要
1. セッション読込
2. stdin 全文読取
3. 対象 `Sheet:ShapeName` 解決
4. 既存文字列取得
5. モードに応じて新文字列生成
6. Excel テキストボックスへ反映
7. 完了表示（必要最小限）

### 5.3.4 注意点
- 保存は行わない
- 改行コードは内部で正規化方針を持つ
- 追記時に必要なら区切り改行の扱いを明確化する

---

## 5.4 els

### 5.4.1 役割
- 対象シート上の利用可能テキストボックス一覧を表示する

### 5.4.2 処理概要
1. セッション読込
2. 対象シート決定
3. Worksheet.Shapes を走査
4. 文字列取得対象 Shape のみ抽出
5. 一覧を stdout へ出力

### 5.4.3 注意点
- 文字列取得可能判定を共通化する
- 順序は Excel の Shapes 順とし、必要に応じて名前順オプションを検討

---

## 5.5 ediff

### 5.5.1 役割
- 2つのテキストボックスの差分表示を簡便に行う
- 内部では `diff` 相当の外部コマンド利用を前提とする

### 5.5.2 処理概要
1. セッション読込
2. 左右の `Sheet:ShapeName` 解決
3. Excel から左右文字列取得
4. 一時ファイル temp1, temp2 作成
5. UTF-8 / LF で内容書込
6. `wsl diff temp1 temp2` 相当を実行
7. diff 結果を stdout に返す
8. 一時ファイル削除

### 5.5.3 注意点
- v0.2 では簡便コマンドであり、上級利用者は `diff <(ecat ...) <(ecat ...)` を利用可能
- `wsl` 呼出しの失敗時メッセージを明確にする
- 一時ファイル削除は finally で保証する

---

## 5.6 einfo

### 5.6.1 役割
- セッション情報を表示する

### 5.6.2 出力対象例
- WorkbookPath
- DefaultSheet
- SessionFilePath

---

## 6. パス処理設計

## 6.1 パス種別
Exshell が受けるパスは次の3種を想定する。
- Windows パス: `D:\work\sample.xlsx`
- WSL パス: `/mnt/d/work/sample.xlsx`
- 相対パス

## 6.2 PathNormalizer
責務:
- 入力文字列を受け取る
- パス種別判定を行う
- 実際に Excel で扱う Windows 絶対パスへ正規化する

### 6.2.1 想定メソッド
- `NormalizeToWindowsPath(string input, string currentDirectoryContext)`
- `TryConvertWslPathToWindowsPath(string wslPath)`
- `IsWindowsPath(string input)`
- `IsWslPath(string input)`

## 6.3 相対パスの扱い
相対パスは、起動時カレントディレクトリ基準で解決する。
ただし WSL から起動された場合、Windows 側プロセスから見たカレントディレクトリとの差異に注意が必要である。
初期版では以下のどちらかを採用する。
- WSL からは絶対パス利用を推奨する
- もしくは起動時ディレクトリ引継ぎ方法を設計する

v0.2 時点では実装容易性を優先し、絶対パス利用を推奨とする。

---

## 7. セッション管理設計

## 7.1 セッション保持項目
- `WorkbookPath`
- `DefaultSheet`
- `UpdatedAt`
- `Version`

## 7.2 セッション保存先
初期案:
- `%APPDATA%\Exshell\session.json`

## 7.3 SessionStore の責務
- セッション読込
- セッション保存
- セッション存在確認
- JSON バージョン管理

### 7.3.1 注意点
- 壊れた JSON への対処
- 旧バージョンとの互換性
- 将来的な複数セッション対応余地確保

---

## 8. Excel Interop 設計

## 8.1 ExcelInteropGateway の責務
- Excel.Application 取得
- Workbook 取得/オープン
- Worksheet 取得
- Shape 取得
- Shape テキスト読取/書込

## 8.2 主要操作
### 8.2.1 Application 取得
- 既存 Excel へ接続を試みる
- 取得できない場合は新規起動

### 8.2.2 Workbook 取得
- 開いているブック一覧からフルパス一致を探す
- 見つからなければ Open

### 8.2.3 Worksheet 取得
- 名前一致で取得
- 未指定時は既定シート

### 8.2.4 Shape 取得
- `Worksheet.Shapes` を走査し、名前一致で取得

### 8.2.5 テキスト取得/設定
- `TextFrame2` 優先
- 必要に応じて `TextFrame` へフォールバック
- 文字列取得不能 Shape は対象外

## 8.3 COM 解放方針
- 取得した COM オブジェクトは finally で解放する
- `Marshal.ReleaseComObject` を適切に使用する
- 参照の二重解放を避ける
- コレクション反復時の一時 COM 参照にも注意する

---

## 9. 標準入出力・文字コード設計

## 9.1 stdout
- 本文出力のみ
- 装飾なし
- WSL パイプ処理前提

## 9.2 stderr
- エラー内容を簡潔に出力
- トラブルシューティング可能な情報を含める

## 9.3 stdin
- `cate` は stdin 全文を読み込む
- 初期版ではストリーミング書込ではなく一括読込とする

## 9.4 文字コード
- 標準入出力は UTF-8 前提を基本とする
- 一時ファイルも UTF-8 で作成する

## 9.5 改行コード
- Excel から取得したテキストは内部で LF 正規化可能とする
- 一時ファイルは LF で出力する
- Excel 書込時の改行は表示確認に基づき方針調整可能とする

---

## 10. 外部プロセス実行設計

## 10.1 ProcessRunner の責務
- 外部コマンド起動
- stdout/stderr 取得
- 終了コード取得

## 10.2 ediff での利用
呼出し例の概念:
- `wsl diff <temp1> <temp2>`

## 10.3 注意点
- temp パスの WSL 参照可能性確認
- スペースを含むパスのクォート処理
- `wsl` が利用できない環境へのエラーメッセージ

---

## 11. 主要クラス案

### 11.1 CLI 関連
- `Program`
- `CommandLineParser`
- `ParsedArguments`
- `ExitCode`

### 11.2 UseCase 関連
- `OpenWorkbookUseCase`
- `ReadTextboxUseCase`
- `WriteTextboxUseCase`
- `ListTextboxUseCase`
- `DiffTextboxUseCase`
- `SessionInfoUseCase`

### 11.3 Model 関連
- `ExshellSession`
- `SheetShapeReference`
- `WorkbookPathInfo`
- `WriteMode`

### 11.4 Infra 関連
- `ExcelInteropGateway`
- `SessionStore`
- `PathNormalizer`
- `TempFileService`
- `ProcessRunner`
- `TextNormalizationService`

---

## 12. 例外・エラー設計

## 12.1 想定エラー区分
- 引数不正
- セッション未存在
- ファイル未存在
- シート未存在
- Shape 未存在
- Shape がテキスト対象外
- Excel 起動失敗
- WSL / diff 呼出し失敗
- COM 操作失敗

## 12.2 終了コード案
- 0: 成功
- 1: 引数エラー
- 2: セッションエラー
- 3: ファイル/ブックエラー
- 4: シートエラー
- 5: Shape エラー
- 6: Excel/COM エラー
- 7: 外部コマンド実行エラー
- 8: 入出力エラー

---

## 13. ログ方針

v0.2 では常時ログファイル出力は必須としない。
ただしデバッグ容易性のため、将来以下を検討可能とする。
- `--verbose`
- `--debug`
- 実行内容ログの任意出力

---

## 14. テスト観点

## 14.1 単体テスト対象
- パス正規化
- 識別子解析
- セッション JSON 読書き
- 書込モード分岐
- 改行正規化
- 外部プロセス引数構築

## 14.2 結合テスト対象
- 実 Excel ブックへの eopen
- ecat / cate / els
- WSL 経由 ediff
- WSL パス入力での eopen

## 14.3 手動確認対象
- Excel テキストボックス種別差異
- 日本語入出力
- 長文テキスト
- 未保存状態の挙動

---

## 15. MVP 実装順序

1. 共通基盤
   - 引数解析
   - セッション管理
   - パス正規化
2. Excel 読取
   - eopen
   - einfo
   - els
   - ecat
3. Excel 書込
   - cate 上書き
   - cate 追記
4. 差分補助
   - ediff
5. 仕上げ
   - エラーコード整理
   - テスト拡充
   - WSL 運用確認

---

## 16. 今後の拡張候補
- `eclear`
- `eappend`
- `eimport`
- `eexport`
- 複数セッション
- 名前付きプロファイル
- 対象 Shape 自動候補表示
- process substitution 前提の運用ガイド整備

---

## 17. まとめ

v0.2 の内部設計では、Exshell を「Windows 上で動作し、WSL から利用される Excel ブリッジ」として定義した。
この方針により、Exshell 側は Excel / COM / セッション処理に集中でき、汎用テキスト処理は WSL 標準コマンドへ委譲できる。

設計の要点は以下である。
- 主戦場は WSL ターミナル
- Exshell は Windows 実行ファイル
- COM 参照は短命に保つ
- セッションファイルで状態共有する
- path / encoding / newline を明示的に扱う
