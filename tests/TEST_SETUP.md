# Exshell テスト環境設定ガイド

## 前提条件

### 1. WSL 環境
- WSL がインストールされ、有効化されていること
- `wsl --version` コマンドが正常に動作すること
- 基本的な Unix コマンド（`echo`, `diff`, `cat`, `ls` など）が利用可能

### 2. .NET 環境
- .NET 8.0 SDK がインストールされていること
- `dotnet test` コマンドが実行可能

### 3. ファイルシステム権限
- 一時ディレクトリ（`%TEMP%`）への読み書き権限
- `%APPDATA%` への読み書き権限

## テスト実行

### 全テスト実行
```powershell
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj
```

### 特定のテストクラス実行
```powershell
# PathConverter のテストのみ
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --filter "FullyQualifiedName~PathConverterTests"

# SessionStore のテストのみ
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --filter "FullyQualifiedName~SessionStoreTests"

# ProcessRunner のテストのみ（WSL 依存）
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --filter "FullyQualifiedName~ProcessRunnerTests"
```

### 詳細出力での実行
```powershell
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --verbosity normal
```

## WSL 関連テストについて

### 動作確認
ProcessRunner のテストは WSL 環境に依存します。事前に以下を確認：

```powershell
# WSL の状態確認
wsl --status

# WSL でのコマンド実行確認
wsl echo "test"
wsl diff --help
```

### トラブルシューティング

#### WSL が利用不可の場合
ProcessRunner のテストは自動的にスキップされます。ログに以下が表示されます：
```
WSL is not available - test skipped
```

#### WSL コマンド失敗の場合
1. WSL ディストリビューションが正常にインストールされているか確認
2. 必要なコマンド（`echo`, `diff`）がインストールされているか確認

## テスト環境の初期化

### 一時ファイル削除
テスト実行前に古い一時ファイルを削除：

```powershell
# 一時ディレクトリの Exshell 関連ファイル削除
Remove-Item -Path "$env:TEMP\ExshellTests*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:APPDATA\Exshell\session.json" -Force -ErrorAction SilentlyContinue
```

## CI/CD 環境での注意点

### GitHub Actions / Azure DevOps
- WSL を有効にする手順が必要
- Windows runners を使用する場合は WSL のセットアップが必要

### 例（GitHub Actions）
```yaml
- name: Enable WSL
  run: |
    wsl --install --no-distribution
    wsl --set-default-version 2
```

## テスト独立性

各テストは以下を保証：
- 他のテストに影響しない（一時ファイル使用）
- 実行順序に依存しない
- 環境変数を適切にクリーンアップ

## パフォーマンス

### 高速実行のために
```powershell
# 並列実行（デフォルト）
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj

# シーケンシャル実行（トラブル時）
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --parallel 1
```

### テスト時間の目安
- PathConverter テスト: < 1秒
- SessionStore テスト: < 2秒（ファイル I/O のため）
- TempFileService テスト: < 2秒
- ProcessRunner テスト: 2-5秒（WSL 起動のため）

## トラブルシューティング

### よくある問題

#### "WSL is not available"
```powershell
# WSL の状態確認
wsl --status
wsl --list --verbose

# WSL が未インストールの場合
wsl --install
```

#### "Access denied" エラー
- PowerShell を管理者権限で実行
- ウイルス対策ソフトの除外設定を確認

#### テスト失敗
```powershell
# 詳細ログ出力
dotnet test tests/Exshell.Tests/Exshell.Tests.csproj --verbosity diagnostic
```