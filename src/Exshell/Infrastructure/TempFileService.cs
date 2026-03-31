using System.Text;

namespace Exshell.Infrastructure;

/// <summary>
/// 一時ファイルの払い出し・UTF-8 LF 書き込み・後始末を担当する。
/// </summary>
public static class TempFileService
{
    /// <summary>
    /// テキストを UTF-8 BOM なし LF で一時ファイルへ書き出し、パスを返す。
    /// </summary>
    public static string WriteUtf8Lf(string text)
    {
        var path = Path.GetTempFileName();
        try
        {
            // 改行を LF に統一
            var lf = text.Replace("\r\n", "\n").Replace("\r", "\n");
            File.WriteAllText(path, lf, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return path;
        }
        catch (Exception ex)
        {
            TryDelete(path);
            throw new ExshellException(
                $"Failed to write temp file: {ex.Message}",
                ExitCodes.TempFileError
            );
        }
    }

    public static void TryDelete(string? path)
    {
        if (path == null) return;
        try { File.Delete(path); } catch { }
    }
}
