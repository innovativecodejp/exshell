using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.ExcelInterop;

/// <summary>
/// Shape からのテキスト取得・設定を担当する。
/// TextFrame2 と TextFrame の差を吸収し、改行コードを正規化する。
/// </summary>
public static class ShapeTextAccessor
{
    /// <summary>
    /// Shape のテキストを読み取り、改行コードを LF に正規化して返す。
    /// </summary>
    public static string GetShapeText(XL.Worksheet ws, string shapeName)
    {
        var shape = ShapeResolver.FindShape(ws, shapeName);
        var raw   = ReadRaw(shape);
        // Excel 内部改行 (\r) および \r\n を \n に統一
        return raw.Replace("\r\n", "\n").Replace("\r", "\n");
    }

    /// <summary>
    /// テキストを Shape へ書き込む（上書きまたは追記）。
    /// </summary>
    public static void SetShapeText(XL.Worksheet ws, string shapeName, string text, bool append)
    {
        var shape = ShapeResolver.FindShape(ws, shapeName);

        if (append)
        {
            var current = ReadRaw(shape);
            var suffix  = text.Replace("\r\n", "\r").Replace("\n", "\r");
            // 既存末尾に改行がなければ \r を挿入してから追記
            text = (current.Length > 0 && current[^1] != '\r')
                ? current + "\r" + suffix
                : current + suffix;
        }
        else
        {
            // Excel Shape 内部は \r を改行として使用
            text = text.Replace("\r\n", "\r").Replace("\n", "\r");
        }

        WriteRaw(shape, text);
    }

    // ---------------------------------------------------------------
    // 内部ヘルパー
    // ---------------------------------------------------------------

    internal static string ReadRaw(XL.Shape shape)
    {
        try
        {
            return shape.TextFrame2.TextRange.Text ?? string.Empty;
        }
        catch { }

        try
        {
            return shape.TextFrame.Characters().Text ?? string.Empty;
        }
        catch
        {
            throw new ExshellException(
                $"Cannot read text from shape '{shape.Name}'",
                ExitCodes.ExcelOperationFailed
            );
        }
    }

    private static void WriteRaw(XL.Shape shape, string text)
    {
        try
        {
            shape.TextFrame2.TextRange.Text = text;
            return;
        }
        catch { }

        try
        {
            shape.TextFrame.Characters().Text = text;
            return;
        }
        catch
        {
            throw new ExshellException(
                $"Cannot write text to shape '{shape.Name}'",
                ExitCodes.ExcelOperationFailed
            );
        }
    }
}
