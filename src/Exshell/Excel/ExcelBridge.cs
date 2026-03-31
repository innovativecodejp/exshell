using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.Excel;

/// <summary>
/// Excel COM Interop ラッパー。
/// </summary>
public static class ExcelBridge
{
    // ---------------------------------------------------------------
    // COM P/Invoke helpers (Marshal.GetActiveObject は .NET 8 で非対応)
    // ---------------------------------------------------------------

    [DllImport("ole32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    private const int S_OK = 0;

    // ---------------------------------------------------------------
    // Application
    // ---------------------------------------------------------------

    /// <summary>
    /// 既に起動中の Excel を取得するか、なければ新規起動する。
    /// </summary>
    public static XL.Application GetOrCreateApplication()
    {
        var running = TryGetRunningExcel();
        if (running != null)
            return running;

        var app = new XL.Application();
        app.Visible = true;
        return app;
    }

    /// <summary>
    /// 実行中の Excel.Application を取得する。未起動なら null を返す。
    /// </summary>
    public static XL.Application? TryGetRunningExcel()
    {
        if (CLSIDFromProgID("Excel.Application", out var clsid) != S_OK)
            return null;
        if (GetActiveObject(ref clsid, IntPtr.Zero, out var obj) != S_OK)
            return null;
        return obj as XL.Application;
    }

    // ---------------------------------------------------------------
    // Workbook
    // ---------------------------------------------------------------

    /// <summary>
    /// 指定パスのブックが既に開かれていればそれを返し、なければ開く。
    /// </summary>
    public static XL.Workbook OpenOrGetWorkbook(XL.Application app, string fullPath)
    {
        fullPath = Path.GetFullPath(fullPath);

        foreach (XL.Workbook wb in app.Workbooks)
        {
            try
            {
                if (string.Equals(wb.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                    return wb;
            }
            catch { /* 列挙中に閉じられた場合は無視 */ }
        }

        if (!File.Exists(fullPath))
            throw new ExshellException(
                $"Excel file not found: {fullPath}",
                ExitCodes.WorkbookNotFound
            );

        return app.Workbooks.Open(fullPath);
    }

    // ---------------------------------------------------------------
    // Worksheet
    // ---------------------------------------------------------------

    /// <summary>
    /// シート名を指定してワークシートを取得する。null の場合はアクティブシートを返す。
    /// </summary>
    public static XL.Worksheet GetWorksheet(XL.Workbook wb, string? sheetName)
    {
        if (sheetName == null)
            return (XL.Worksheet)wb.ActiveSheet;

        foreach (XL.Worksheet ws in wb.Worksheets)
        {
            if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                return ws;
        }

        throw new ExshellException(
            $"Sheet not found: {sheetName}",
            ExitCodes.WorksheetNotFound
        );
    }

    // ---------------------------------------------------------------
    // Shape listing
    // ---------------------------------------------------------------

    /// <summary>
    /// ワークシート上のテキストを持つ Shape の名前一覧を返す。
    /// </summary>
    public static IReadOnlyList<string> ListTextShapes(XL.Worksheet ws)
    {
        var result = new List<string>();
        foreach (XL.Shape shape in ws.Shapes)
        {
            // TextFrame2 にアクセスできれば「テキストを持つ Shape」とみなす
            try
            {
                var tf2 = shape.TextFrame2;
                if (tf2 != null)
                    result.Add(shape.Name);
            }
            catch { /* テキストフレームを持たない Shape は無視 */ }
        }
        return result;
    }

    // ---------------------------------------------------------------
    // Text I/O
    // ---------------------------------------------------------------

    /// <summary>
    /// Shape のテキストを読み取り、改行コードを LF に正規化して返す。
    /// </summary>
    public static string GetShapeText(XL.Worksheet ws, string shapeName)
    {
        var shape = FindShape(ws, shapeName);
        var raw   = ReadRaw(shape);
        // Excel 内部改行 (\r) および \r\n を \n に統一
        return raw.Replace("\r\n", "\n").Replace("\r", "\n");
    }

    /// <summary>
    /// stdin テキストを Shape へ書き込む（上書きまたは追記）。
    /// </summary>
    public static void SetShapeText(XL.Worksheet ws, string shapeName, string text, bool append)
    {
        var shape = FindShape(ws, shapeName);

        if (append)
        {
            var current = ReadRaw(shape);
            // 既存テキストに追記。末尾が改行でなければ \r を挿入
            var suffix = text.Replace("\r\n", "\r").Replace("\n", "\r");
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
    // Private helpers
    // ---------------------------------------------------------------

    private static XL.Shape FindShape(XL.Worksheet ws, string shapeName)
    {
        foreach (XL.Shape shape in ws.Shapes)
        {
            try
            {
                if (string.Equals(shape.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                    return shape;
            }
            catch { }
        }
        throw new ExshellException(
            $"Shape not found: {shapeName}",
            ExitCodes.ShapeNotFound
        );
    }

    private static string ReadRaw(XL.Shape shape)
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
