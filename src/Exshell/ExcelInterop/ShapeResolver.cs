using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.ExcelInterop;

/// <summary>
/// Shape 名から対象 Shape を取得し、テキストを持つ Shape か判定する。
/// </summary>
public static class ShapeResolver
{
    /// <summary>
    /// ワークシート上のテキストを持つ Shape の名前一覧を返す。
    /// TextFrame2 にアクセスできる Shape を「テキスト保持 Shape」とみなす。
    /// </summary>
    public static IReadOnlyList<string> ListTextShapes(XL.Worksheet ws)
    {
        var result = new List<string>();
        foreach (XL.Shape shape in ws.Shapes)
        {
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

    /// <summary>
    /// Shape 名に一致する Shape を返す。見つからなければ例外。
    /// </summary>
    public static XL.Shape FindShape(XL.Worksheet ws, string shapeName)
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
}
