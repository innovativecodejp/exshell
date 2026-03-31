using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.ExcelInterop;

/// <summary>
/// Sheet 名から Worksheet を取得し、既定シート解決を行う。
/// </summary>
public static class WorksheetResolver
{
    /// <summary>
    /// シート名を指定してワークシートを取得する。
    /// null の場合はアクティブシートを返す。
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
}
