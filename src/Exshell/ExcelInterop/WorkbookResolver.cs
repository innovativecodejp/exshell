using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.ExcelInterop;

/// <summary>
/// 指定パスに対応する Workbook を開く、または既存のものを返す。
/// </summary>
public static class WorkbookResolver
{
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
}
