using Exshell.Excel;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// einfo
/// 現在セッション情報を表示する。
/// </summary>
public static class EinfoCommand
{
    public static int Run(string[] args)
    {
        try
        {
            var session = SessionStore.Load();

            if (session == null || string.IsNullOrEmpty(session.WorkbookPath))
            {
                Console.WriteLine("Workbook : (none)");
                Console.WriteLine("Sheet    : (none)");
                Console.WriteLine("Excel    : Unknown");
                return ExitCodes.SessionNotEstablished;
            }

            var excelRunning = ExcelBridge.TryGetRunningExcel() != null ? "Running" : "Not running";

            Console.WriteLine($"Workbook : {session.WorkbookPath}");
            Console.WriteLine($"Sheet    : {session.DefaultSheet ?? "(active)"}");
            Console.WriteLine($"Excel    : {excelRunning}");

            return ExitCodes.Success;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return ExitCodes.ExcelOperationFailed;
        }
    }
}
