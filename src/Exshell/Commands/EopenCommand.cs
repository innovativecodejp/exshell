using Exshell.ExcelInterop;
using Exshell.Infrastructure;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// eopen &lt;excel file&gt; [--sheet &lt;sheet name&gt;]
/// </summary>
public static class EopenCommand
{
    public static int Run(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: eopen <excel file> [--sheet <sheet name>]");
            return ExitCodes.ArgumentError;
        }

        var filePath  = args[0];
        string? sheet = null;

        for (int i = 1; i < args.Length; i++)
        {
            if (args[i] == "--sheet" && i + 1 < args.Length)
                sheet = args[++i];
            else
            {
                Console.Error.WriteLine($"Unknown option: {args[i]}");
                return ExitCodes.ArgumentError;
            }
        }

        try
        {
            var fullPath = PathConverter.NormalizeWindowsPath(filePath);
            var app      = ExcelAppGateway.GetOrCreateApplication();
            var wb       = WorkbookResolver.OpenOrGetWorkbook(app, fullPath);

            // シート名が指定されていれば存在確認、なければアクティブシートを採用
            var ws              = WorksheetResolver.GetWorksheet(wb, sheet);
            var defaultSheetName = ws.Name;

            var session = new SessionInfo
            {
                WorkbookPath     = fullPath,
                DefaultSheetName = defaultSheetName,
            };
            SessionStore.Save(session);

            Console.WriteLine($"Workbook : {fullPath}");
            Console.WriteLine($"Sheet    : {defaultSheetName}");
            return ExitCodes.Success;
        }
        catch (ExshellException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return ex.ExitCode;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return ExitCodes.ExcelOperationFailed;
        }
    }
}
