using Exshell.Excel;
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
            var fullPath = Path.GetFullPath(filePath);
            var app      = ExcelBridge.GetOrCreateApplication();
            var wb       = ExcelBridge.OpenOrGetWorkbook(app, fullPath);

            // シート名が指定されていれば存在確認
            var ws          = ExcelBridge.GetWorksheet(wb, sheet);
            var defaultSheet = ws.Name;

            var session = new ExshellSession
            {
                WorkbookPath = fullPath,
                DefaultSheet = defaultSheet,
            };
            SessionStore.Save(session);

            Console.WriteLine($"Workbook : {fullPath}");
            Console.WriteLine($"Sheet    : {defaultSheet}");
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
