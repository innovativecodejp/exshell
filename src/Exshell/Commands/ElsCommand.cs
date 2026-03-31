using Exshell.Excel;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// els [--sheet &lt;sheet name&gt;]
/// </summary>
public static class ElsCommand
{
    public static int Run(string[] args)
    {
        string? sheetOverride = null;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--sheet" && i + 1 < args.Length)
                sheetOverride = args[++i];
            else
            {
                Console.Error.WriteLine($"Unknown option: {args[i]}");
                return ExitCodes.ArgumentError;
            }
        }

        try
        {
            var session = SessionStore.LoadOrThrow();
            var sheet   = sheetOverride ?? session.DefaultSheet;

            var app = ExcelBridge.GetOrCreateApplication();
            var wb  = ExcelBridge.OpenOrGetWorkbook(app, session.WorkbookPath);
            var ws  = ExcelBridge.GetWorksheet(wb, sheet);

            var shapes = ExcelBridge.ListTextShapes(ws);

            Console.WriteLine($"[{ws.Name}]");
            foreach (var name in shapes)
                Console.WriteLine(name);

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
