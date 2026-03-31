using Exshell.Application.Models;
using Exshell.ExcelInterop;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// cate &lt;textbox&gt; [--append]
/// stdin をテキストボックスへ書き込む。
/// </summary>
public static class CateCommand
{
    public static int Run(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: cate <textbox> [--append]");
            Console.Error.WriteLine("  textbox: ShapeName or SheetName:ShapeName");
            return ExitCodes.ArgumentError;
        }

        var target = args[0];
        var append = false;

        for (int i = 1; i < args.Length; i++)
        {
            if (args[i] == "--append")
                append = true;
            else
            {
                Console.Error.WriteLine($"Unknown option: {args[i]}");
                return ExitCodes.ArgumentError;
            }
        }

        try
        {
            var session  = SessionStore.LoadOrThrow();
            var shapeRef = ShapeReference.Parse(target, session.DefaultSheetName);
            var text     = Console.In.ReadToEnd();

            var app = ExcelAppGateway.GetOrCreateApplication();
            var wb  = WorkbookResolver.OpenOrGetWorkbook(app, session.WorkbookPath);
            var ws  = WorksheetResolver.GetWorksheet(wb, shapeRef.SheetName);

            ShapeTextAccessor.SetShapeText(ws, shapeRef.ShapeName, text, append);

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
