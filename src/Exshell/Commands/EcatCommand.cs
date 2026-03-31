using Exshell.Application.Models;
using Exshell.ExcelInterop;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// ecat &lt;textbox&gt;
/// テキストボックス内容を stdout へ出力する。
/// </summary>
public static class EcatCommand
{
    public static int Run(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: ecat <textbox>");
            Console.Error.WriteLine("  textbox: ShapeName or SheetName:ShapeName");
            return ExitCodes.ArgumentError;
        }

        try
        {
            var session  = SessionStore.LoadOrThrow();
            var shapeRef = ShapeReference.Parse(args[0], session.DefaultSheetName);

            var app  = ExcelAppGateway.GetOrCreateApplication();
            var wb   = WorkbookResolver.OpenOrGetWorkbook(app, session.WorkbookPath);
            var ws   = WorksheetResolver.GetWorksheet(wb, shapeRef.SheetName);
            var text = ShapeTextAccessor.GetShapeText(ws, shapeRef.ShapeName);

            Console.Write(text);

            // 末尾に改行がない場合は追加してパイプ利用しやすくする
            if (text.Length > 0 && text[^1] != '\n')
                Console.WriteLine();

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
