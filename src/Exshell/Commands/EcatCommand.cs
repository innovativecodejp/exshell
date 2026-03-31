using Exshell.Excel;
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

        var target = args[0];

        try
        {
            var session = SessionStore.LoadOrThrow();
            var (sheetName, shapeName) = ParseTarget(target, session.DefaultSheet);

            var app  = ExcelBridge.GetOrCreateApplication();
            var wb   = ExcelBridge.OpenOrGetWorkbook(app, session.WorkbookPath);
            var ws   = ExcelBridge.GetWorksheet(wb, sheetName);
            var text = ExcelBridge.GetShapeText(ws, shapeName);

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

    /// <summary>
    /// "Sheet:Shape" または "Shape" を解析する。
    /// </summary>
    internal static (string? sheetName, string shapeName) ParseTarget(string target, string? defaultSheet)
    {
        var colon = target.IndexOf(':');
        if (colon >= 0)
            return (target[..colon], target[(colon + 1)..]);
        return (defaultSheet, target);
    }
}
