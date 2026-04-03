using Exshell.Application.Models;
using Exshell.ExcelInterop;
using Exshell.Infrastructure;
using Exshell.Session;

namespace Exshell.Commands;

/// <summary>
/// ediff &lt;textbox1&gt; &lt;textbox2&gt;
/// 2 つのテキストボックスを一時ファイル化して WSL diff を実行する。
/// </summary>
public static class EdiffCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: ediff <textbox1> <textbox2>");
            return ExitCodes.ArgumentError;
        }

        string? tmp1 = null;
        string? tmp2 = null;

        try
        {
            var session   = SessionStore.LoadOrThrow();
            var shapeRef1 = ShapeReference.Parse(args[0], session.DefaultSheetName);
            var shapeRef2 = ShapeReference.Parse(args[1], session.DefaultSheetName);

            var app = ExcelAppGateway.GetOrCreateApplication();
            var wb  = WorkbookResolver.OpenOrGetWorkbook(app, session.WorkbookPath);

            var text1 = ShapeTextAccessor.GetShapeText(
                WorksheetResolver.GetWorksheet(wb, shapeRef1.SheetName),
                shapeRef1.ShapeName
            );
            var text2 = ShapeTextAccessor.GetShapeText(
                WorksheetResolver.GetWorksheet(wb, shapeRef2.SheetName),
                shapeRef2.ShapeName
            );

            // 一時ファイルに UTF-8 LF で書き出す
            tmp1 = TempFileService.WriteUtf8Lf(text1);
            tmp2 = TempFileService.WriteUtf8Lf(text2);

            // WSL パスに変換して diff 実行
            var wslPath1 = PathConverter.ToWslPath(tmp1);
            var wslPath2 = PathConverter.ToWslPath(tmp2);

            var exitCode = ProcessRunner.RunWsl($"diff \"{wslPath1}\" \"{wslPath2}\"");

            // diff: 0=差分なし, 1=差分あり, 2=エラー
            if (exitCode == 2)
            {
                Console.Error.WriteLine("WSL diff execution failed.");
                return ExitCodes.DiffExecutionFailed;
            }

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
            return ExitCodes.UnexpectedError;
        }
        finally
        {
            TempFileService.TryDelete(tmp1);
            TempFileService.TryDelete(tmp2);
        }
    }
}
