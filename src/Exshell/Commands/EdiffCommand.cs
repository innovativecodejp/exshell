using System.Diagnostics;
using System.Text;
using Exshell.Excel;
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

        var target1 = args[0];
        var target2 = args[1];

        string? tmp1 = null;
        string? tmp2 = null;

        try
        {
            var session = SessionStore.LoadOrThrow();
            var (sheet1, shape1) = EcatCommand.ParseTarget(target1, session.DefaultSheet);
            var (sheet2, shape2) = EcatCommand.ParseTarget(target2, session.DefaultSheet);

            var app = ExcelBridge.GetOrCreateApplication();
            var wb  = ExcelBridge.OpenOrGetWorkbook(app, session.WorkbookPath);

            var ws1   = ExcelBridge.GetWorksheet(wb, sheet1);
            var text1 = ExcelBridge.GetShapeText(ws1, shape1);

            var ws2   = ExcelBridge.GetWorksheet(wb, sheet2);
            var text2 = ExcelBridge.GetShapeText(ws2, shape2);

            // 一時ファイルに UTF-8 LF で書き出す
            tmp1 = Path.GetTempFileName();
            tmp2 = Path.GetTempFileName();
            WriteUtf8Lf(tmp1, text1);
            WriteUtf8Lf(tmp2, text2);

            // Windows パス → WSL パスに変換
            var wslPath1 = ToWslPath(tmp1);
            var wslPath2 = ToWslPath(tmp2);

            // wsl diff を実行
            var exitCode = RunWslDiff(wslPath1, wslPath2);

            // diff は差分あり=1、エラー=2
            if (exitCode == 2)
            {
                Console.Error.WriteLine("WSL diff execution failed.");
                return ExitCodes.WslExecutionFailed;
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
            return ExitCodes.ExcelOperationFailed;
        }
        finally
        {
            TryDelete(tmp1);
            TryDelete(tmp2);
        }
    }

    private static void WriteUtf8Lf(string path, string text)
    {
        // 改行を LF に統一
        var lf = text.Replace("\r\n", "\n").Replace("\r", "\n");
        File.WriteAllText(path, lf, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>
    /// C:\foo\bar.txt → /mnt/c/foo/bar.txt
    /// </summary>
    private static string ToWslPath(string winPath)
    {
        if (winPath.Length >= 2 && winPath[1] == ':')
        {
            var drive = char.ToLower(winPath[0]);
            var rest  = winPath[2..].Replace('\\', '/');
            return $"/mnt/{drive}{rest}";
        }
        return winPath.Replace('\\', '/');
    }

    private static int RunWslDiff(string wslPath1, string wslPath2)
    {
        var psi = new ProcessStartInfo
        {
            FileName               = "wsl",
            Arguments              = $"diff \"{wslPath1}\" \"{wslPath2}\"",
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
        };

        using var proc = Process.Start(psi)
            ?? throw new ExshellException("Failed to start WSL process.", ExitCodes.WslExecutionFailed);

        // diff の出力をそのまま stdout へ流す
        proc.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null) Console.WriteLine(e.Data);
        };
        proc.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null) Console.Error.WriteLine(e.Data);
        };
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();
        proc.WaitForExit();

        return proc.ExitCode;
    }

    private static void TryDelete(string? path)
    {
        if (path == null) return;
        try { File.Delete(path); } catch { }
    }
}
