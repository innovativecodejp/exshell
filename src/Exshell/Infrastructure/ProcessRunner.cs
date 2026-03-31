using System.Diagnostics;

namespace Exshell.Infrastructure;

/// <summary>
/// 外部プロセス（主に wsl）の起動・実行結果取得を担当する。
/// </summary>
public static class ProcessRunner
{
    /// <summary>
    /// wsl コマンドを実行し、stdout を Console.Out へ、
    /// stderr を Console.Error へそれぞれ流す。
    /// </summary>
    /// <returns>プロセス終了コード</returns>
    public static int RunWsl(string arguments)
    {
        var psi = new ProcessStartInfo
        {
            FileName               = "wsl",
            Arguments              = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
        };

        using var proc = Process.Start(psi)
            ?? throw new ExshellException("Failed to start WSL process.", ExitCodes.WslExecutionFailed);

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
}
