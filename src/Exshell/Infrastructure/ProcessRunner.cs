using System.Diagnostics;
using System.Text;

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
        if (string.IsNullOrEmpty(arguments))
            throw new ExshellException("WSL command is required.", ExitCodes.ArgumentError);

        var psi = new ProcessStartInfo
        {
            FileName               = "wsl",
            Arguments              = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding  = Encoding.UTF8,
            UseShellExecute        = false,
            CreateNoWindow         = true,
        };

        using var proc = Process.Start(psi)
            ?? throw new ExshellException("Failed to start WSL process.", ExitCodes.DiffExecutionFailed);

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        proc.WaitForExit();

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();

        if (!string.IsNullOrEmpty(stdout))
            Console.Write(stdout);

        if (!string.IsNullOrEmpty(stderr))
            Console.Error.Write(stderr);

        return proc.ExitCode;
    }
}
