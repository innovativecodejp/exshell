namespace Exshell;

public sealed class ExshellException : Exception
{
    public int ExitCode { get; }

    public ExshellException(string message, int exitCode) : base(message)
    {
        ExitCode = exitCode;
    }
}
