namespace Exshell.Session;

/// <summary>
/// セッション情報。~/.exshell/session.json に永続化される。
/// </summary>
public sealed class SessionInfo
{
    public string   WorkbookPath     { get; set; } = string.Empty;
    public string?  DefaultSheetName { get; set; }
    public DateTime UpdatedAt        { get; set; }
}
