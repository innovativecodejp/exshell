using System.Text.Json;

namespace Exshell.Session;

public static class SessionStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static string SessionPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Exshell",
            "session.json"
        );

    public static SessionInfo? Load()
    {
        if (!File.Exists(SessionPath))
            return null;

        try
        {
            var json = File.ReadAllText(SessionPath);
            return JsonSerializer.Deserialize<SessionInfo>(json);
        }
        catch
        {
            return null;
        }
    }

    public static void Save(SessionInfo session)
    {
        session.UpdatedAt = DateTime.Now;
        var dir = Path.GetDirectoryName(SessionPath)!;
        Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(session, JsonOptions);
        File.WriteAllText(SessionPath, json);
    }

    public static SessionInfo LoadOrThrow()
    {
        var session = Load();
        if (session == null || string.IsNullOrEmpty(session.WorkbookPath))
            throw new ExshellException(
                "No active session. Run 'eopen <file>' first.",
                ExitCodes.SessionNotEstablished
            );
        return session;
    }
}
