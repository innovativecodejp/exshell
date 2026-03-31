namespace Exshell;

public static class ExitCodes
{
    public const int Success               = 0;
    public const int ArgumentError         = 1;
    public const int SessionNotEstablished = 2;
    public const int WorkbookNotFound      = 3;
    public const int WorksheetNotFound     = 4;
    public const int ShapeNotFound         = 5;
    public const int ExcelOperationFailed  = 6;
    public const int WslExecutionFailed    = 7;
}
