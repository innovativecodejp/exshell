using Exshell;
using Exshell.Commands;
using System.Text;

var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
Console.InputEncoding = utf8NoBom;
Console.OutputEncoding = utf8NoBom;

if (args.Length == 0)
{
    PrintUsage();
    return ExitCodes.ArgumentError;
}

return args[0].ToLowerInvariant() switch
{
    "eopen" => EopenCommand.Run(args[1..]),
    "els"   => ElsCommand.Run(args[1..]),
    "ecat"  => EcatCommand.Run(args[1..]),
    "cate"  => CateCommand.Run(args[1..]),
    "ediff" => EdiffCommand.Run(args[1..]),
    "einfo" => EinfoCommand.Run(args[1..]),
    _       => UnknownCommand(args[0]),
};

static void PrintUsage()
{
    Console.Error.WriteLine("Exshell - Excel textbox CLI bridge");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Usage: exshell <command> [args]");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Commands:");
    Console.Error.WriteLine("  eopen <file> [--sheet <name>]   Open workbook and set session");
    Console.Error.WriteLine("  els [--sheet <name>]            List textboxes on sheet");
    Console.Error.WriteLine("  ecat <textbox>                  Print textbox content to stdout");
    Console.Error.WriteLine("  cate <textbox> [--append]       Write stdin to textbox");
    Console.Error.WriteLine("  ediff <textbox1> <textbox2>     Diff two textboxes via WSL diff");
    Console.Error.WriteLine("  einfo                           Show current session info");
}

static int UnknownCommand(string name)
{
    Console.Error.WriteLine($"Unknown command: {name}");
    Console.Error.WriteLine("Run 'exshell' with no arguments to see usage.");
    return ExitCodes.ArgumentError;
}
