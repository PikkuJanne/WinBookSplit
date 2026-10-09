using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

// Native acceptance control, never shipped as an application dependency.
public static class OwnedConverter {
    public static int Main(string[] args) {
        if (args.Length == 1 && args[0] == "--version") {
            Console.WriteLine("ebook-convert.exe (calibre 9.15.0)");
            return 0;
        }
        string receipt = Environment.GetEnvironmentVariable("WBS_OUTCOME_RECEIPT");
        string childReceipt = Environment.GetEnvironmentVariable("WBS_OUTCOME_CHILD_RECEIPT");
        if (args.Length == 1 && args[0] == "--owned-child") {
            File.WriteAllText(childReceipt, "{\"pid\":" + Process.GetCurrentProcess().Id + "}");
            Thread.Sleep(20000);
            return 0;
        }
        if (args.Length < 2) return 3;
        string kind = Environment.GetEnvironmentVariable("WBS_OUTCOME_KIND");
        string escaped = args[1].Replace("\\", "\\\\").Replace("\"", "\\\"");
        File.WriteAllText(receipt, "{\"pid\":" + Process.GetCurrentProcess().Id +
            ",\"output\":\"" + escaped + "\",\"kind\":\"" + kind + "\"}");
        Console.Error.WriteLine("AUTHORED_CONVERTER_FINAL_STDERR");
        if (kind == "converter-failure") return 17;
        var info = new ProcessStartInfo(Process.GetCurrentProcess().MainModule.FileName, "--owned-child");
        info.UseShellExecute = false;
        info.CreateNoWindow = true;
        Process.Start(info);
        for (int i = 0; i < 200 && !File.Exists(childReceipt); i++) Thread.Sleep(5);
        Thread.Sleep(20000);
        return 0;
    }
}
