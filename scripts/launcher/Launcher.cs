using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

[assembly: AssemblyTitle("WorkSpace")]
[assembly: AssemblyDescription("WorkSpace Code & Document Editor")]
[assembly: AssemblyCompany("WorkSpace")]
[assembly: AssemblyProduct("WorkSpace")]
[assembly: AssemblyCopyright("Copyright © 2026")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string here = AppDomain.CurrentDomain.BaseDirectory;
            string app = Path.Combine(here, "app", "WorkSpace.exe");

            if (!File.Exists(app))
            {
                return 1;
            }

            ProcessStartInfo start = new ProcessStartInfo(app);
            start.WorkingDirectory = Path.GetDirectoryName(app);
            start.UseShellExecute = true;

            if (args != null && args.Length > 0)
            {
                start.Arguments = string.Join(" ", Array.ConvertAll(args, a =>
                    a.Contains(" ") ? "\"" + a + "\"" : a));
            }

            Process.Start(start);
            return 0;
        }
        catch
        {
            return 1;
        }
    }
}
