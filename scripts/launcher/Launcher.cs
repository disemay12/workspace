using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

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
            start.UseShellExecute = false;

            // ELECTRON_RUN_AS_NODE превращает бинарник в обычный Node: окно не
            // создаётся и процесс сразу завершается. Среда запуска может
            // унаследовать переменную, поэтому она сбрасывается явно.
            start.EnvironmentVariables.Remove("ELECTRON_RUN_AS_NODE");

            string arguments = "";

            // При отключённом UAC процесс получает полный токен администратора,
            // и Chromium не может поднять sandbox: бинарник с именем, отличным
            // от electron.exe, падает с STATUS_BREAKPOINT ещё до загрузки
            // main.js. Флаг должен прийти снаружи — повторный запуск из main.js
            // опаздывает. Повышенный токен и означает этот режим.
            if (HasElevatedToken())
            {
                arguments = "--no-sandbox";
            }

            if (args != null && args.Length > 0)
            {
                string passed = string.Join(" ", Array.ConvertAll(args, a =>
                    a.Contains(" ") ? "\"" + a + "\"" : a));
                arguments = arguments.Length == 0 ? passed : arguments + " " + passed;
            }

            start.Arguments = arguments;

            Process.Start(start);
            return 0;
        }
        catch
        {
            return 1;
        }
    }

    // TokenElevation (20): ненулевое значение означает полный токен
    // администратора, который выдаётся либо при отключённом UAC, либо при
    // запуске с повышением. В обоих случаях Chromium sandbox неработоспособен.
    private static bool HasElevatedToken()
    {
        try
        {
            IntPtr token;
            if (!OpenProcessToken((IntPtr)(-1), 0x8, out token))
            {
                return false;
            }
            try
            {
                TokenElevation elevation;
                int returned;
                int size = Marshal.SizeOf(typeof(TokenElevation));
                if (!GetTokenInformation(token, 20, out elevation, size, out returned))
                {
                    return false;
                }
                return elevation.TokenIsElevated != 0;
            }
            finally
            {
                CloseHandle(token);
            }
        }
        catch
        {
            return false;
        }
    }

    private struct TokenElevation
    {
        public uint TokenIsElevated;
    }

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool GetTokenInformation(
        IntPtr token, int cls, out TokenElevation elevation, int size, out int returned);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);
}
