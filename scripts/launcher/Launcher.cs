using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

/// <summary>
/// Оконный загрузчик WorkSpace.
///
/// Chromium не может поднять sandbox, когда в Windows отключён UAC
/// (EnableLUA = 0), и процесс завершается ещё до появления окна.
/// Загрузчик проверяет это один раз и при необходимости запускает
/// приложение с --no-sandbox. Консольного окна нет ни в одном режиме.
/// </summary>
internal static class Program
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetCommandLineW();

    private static int Main()
    {
        string here = AppDomain.CurrentDomain.BaseDirectory;
        string app = Path.Combine(here, "app", "WorkSpace.exe");

        if (!File.Exists(app))
        {
            return 1;
        }

        ProcessStartInfo start = new ProcessStartInfo();
        start.FileName = app;
        start.WorkingDirectory = Path.GetDirectoryName(app);
        start.UseShellExecute = false;
        start.CreateNoWindow = true;

        // Родительская оболочка может выставить ELECTRON_RUN_AS_NODE,
        // и тогда Electron работает как обычный Node и отвергает свои флаги.
        start.EnvironmentVariables.Remove("ELECTRON_RUN_AS_NODE");

        string userArgs = UserArguments();
        if (!IsUacEnabled())
        {
            start.Arguments = string.IsNullOrEmpty(userArgs) ? "--no-sandbox" : "--no-sandbox " + userArgs;
        }
        else
        {
            start.Arguments = userArgs;
        }

        try
        {
            Process.Start(start);
        }
        catch
        {
            return 1;
        }

        return 0;
    }

    /// <summary>Аргументы пользователя без пути к самому загрузчику.</summary>
    private static string UserArguments()
    {
        string raw = Marshal.PtrToStringUni(GetCommandLineW()) ?? string.Empty;
        raw = raw.Trim();
        if (raw.Length == 0)
        {
            return string.Empty;
        }

        if (raw[0] == '"')
        {
            int end = raw.IndexOf('"', 1);
            if (end >= 0 && end + 1 < raw.Length)
            {
                return raw.Substring(end + 1).Trim();
            }
            return string.Empty;
        }

        int space = raw.IndexOf(' ');
        if (space >= 0)
        {
            return raw.Substring(space + 1).Trim();
        }
        return string.Empty;
    }

    /// <summary>true, если UAC включён и sandbox Chromium работоспособен.</summary>
    private static bool IsUacEnabled()
    {
        try
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"))
            {
                if (key != null)
                {
                    object value = key.GetValue("EnableLUA");
                    if (value is int)
                    {
                        return (int)value != 0;
                    }
                }
            }
        }
        catch
        {
            // Политику прочитать не удалось — не отключаем sandbox.
        }

        return true;
    }
}
