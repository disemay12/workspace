using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

[assembly: AssemblyTitle("WorkSpace")]
[assembly: AssemblyDescription("WorkSpace Code & Document Editor")]
[assembly: AssemblyCompany("WorkSpace")]
[assembly: AssemblyProduct("WorkSpace")]
[assembly: AssemblyCopyright("Copyright © 2026")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

/**
 * Загрузчик portable-сборки.
 *
 * Кроме запуска приложения он применяет подготовленное обновление: Electron
 * держит resources/app.asar открытым, поэтому заменять его из работающего
 * приложения нельзя. Приложение скачивает архив, распаковывает его и пишет
 * маркер pending-update.txt, а загрузчик копирует файлы до старта приложения
 * и перезапускает его после выхода с кодом 42.
 */
internal static class Program
{
    /** Приложение завершилось с этим кодом — нужно применить обновление и перезапустить. */
    private const int UpdateExitCode = 42;

    [STAThread]
    private static int Main(string[] args)
    {
        string here = AppDomain.CurrentDomain.BaseDirectory;
        string app = Path.Combine(here, "app", "WorkSpace.exe");
        string userData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WorkSpace");

        try
        {
            Directory.CreateDirectory(userData);
        }
        catch
        {
            // Профиль недоступен — работаем без него, приложение создаст свой.
        }

        WriteLauncherInfo(userData);
        ApplyPendingUpdate(here, userData);
        ReplaceStagedExecutables(here, userData);

        // Обычно цикл выполняется один раз; вторая итерация — перезапуск
        // после успешно применённого обновления.
        for (int attempt = 0; attempt < 2; attempt++)
        {
            if (!File.Exists(app))
            {
                Log(userData, "app not found: " + app);
                return 1;
            }

            try
            {
                ProcessStartInfo start = BuildStartInfo(app, args, here);
                using (Process child = Process.Start(start))
                {
                    if (child == null)
                    {
                        return 1;
                    }

                    child.WaitForExit();
                    int code = child.ExitCode;

                    if (code == UpdateExitCode)
                    {
                        // Приложение попросило применить обновление. Запускаем
                        // его снова даже при неудаче: молчаливый выход, после
                        // которого ничего не происходит, хуже всего.
                        bool applied = ApplyPendingUpdate(here, userData);
                        ReplaceStagedExecutables(here, userData);
                        Log(userData, "update requested, applied=" + applied);
                        continue;
                    }

                    return code;
                }
            }
            catch (Exception error)
            {
                Log(userData, "start failed: " + error.Message);
                return 1;
            }
        }

        return 0;
    }

    private static ProcessStartInfo BuildStartInfo(string app, string[] args, string here)
    {
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
        // от electron.exe, падает с STATUS_BREAKPOINT ещё до загрузки main.js.
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
        return start;
    }

    /** Сообщает приложению, что загрузчик умеет применять обновление и перезапускать его. */
    private static void WriteLauncherInfo(string userData)
    {
        try
        {
            string json = "{\"autoRestart\":true,\"launcherVersion\":\""
                + Assembly.GetExecutingAssembly().GetName().Version.ToString()
                + "\",\"writtenAt\":\"" + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) + "\"}";
            File.WriteAllText(Path.Combine(userData, "launcher.json"), json, new UTF8Encoding(false));
        }
        catch
        {
            // Не критично: приложение просто предупредит о ручном перезапуске.
        }
    }

    /** Ставит на место новые WorkSpace.exe, скачанные предыдущим обновлением. */
    private static void ReplaceStagedExecutables(string here, string userData)
    {
        string appExe = Path.Combine(here, "app", "WorkSpace.exe");
        string stagedApp = appExe + ".new";
        if (File.Exists(stagedApp))
        {
            try
            {
                File.Copy(stagedApp, appExe, true);
                File.Delete(stagedApp);
                Log(userData, "electron runtime replaced");
            }
            catch (Exception error)
            {
                Log(userData, "electron runtime replace failed: " + error.Message);
            }
        }

        string launcher = Path.Combine(here, "WorkSpace.exe");
        string stagedLauncher = launcher + ".new";
        if (File.Exists(stagedLauncher))
        {
            try
            {
                // Запущенный образ нельзя перезаписать, но можно переименовать.
                string retired = launcher + ".old";
                if (File.Exists(retired))
                {
                    File.Delete(retired);
                }
                File.Move(launcher, retired);
                File.Move(stagedLauncher, launcher);
                Log(userData, "launcher replaced");
            }
            catch (Exception error)
            {
                Log(userData, "launcher replace failed: " + error.Message);
            }
        }
    }

    /** Применяет подготовленное обновление. true — файлы обновлены. */
    private static bool ApplyPendingUpdate(string here, string userData)
    {
        string marker = Path.Combine(userData, "pending-update.txt");
        if (!File.Exists(marker))
        {
            return false;
        }

        string source = null;
        string version = "unknown";
        try
        {
            foreach (string line in File.ReadAllLines(marker))
            {
                int separator = line.IndexOf('=');
                if (separator <= 0)
                {
                    continue;
                }
                string key = line.Substring(0, separator).Trim();
                string value = line.Substring(separator + 1).Trim();
                if (key.Equals("source", StringComparison.OrdinalIgnoreCase))
                {
                    source = value;
                }
                else if (key.Equals("version", StringComparison.OrdinalIgnoreCase))
                {
                    version = value;
                }
            }
        }
        catch (Exception error)
        {
            Log(userData, "pending update read failed: " + error.Message);
        }

        if (string.IsNullOrEmpty(source) || !File.Exists(Path.Combine(source, "app", "WorkSpace.exe")))
        {
            Log(userData, "pending update has no valid source: " + (source ?? "<none>"));
            TryDelete(marker);
            return false;
        }

        string asar = Path.Combine(here, "app", "resources", "app.asar");
        try
        {
            if (File.Exists(asar))
            {
                File.Copy(asar, asar + ".bak-before-" + version, true);
            }
        }
        catch (Exception error)
        {
            Log(userData, "asar backup failed: " + error.Message);
        }

        int copied = 0;
        int failed = 0;
        CopyTree(source, here, userData, ref copied, ref failed);

        Log(userData, "applied " + version + ": copied " + copied + ", failed " + failed);

        string root = Path.GetDirectoryName(source);
        TryDelete(marker);
        try
        {
            Directory.Delete(source, true);
        }
        catch
        {
            // Каталог обновления мог остаться заблокированным — не критично.
        }
        if (!string.IsNullOrEmpty(root) && root.IndexOf("WorkSpace-update-", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            try
            {
                Directory.Delete(root, true);
            }
            catch
            {
                // Архив удалится при следующем обновлении.
            }
        }

        return failed == 0;
    }

    /** Рекурсивно копирует дерево обновления на место установленной программы. */
    private static void CopyTree(string source, string target, string userData, ref int copied, ref int failed)
    {
        string[] files;
        try
        {
            files = Directory.GetFiles(source, "*", SearchOption.AllDirectories);
        }
        catch (Exception error)
        {
            Log(userData, "update listing failed: " + error.Message);
            failed += 1;
            return;
        }

        foreach (string file in files)
        {
            string relative = file.Substring(source.Length).TrimStart('\\', '/');
            if (ShouldSkip(relative))
            {
                continue;
            }

            string destination = Path.Combine(target, relative);
            // Загрузчик, который сейчас выполняется, заменяем только файлом .new.
            if (relative.Equals("WorkSpace.exe", StringComparison.OrdinalIgnoreCase))
            {
                destination += ".new";
            }

            try
            {
                string directory = Path.GetDirectoryName(destination);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                File.Copy(file, destination, true);
                copied += 1;
            }
            catch (Exception error)
            {
                failed += 1;
                Log(userData, "copy failed " + relative + ": " + error.Message);
            }
        }
    }

    private static bool ShouldSkip(string relative)
    {
        string name = Path.GetFileName(relative);
        if (name.Equals("pending-update.txt", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        // Резервные копии и снятые образы прошлых обновлений
        return relative.IndexOf(".bak-", StringComparison.OrdinalIgnoreCase) >= 0
            || name.EndsWith(".old", StringComparison.OrdinalIgnoreCase);
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // Файл занят — маркер удалится при следующем запуске.
        }
    }

    private static void Log(string userData, string message)
    {
        try
        {
            string line = "[" + DateTime.Now.ToString("o", CultureInfo.InvariantCulture) + "] " + message
                + Environment.NewLine;
            File.AppendAllText(Path.Combine(userData, "update.log"), line, new UTF8Encoding(false));
        }
        catch
        {
            // Логирование не должно мешать запуску.
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
