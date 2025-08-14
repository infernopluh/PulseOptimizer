using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Management;


namespace PCOptimizer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Tabs.SelectedIndex = 0;
            UpdateTitle();
            LoadSysInfo();
            LoadStartupItems();
        }

        private void UpdateTitle()
        {
            TitleBlock.Text = ((Tabs.SelectedItem as TabItem)?.Header?.ToString()) ?? "Dashboard";
        }

        private void Nav_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && int.TryParse((string)b.Tag, out int idx))
            {
                Tabs.SelectedIndex = idx;
                UpdateTitle();
            }
        }

        #region System Info
        private void LoadSysInfo()
        {
            try
            {
                var os = Environment.OSVersion.VersionString;
                var mem = GetTotalMemoryInGB();
                var isAdmin = IsRunningAsAdmin() ? "Yes" : "No";
                SysInfoText.Text = $"OS: {os}\nRAM: ~{mem:0.0} GB\nAdmin: {isAdmin}";
            }
            catch (Exception ex)
            {
                SysInfoText.Text = "Unable to read system info: " + ex.Message;
            }
        }

        private static double GetTotalMemoryInGB()
        {
            try
            {
                MEMORYSTATUSEX memStatus = new MEMORYSTATUSEX();
                if (GlobalMemoryStatusEx(memStatus))
                {
                    return memStatus.ullTotalPhys / 1024.0 / 1024.0 / 1024.0;
                }
                return 0.0;
            }
            catch
            {
                return 0.0;
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
            public MEMORYSTATUSEX()
            {
                dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]

        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);
        private static bool IsRunningAsAdmin()
        {
            using var id = WindowsIdentity.GetCurrent();
            var pr = new WindowsPrincipal(id);
            return pr.IsInRole(WindowsBuiltInRole.Administrator);
        }
        #endregion

        #region One-Click Optimize
        private async void RunOptimize_Click(object sender, RoutedEventArgs e)
        {
            StatusBlock.Text = "Starting optimization...";
            MainProgress.Value = 0;

            await Task.Delay(150);
            var tempResult = await Task.Run(() => Cleaner.ClearTemp());
            MainProgress.Value = 25; StatusBlock.Text = "Temp cleaned.";

            var browserResult = await Task.Run(() => Cleaner.ClearBrowserCaches());
            MainProgress.Value = 45; StatusBlock.Text = "Browser cache cleared.";

            var wuResult = await Task.Run(() => Cleaner.ClearWindowsUpdateCache());
            MainProgress.Value = 65; StatusBlock.Text = "Windows Update cache cleared.";

            var binResult = Cleaner.EmptyRecycleBin();
            MainProgress.Value = 80; StatusBlock.Text = "Recycle Bin emptied.";

            var freed = await Task.Run(() => Performance.FreeRAM());
            MainProgress.Value = 95; StatusBlock.Text = "RAM optimized.";

            MainProgress.Value = 100;
            StatusBlock.Text = $"Done. Temp: {FormatBytes(tempResult.BytesDeleted)}, Browser: {FormatBytes(browserResult.BytesDeleted)}, WU: {FormatBytes(wuResult.BytesDeleted)}, Bin: {binResult}. RAM Freed: ~{FormatBytes(freed)}";
        }
        #endregion

        #region Cleaner Buttons
        private void ClearTemp_Click(object sender, RoutedEventArgs e)
        {
            var r = Cleaner.ClearTemp();
            CleanerResult.Text = $"Temp cleaned. Deleted {r.FilesDeleted} files, {FormatBytes(r.BytesDeleted)}.";
        }

        private void ClearBrowser_Click(object sender, RoutedEventArgs e)
        {
            var r = Cleaner.ClearBrowserCaches();
            CleanerResult.Text = $"Browser caches cleared. Deleted {r.FilesDeleted} files, {FormatBytes(r.BytesDeleted)}.";
        }

        private void ClearWU_Click(object sender, RoutedEventArgs e)
        {
            var r = Cleaner.ClearWindowsUpdateCache();
            CleanerResult.Text = $"Windows Update cache cleared. Deleted {r.FilesDeleted} files, {FormatBytes(r.BytesDeleted)}.";
        }

        private void EmptyBin_Click(object sender, RoutedEventArgs e)
        {
            var ok = Cleaner.EmptyRecycleBin();
            CleanerResult.Text = ok ? "Recycle Bin emptied." : "Recycle Bin already empty or not accessible.";
        }
        #endregion

        #region Startup Manager
        private void LoadStartupItems()
        {
            try
            {
                var items = StartupManager.GetAll();
                StartupList.ItemsSource = items;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load startup items: " + ex.Message);
            }
        }

        private void StartupRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadStartupItems();
        }

        private void StartupDisable_Click(object sender, RoutedEventArgs e)
        {
            if (StartupList.SelectedItem is StartupItem item)
            {
                if (StartupManager.Disable(item))
                {
                    PerfResult.Text = $"Disabled: {item.Name}";
                    LoadStartupItems();
                }
                else
                {
                    PerfResult.Text = $"Failed to disable: {item.Name}";
                }
            }
        }

        private void StartupEnable_Click(object sender, RoutedEventArgs e)
        {
            if (StartupList.SelectedItem is StartupItem item)
            {
                if (StartupManager.Enable(item))
                {
                    PerfResult.Text = $"Enabled: {item.Name}";
                    LoadStartupItems();
                }
                else
                {
                    PerfResult.Text = $"Failed to enable: {item.Name}";
                }
            }
        }
        #endregion

        #region Performance Buttons
        private void FreeRAM_Click(object sender, RoutedEventArgs e)
        {
            var freed = Performance.FreeRAM();
            PerfResult.Text = $"Requested RAM trim. Estimated freed: {FormatBytes(freed)} (approx).";
        }

        private void KillBloat_Click(object sender, RoutedEventArgs e)
        {
            var (killed, failed) = Performance.KillCommonBloat();
            PerfResult.Text = $"Killed {killed} background processes; {failed} could not be killed.";
        }

        private void SetBalanced_Click(object sender, RoutedEventArgs e)
        {
            var ok = PowerPlans.SetBalanced();
            StatusBlock.Text = ok ? "Balanced power plan set." : "Failed to set Balanced plan.";
        }

        private void SetHighPerf_Click(object sender, RoutedEventArgs e)
        {
            var ok = PowerPlans.SetHighPerformance();
            StatusBlock.Text = ok ? "High Performance power plan set." : "Failed to set High Performance.";
        }
        #endregion

        private void OpenLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dir = Logger.LogDirectory;
                Directory.CreateDirectory(dir);
                Process.Start(new ProcessStartInfo("explorer.exe", dir) { UseShellExecute = true });
            }
            catch { }
        }

        private static string FormatBytes(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1) { order++; len = len / 1024; }
            return $"{len:0.##} {sizes[order]}";
        }
    }

    public static class Logger
    {
        public static string LogDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PulseOptimizer", "Logs");
        private static string LogFile => Path.Combine(LogDirectory, DateTime.Now.ToString("yyyyMMdd") + ".log");

        public static void Info(string message)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(LogFile, $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            }
            catch { /* ignore */ }
        }
    }

    public record DeleteResult(int FilesDeleted, long BytesDeleted);

    public static class Cleaner
    {
        [Flags]
        private enum RecycleFlags : int
        {
            SHERB_NOCONFIRMATION = 0x00000001,
            SHERB_NOPROGRESSUI = 0x00000002,
            SHERB_NOSOUND = 0x00000004
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, RecycleFlags dwFlags);

        public static bool EmptyRecycleBin()
        {
            try
            {
                return SHEmptyRecycleBin(IntPtr.Zero, null, RecycleFlags.SHERB_NOCONFIRMATION | RecycleFlags.SHERB_NOSOUND | RecycleFlags.SHERB_NOPROGRESSUI) == 0;
            }
            catch (Exception ex)
            {
                Logger.Info("EmptyRecycleBin error: " + ex.Message);
                return false;
            }
        }

        public static DeleteResult ClearTemp()
        {
            var total = new DeleteResult(0, 0);
            total = Sum(total, DeleteFrom(Environment.GetEnvironmentVariable("TEMP") ?? ""));
            total = Sum(total, DeleteFrom(Path.GetTempPath()));
            total = Sum(total, DeleteFrom(@"C:\Windows\Prefetch"));
            total = Sum(total, DeleteFrom(@"C:\Windows\Logs"));
            Logger.Info($"ClearTemp: {total.FilesDeleted} files, {total.BytesDeleted} bytes.");
            return total;
        }

        public static DeleteResult ClearWindowsUpdateCache()
        {
            return DeleteFrom(@"C:\Windows\SoftwareDistribution\Download");
        }

        public static DeleteResult ClearBrowserCaches()
        {
            var total = new DeleteResult(0, 0);
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            total = Sum(total, DeleteFrom(Path.Combine(localAppData, @"Google\Chrome\User Data\Default\Cache")));
            total = Sum(total, DeleteFrom(Path.Combine(localAppData, @"Microsoft\Edge\User Data\Default\Cache")));
            var ffProfiles = Path.Combine(appData, @"Mozilla\Firefox\Profiles");
            if (Directory.Exists(ffProfiles))
            {
                foreach (var dir in Directory.GetDirectories(ffProfiles))
                {
                    total = Sum(total, DeleteFrom(Path.Combine(dir, "cache2")));
                }
            }
            Logger.Info($"ClearBrowserCaches: {total.FilesDeleted} files, {total.BytesDeleted} bytes.");
            return total;
        }

        private static DeleteResult DeleteFrom(string path)
        {
            int count = 0;
            long bytes = 0;
            try
            {
                if (Directory.Exists(path))
                {
                    foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
                    {
                        try
                        {
                            var fi = new FileInfo(file);
                            bytes += fi.Length;
                            File.SetAttributes(file, FileAttributes.Normal);
                            fi.Delete();
                            count++;
                        }
                        catch { }
                    }
                    foreach (var dir in Directory.EnumerateDirectories(path, "*", SearchOption.AllDirectories).OrderByDescending(d => d.Length))
                    {
                        try { Directory.Delete(dir, true); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"DeleteFrom error {path}: {ex.Message}");
            }
            return new DeleteResult(count, bytes);
        }

        private static DeleteResult Sum(DeleteResult a, DeleteResult b)
        {
            return new DeleteResult(a.FilesDeleted + b.FilesDeleted, a.BytesDeleted + b.BytesDeleted);
        }
    }

    public class StartupItem
    {
        public string Name { get; set; } = "";
        public string Location { get; set; } = "";
        public string Command { get; set; } = "";
        public bool Enabled { get; set; } = true;
        public string? RegistryRoot { get; set; }
        public string? RegistryPath { get; set; }
        public string? RegistryValueName { get; set; }
        public string? ShortcutPath { get; set; }
    }

    public static class StartupManager
    {
        private const string RunKeyCU = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunKeyLM = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string DisabledKeyCU = @"Software\Microsoft\Windows\CurrentVersion\Run_DisabledByPulse";
        private const string DisabledKeyLM = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run_DisabledByPulse";

        public static List<StartupItem> GetAll()
        {
            var list = new List<StartupItem>();
            list.AddRange(ReadRunKey(Registry.CurrentUser, RunKeyCU, "HKCU", enabled: true));
            list.AddRange(ReadRunKey(Registry.CurrentUser, DisabledKeyCU, "HKCU", enabled: false));
            list.AddRange(ReadRunKey(Registry.LocalMachine, RunKeyLM, "HKLM", enabled: true));
            list.AddRange(ReadRunKey(Registry.LocalMachine, DisabledKeyLM, "HKLM", enabled: false));
            var startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            var disabledFolder = Path.Combine(Path.GetDirectoryName(startupFolder) ?? "", "Startup_DisabledByPulse");
            list.AddRange(ReadStartupFolder(startupFolder, true));
            list.AddRange(ReadStartupFolder(disabledFolder, false));
            return list.OrderBy(i => i.Name).ToList();
        }

        private static IEnumerable<StartupItem> ReadRunKey(RegistryKey root, string subKey, string rootName, bool enabled)
        {
            var items = new List<StartupItem>();
            try
            {
                using var key = root.OpenSubKey(subKey, writable: false);
                if (key != null)
                {
                    foreach (var name in key.GetValueNames())
                    {
                        var cmd = key.GetValue(name)?.ToString() ?? "";
                        items.Add(new StartupItem { Name = name, Location = $"{rootName}\\Run", Command = cmd, Enabled = enabled, RegistryRoot = rootName, RegistryPath = subKey, RegistryValueName = name });
                    }
                }
            }
            catch { }
            return items;
        }

        private static IEnumerable<StartupItem> ReadStartupFolder(string folder, bool enabled)
        {
            var items = new List<StartupItem>();
            try
            {
                if (Directory.Exists(folder))
                {
                    foreach (var lnk in Directory.EnumerateFiles(folder, "*.lnk"))
                    {
                        items.Add(new StartupItem { Name = Path.GetFileNameWithoutExtension(lnk), Location = "StartupFolder", Command = lnk, Enabled = enabled, ShortcutPath = lnk });
                    }
                }
            }
            catch { }
            return items;
        }

        public static bool Disable(StartupItem item)
        {
            try
            {
                if (!item.Enabled) return true;
                if (item.RegistryPath != null && item.RegistryValueName != null && item.RegistryRoot != null)
                {
                    var disabledKey = item.RegistryRoot == "HKCU" ? DisabledKeyCU : DisabledKeyLM;
                    using var root = item.RegistryRoot == "HKCU" ? Registry.CurrentUser : Registry.LocalMachine;
                    using var src = root.OpenSubKey(item.RegistryPath, true);
                    using var dst = root.CreateSubKey(disabledKey, true);
                    if (src != null && dst != null)
                    {
                        var val = src.GetValue(item.RegistryValueName);
                        if (val != null)
                        {
                            dst.SetValue(item.RegistryValueName, val);
                            src.DeleteValue(item.RegistryValueName);
                            return true;
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(item.ShortcutPath))
                {
                    var startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                    var disabledFolder = Path.Combine(Path.GetDirectoryName(startupFolder) ?? "", "Startup_DisabledByPulse");
                    Directory.CreateDirectory(disabledFolder);
                    var dest = Path.Combine(disabledFolder, Path.GetFileName(item.ShortcutPath));
                    File.Move(item.ShortcutPath!, dest, true);
                    return true;
                }
            }
            catch (Exception ex) { Logger.Info("Disable startup error: " + ex.Message); }
            return false;
        }

        public static bool Enable(StartupItem item)
        {
            try
            {
                if (item.Enabled) return true;
                if (item.RegistryPath != null && item.RegistryValueName != null && item.RegistryRoot != null)
                {
                    var disabledKey = item.RegistryRoot == "HKCU" ? DisabledKeyCU : DisabledKeyLM;
                    using var root = item.RegistryRoot == "HKCU" ? Registry.CurrentUser : Registry.LocalMachine;
                    using var src = root.OpenSubKey(disabledKey, true);
                    using var dst = root.CreateSubKey(item.RegistryRoot == "HKCU" ? RunKeyCU : RunKeyLM, true);
                    if (src != null && dst != null)
                    {
                        var val = src.GetValue(item.RegistryValueName);
                        if (val != null)
                        {
                            dst.SetValue(item.RegistryValueName, val);
                            src.DeleteValue(item.RegistryValueName);
                            return true;
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(item.ShortcutPath))
                {
                    var startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                    var disabledFolder = Path.Combine(Path.GetDirectoryName(startupFolder) ?? "", "Startup_DisabledByPulse");
                    var src = Path.Combine(disabledFolder, Path.GetFileName(item.ShortcutPath));
                    var dest = Path.Combine(startupFolder, Path.GetFileName(item.ShortcutPath));
                    if (File.Exists(src))
                    {
                        File.Move(src, dest, true);
                        return true;
                    }
                }
            }
            catch (Exception ex) { Logger.Info("Enable startup error: " + ex.Message); }
            return false;
        }
    }

    public static class Performance
    {
        [DllImport("psapi.dll")]
        static extern bool EmptyWorkingSet(IntPtr hProcess);

        public static long FreeRAM()
        {
            long freed = 0;
            foreach (var p in Process.GetProcesses())
            {
                try
                {
                    if (p.Id == 0) continue;
                    string name = p.ProcessName.ToLowerInvariant();
                    if (name is "system" or "idle" or "csrss" or "wininit" or "winlogon" or "services" or "lsass") continue;
                    var before = p.WorkingSet64;
                    EmptyWorkingSet(p.Handle);
                    p.Refresh();
                    var after = p.WorkingSet64;
                    if (after < before) freed += (before - after);
                }
                catch { /* ignore */ }
            }
            Logger.Info($"FreeRAM estimated: {freed}");
            return freed;
        }

        public static (int killed, int failed) KillCommonBloat()
        {
            string[] targets = new[] { "OneDrive", "Teams", "YourPhone", "Cortana", "SkypeBackground", "AdobeCollabSync", "Discord", "SteamWebHelper" };
            int killed = 0, failed = 0;
            foreach (var p in Process.GetProcesses())
            {
                try
                {
                    if (targets.Any(t => p.ProcessName.Contains(t, StringComparison.OrdinalIgnoreCase)))
                    {
                        p.Kill(true);
                        killed++;
                    }
                }
                catch { failed++; }
            }
            return (killed, failed);
        }
    }

    public static class PowerPlans
    {
        private static bool RunPowerCfg(string args)
        {
            try
            {
                var psi = new ProcessStartInfo("powercfg", args)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                var p = Process.Start(psi);
                p!.WaitForExit(5000);
                Logger.Info($"powercfg {args} => {p.ExitCode}");
                return p.ExitCode == 0;
            }
            catch (Exception ex) { Logger.Info("PowerCfg error: " + ex.Message); return false; }
        }

        public static bool SetHighPerformance()
        {
            return RunPowerCfg("/S 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c");
        }

        public static bool SetBalanced()
        {
            return RunPowerCfg("/S 381b4222-f694-41f0-9685-ff5bb260df2e");
        }
    }
}
