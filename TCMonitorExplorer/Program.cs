using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

class Program
{
    [STAThread]
    static void Main()
    {
        var tcManager = new TotalCommanderManager();
        if (!tcManager.Initialize())
        {
            Console.WriteLine("未找到 Total Commander，程序退出。");
            return;
        }

        var explorerManager = new ExplorerManager(tcManager);
        explorerManager.MainLoop();
    }
}

#region Explorer相关
class ExplorerManager
{
    private readonly HashSet<string> handledFiles = new();
    private readonly Dictionary<int, DateTime> explorerBirth = new();
    private readonly TotalCommanderManager tcManager;
    private dynamic shellApp;

    public ExplorerManager(TotalCommanderManager tcMgr)
    {
        tcManager = tcMgr;
        shellApp = Activator.CreateInstance(
            Type.GetTypeFromProgID("Shell.Application")
        );
    }

    public void MainLoop()
    {
        Console.WriteLine("TC Shell Listener (Hide + Quit Explorer)");
        while (true)
        {
            ScanExplorerWindows();
            Thread.Sleep(100);
        }
    }

    private void ScanExplorerWindows()
    {
        foreach (dynamic window in shellApp.Windows())
        {
            try
            {
                if (window.Name != "文件资源管理器") continue;
                int hwnd = (int)window.HWND;
                if (!explorerBirth.ContainsKey(hwnd))
                {
                    explorerBirth[hwnd] = DateTime.Now;
                    Win32Helper.MoveExplorerOffScreen((IntPtr)hwnd);
                    continue;
                }
                var age = (DateTime.Now - explorerBirth[hwnd]).TotalMilliseconds;
                if (age > 1500)
                {
                    window.Quit();
                    explorerBirth.Remove(hwnd);
                    continue;
                }
                dynamic items = window.Document.SelectedItems();
                if (items == null || items.Count == 0) continue;

                foreach (dynamic item in items)
                {
                    string path = item.Path;
                    if (string.IsNullOrEmpty(path) || handledFiles.Contains(path)) continue;
                    handledFiles.Add(path);
                    tcManager.OpenInTotalCommander(path);
                    Thread.Sleep(200);
                    window.Quit();
                    return;
                }
            }
            catch { /* Explorer初始化失败或已关闭 */ }
        }
    }
}
#endregion

#region TC和Win32管理
class TotalCommanderManager
{
    public string TcPath { get; private set; }
    
    private static readonly string ConfigFilePath = 
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tc_config.txt");

    public bool Initialize()
    {
        TcPath = LoadTcPathFromConfig();
        if (!string.IsNullOrEmpty(TcPath) && File.Exists(TcPath))
        {
            Console.WriteLine("TOTALCMD路径已获取");
            return true;
        }
        
        TcPath = FindTotalCommanderPath();
        if (!string.IsNullOrEmpty(TcPath))
        {
            Console.WriteLine("TOTALCMD路径已保存");
            SaveTcPathToConfig(TcPath);
        }
        
        return !string.IsNullOrEmpty(TcPath);
    }

    private string LoadTcPathFromConfig()
    {
        if (File.Exists(ConfigFilePath))
            return File.ReadAllText(ConfigFilePath).Trim();
        return null;
    }
    private void SaveTcPathToConfig(string path)
    {
        File.WriteAllText(ConfigFilePath, path ?? "");
    }

    public void OpenInTotalCommander(string file)
    {
        var psi = new ProcessStartInfo
        {
            FileName = TcPath,
            Arguments = $"/O /T \"{file}\"",
            UseShellExecute = false
        };
        Process.Start(psi);
    }

    private string FindTotalCommanderPath()
    {
        var procs = Process.GetProcessesByName("TOTALCMD64");
        if (!procs.Any())
            procs = Process.GetProcessesByName("TOTALCMD");

        if (procs.Any())
        {
            try
            {
                return procs[0].MainModule.FileName;
            }
            catch { }
        }

        using var ofd = new OpenFileDialog();
        ofd.Filter = "TOTALCMD|TOTALCMD64.EXE";
        ofd.Title = "请选择 Total Commander 的路径";
        return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
    }
}

static class Win32Helper
{
    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags
    );

    private static readonly IntPtr HWND_BOTTOM = new(1);
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;

    public static void MoveExplorerOffScreen(IntPtr hwnd)
    {
        SetWindowPos(
            hwnd, HWND_BOTTOM, -32000, -32000, 0, 0,
            SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW
        );
    }
}
#endregion