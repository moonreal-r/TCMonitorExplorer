using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

class Program
{
    // 已处理的文件，防止重复触发
    static readonly HashSet<string> HandledFiles = new();

    // Explorer 窗口创建时间记录（HWND -> 时间）
    static readonly Dictionary<int, DateTime> ExplorerBirth = new();

    // 已隐藏的 Explorer，避免重复 Hide
    static readonly HashSet<int> HiddenExplorers = new();

    // Shell.Application 只创建一次
    static dynamic ShellApp;

    // Win32
    [DllImport("user32.dll")]
    static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags
    );
    
    static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

    const uint SWP_NOSIZE = 0x0001;
    const uint SWP_NOZORDER = 0x0004;
    const uint SWP_NOACTIVATE = 0x0010;
    const uint SWP_SHOWWINDOW = 0x0040;
    
    static void MoveExplorerOffScreen(IntPtr hwnd)
    {
        SetWindowPos(
            hwnd,
            HWND_BOTTOM,
            -32000,   // ⭐ 系统级不可见坐标
            -32000,
            0,
            0,
            SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW
        );
    }


    static void Main()
    {
        Console.WriteLine("TC Shell Listener (Hide + Quit Explorer)");

        ShellApp = Activator.CreateInstance(
            Type.GetTypeFromProgID("Shell.Application")
        );

        while (true)
        {
            ScanExplorerWindows();
            Thread.Sleep(100); // 已验证的稳定值
        }
    }

    static void ScanExplorerWindows()
    {
        foreach (dynamic window in ShellApp.Windows())
        {
            try
            {
                if (window.Name != "文件资源管理器")
                    continue;

                int hwnd = (int)window.HWND;

                // 第一次看到这个 Explorer
                if (!ExplorerBirth.ContainsKey(hwnd))
                {
                    ExplorerBirth[hwnd] = DateTime.Now;

                    // ⭐ 直接移出屏幕
                    MoveExplorerOffScreen((IntPtr)hwnd);
                    
                    continue;
                }

                var age = (DateTime.Now - ExplorerBirth[hwnd]).TotalMilliseconds;

                // 超过 1500ms 还没选中任何文件 → 直接关闭
                if (age > 1500)
                {
                    window.Quit();
                    ExplorerBirth.Remove(hwnd);
                    continue;
                }
                
                dynamic items = window.Document.SelectedItems();
                if (items == null || items.Count == 0)
                    continue;

                foreach (dynamic item in items)
                {
                    string path = item.Path;
                    if (string.IsNullOrEmpty(path))
                        continue;

                    if (HandledFiles.Contains(path))
                        continue;

                    HandledFiles.Add(path);

                    // 打开 TC
                    OpenInTotalCommander(path);

                    // ⭐ 给 TC 一点时间接管
                    Thread.Sleep(200);

                    // ⭐ 后台关闭 Explorer（用户看不到）
                    window.Quit();
                    return;
                }
            }
            catch
            {
                // Explorer 初始化 / 已关闭，忽略
            }
        }
    }

    static void OpenInTotalCommander(string file)
    {
        var psi = new ProcessStartInfo
        {
            FileName = @"F:\Program Files\TotalCMD64\TOTALCMD64.EXE",
            Arguments = $"/O /T \"{file}\"",
            UseShellExecute = false
        };

        Process.Start(psi);
    }
}
