using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using SHDocVw;
using System.Threading;
using System.Runtime.InteropServices;
using Win32Interop.WinHandles;

//Icon source: https://publicdomainvectors.org/en/free-clipart/White-cross-within-a-red-octagon-vector-image/17388.html

namespace Alice_s_Window_Closer_App
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int PostMessage(int hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private const int WM_Close = 16;
        private static bool force;

        private Program()
        {
            //Resolves S1118:Utility classes should not have public constructors
        }

        private static void Main(string[] args)
        {
            if (args.Length == 1)
            {
                switch (args[0])
                {
                    case "--force":
                        force = true;
                        break;

                    default:
                        Console.Out.WriteLine("Alice's Window Closer App v1.0");
                        Console.Out.WriteLine("--force  Kill any processes that have an open window. (Dangerous!)");
                        Console.Out.WriteLine("--help   Show help flags");
                        break;
                }
            }

            List<Process> taskBarProcesses = Process.GetProcesses().Where(p => !string.IsNullOrEmpty(p.MainWindowTitle)).Where(p => p.MainWindowHandle != GetConsoleWindow()).ToList();

            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();

            foreach (InternetExplorer item in shellWindows)
            {
                item.Quit();
            }

            taskBarProcesses.ForEach((task) =>
            {
#if DEBUG
                if (task.ProcessName == "devenv")
                {
                    return;
                }
#endif

                TopLevelWindowUtils.FindWindows((wh) => IsWindowHandleOwnedBy(wh, task)).ToList()
                .ForEach((wh) =>
                {
                    Console.Out.WriteLine($"Closing {task.ProcessName}");
                    PostMessage(wh.RawPtr.ToInt32(), WM_Close, 0, 0);
                });

                if (force)
                {
                    Thread.Sleep(500);
                    Console.Out.WriteLine($"Killing {task.ProcessName}");
                    try
                    {
                        task.Refresh();
                        if (!task.HasExited) { task.Kill(); }
                    }
                    catch (InvalidOperationException e)
                    {
                        Console.Error.WriteLine($"Something went wrong trying to kill {task.ProcessName}: {e.Message}");
                    }
                }
            });

#if DEBUG
            Console.ReadKey();
#endif
        }

        private static bool IsWindowHandleOwnedBy(WindowHandle wh, Process task) => (
            wh.IsValid && wh.IsVisible() &&
            (GetWindowThreadProcessId(wh.RawPtr, out uint _) == GetWindowThreadProcessId(task.MainWindowHandle, out uint _))
            );
    }
}