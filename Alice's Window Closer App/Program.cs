using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using SHDocVw;
using System.Threading;
using System.Runtime.InteropServices;

//Icon source: https://publicdomainvectors.org/en/free-clipart/White-cross-within-a-red-octagon-vector-image/17388.html


namespace Alice_s_Window_Closer_App
{

    internal class Program
    {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        private static bool force = false;

        private Program()
        {
            //Resolves S1118:Utility classes should not have public constructors
        }

        static void Main(string[] args)
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
                Console.Out.WriteLine($"Closing {task.ProcessName}");

                try
                {
                    task.CloseMainWindow();
                }
                catch (InvalidOperationException e)
                {
                    Console.Error.WriteLine($"Something went wrong trying to close {task.ProcessName}: {e.Message}");
                }
                if (force)
                {
                    Thread.Sleep(500);
                    Console.Out.WriteLine($"Killing {task.ProcessName}");
                    try
                    {
                        if (task.HasExited) { task.Kill(); }
                    }
                    catch (InvalidOperationException e)
                    {
                        Console.Error.WriteLine($"Something went wrong trying to kill {task.ProcessName}: {e.Message}");
                    }

                }

            });



        }
    }
}
