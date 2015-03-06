using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing;
using System.ComponentModel;
using System.Management;
using IWshRuntimeLibrary;
using System.Windows.Forms;
using System.Runtime.InteropServices;

//Add References for system.management, system.drawing, and (COM) Windows Script Host Object Model
namespace AppControl
{
    public static class appcontrol
    {

        /// <summary>
        /// Extracts Icon from application
        /// </summary>
        /// <param name="appPath">Path to executable</param>
        /// <returns></returns>
        public static Icon appIcon(string appPath)
        {
            Icon ico = Icon.ExtractAssociatedIcon(appPath);
            return ico;
        }

        /// <summary>
        /// Opens Application and returns handle.
        /// </summary>
        /// <param name="app">Path to executable.</param>
        /// <returns></returns>
        public static IntPtr openapp(string app)
        {
            IntPtr op = IntPtr.Zero;
            try
            {
                Process fileopen = RunCommand.Command(app);
                fileopen.Exited += fileopen_Exited;
                fileopen.Start();
                op = fileopen.MainWindowHandle;
            }
            catch { op = IntPtr.Zero; }
            return op;
        }



        /// <summary>
        /// Code to run when program has exited
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void fileopen_Exited(object sender, EventArgs e)
        {
            //do something
        }

        /// <summary>
        /// Kill a running application and all child windows of that application.
        /// </summary>
        /// <param name="pid">Process ID</param>
        public static void KillProcessAndChildren(int pid)
        {
            //Run WMI call to local machine to get all children of defined parent window.
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_Process Where ParentProcessID=" + pid);
            ManagementObjectCollection moc = searcher.Get();
            foreach (ManagementObject mo in moc)
            {
                //Recursively search for children of children
                KillProcessAndChildren(Convert.ToInt32(mo["ProcessID"]));
            }
            try
            {
                //Kill All Process found.
                Process proc = Process.GetProcessById(pid);
                proc.Kill();
            }
            catch (ArgumentException)
            { 
                
            }
        }

        /// <summary>
        /// Get path to executable from a shortcut
        /// </summary>
        /// <param name="linkPath"></param>
        /// <returns></returns>
        public static string getAppPath(string linkPath)
        {
            string appPath = "";
            if (System.IO.File.Exists(linkPath))
            {
                WshShell WShell = new WshShell();
                IWshShortcut Link = (IWshShortcut)WShell.CreateShortcut(linkPath);

                appPath = Link.TargetPath;
            }
            return appPath;
        }


        /// <summary>
        /// Get the name of a file from the path string
        /// </summary>
        /// <param name="appPath">Path to file</param>
        /// <returns></returns>
        public static string getAppName(string appPath)
        {
            string[] spl = { "\\" };
            string[] splb = { "." };
            string[] splpath = appPath.Split(spl, StringSplitOptions.None);
            string[] ext = splpath.Last<string>().Split(splb, StringSplitOptions.None);
            string Name = splpath.Last<string>().Replace("." + ext.Last<string>(), "");
            return Name;
        }
        

        /// <summary>
        /// Returns a list of all currently running tasks (Main Windows Title)
        /// </summary>
        /// <returns></returns>
        public static List<string> GetActiveTasks()
        {
            List<string> ar = new List<string>();
            IntPtr child = IntPtr.Zero;

            Process[] process = Process.GetProcesses();
            foreach (Process p in process)
            {
                //WindowData w;
                if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle.Length > 0)
                {
                    ar.Add(p.MainWindowTitle);
                    
                }
            }
            return ar;
        }

        /// <summary>
        /// Returns a list of Processes that are currently running
        /// </summary>
        /// <returns></returns>
        public static List<Process> GetActiveTasksAll()
        {
            List<Process> ar = new List<Process>();
            IntPtr child = IntPtr.Zero;

            Process[] process = Process.GetProcesses();
            foreach (Process p in process)
            {
                //WindowData w;
                if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle.Length > 0)
                {
                    
                    ar.Add(p);

                }
            }
            return ar;
        }

        /// <summary>
        /// Returns a list of the Process names of all running processes
        /// </summary>
        /// <returns></returns>
        public static List<string> GetActiveTasksNames()
        {
            List<string> ar = new List<string>();
            IntPtr child = IntPtr.Zero;

            Process[] process = Process.GetProcesses();
            foreach (Process p in process)
            {
                //WindowData w;
                if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle.Length > 0)
                {
                    ar.Add(p.ProcessName);

                }
            }
            return ar;
        }

        /// <summary>
        /// Returns a windows title based on the handle of a process
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static string GetWindowTitle(IntPtr hwnd)
        {
            string ar = "";
          

            Process[] process = Process.GetProcesses();
            foreach (Process p in process)
            {
                //WindowData w;
                if (p.MainWindowHandle == hwnd)
                {
                    ar = p.ProcessName;

                }
            }
            return ar;
        }


        /// <summary>
        /// Kills an application and its children with user prompts.
        /// </summary>
        /// <param name="App">Process to kill</param>
        public static void AppKill(Process App)
        {
            //Minimize application before processing.
            appMin(App);

            //Verify that user wants to kill app
            DialogResult res = TopMostMessageBox.Show("Would you like to terminate " + App.ProcessName + "\nThis may result in lost work.", "WARNING", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    //kill App.
                    KillProcessAndChildren(App.Id);
                }
                catch
                {
                    MessageBox.Show(App.ProcessName + " could not be killed.");
                }
            }
            if (res == DialogResult.Cancel)
            {
                //Restore App
                appRestore(App);
            }



        }


        /// <summary>
        /// Minimize Application
        /// </summary>
        /// <param name="App">Process to minimize</param>
        public static void appMin(Process App)
        {
            
            WinAPI.ShowWindowAsync(App.MainWindowHandle, windowState.SW_MINIMIZE);
        }
        /// <summary>
        /// Maximize Application
        /// </summary>
        /// <param name="App">Process to Maximize</param>
        public static void appMax(Process App)
        {
            WinAPI.ShowWindowAsync(App.MainWindowHandle, windowState.SW_SHOWMAXIMIZED);
        }

        /// <summary>
        /// Restore Windows size of Application
        /// </summary>
        /// <param name="App">Process to restore</param>
        public static void appRestore(Process App)
        {
            WinAPI.ShowWindowAsync(App.MainWindowHandle, windowState.SW_RESTORE);
        }

        /// <summary>
        /// Minimize all applications except one.
        /// </summary>
        /// <param name="App">Process to not minimize</param>
        public static void appMinAll(Process App)
        {
            foreach (Process blip in appcontrol.GetActiveTasksAll())
            {
                if (blip != App)
                {
                    appcontrol.appMin(blip);
                }
            }
        }

        /// <summary>
        /// Minimize all applications
        /// </summary>
        public static void appMinAll()
        {
            foreach (Process blip in appcontrol.GetActiveTasksAll())
            {
                appcontrol.appMin(blip);
            }
        }

        /// <summary>
        /// Shows an open window using the definition provided in windowState class.
        /// </summary>
        /// <param name="hWnd">A handle to the window. </param>
        /// <param name="nCmdShow">Controls how the window is to be shown.</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// This class contais the defeinitions of the windowState variables used for ShowWindow and ShowWindowAsync
        /// </summary>
        public static class windowState
        {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            public const int SW_HIDE = 0;

            /// <summary>
            /// Activates and displays a window. If the window is minimized or maximized,
            ///  the system restores it to its original size and position. An application
            ///  should specify this flag when displaying the window for the first time.
            /// </summary>
            public const int SW_SHOWNORMAL = 1;

            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            public const int SW_SHOWMINIMIZED = 2;

            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>
            public const int SW_SHOWMAXIMIZED = 3;

            /// <summary>
            /// Displays a window in its most recent size and position. This value is similar
            ///  to SW_SHOWNORMAL, except that the window is not activated.
            /// </summary>
            public const int SW_SHOWNOACTIVATE = 4;

            /// <summary>
            /// Activates the window and displays it in its current size and position. 
            /// </summary>
            public const int SW_SHOW = 5;

            /// <summary>
            /// Minimizes the specified window and activates the next top-level window in the Z order.
            /// </summary>
            public const int SW_MINIMIZE = 6;

            /// <summary>
            /// Displays the window as a minimized window. This value is similar to
            ///  SW_SHOWMINIMIZED, except the window is not activated.
            /// </summary>
            public const int SW_SHOWMINNOACTIVE = 7;

            /// <summary>
            /// Displays the window in its current size and position. This value is 
            /// similar to SW_SHOW, except that the window is not activated.
            /// </summary>
            public const int SW_SHOWNA = 8;

            /// <summary>
            /// Activates and displays the window. If the window is minimized or
            ///  maximized, the system restores it to its original size and position. An
            ///  application should specify this flag when restoring a minimized window.
            /// </summary>
            public const int SW_RESTORE = 9;

            /// <summary>
            /// Sets the show state based on the SW_ value specified in the STARTUPINFO
            ///  structure passed to the CreateProcess function by the program that started
            ///  the application. 
            /// </summary>
            public const int SW_SHOWDEFAULT = 10;

            /// <summary>
            /// Minimizes a window, even if the thread that owns the window is not
            ///  responding. This flag should only be used when minimizing windows
            ///  from a different thread.
            /// </summary>
            public const int SW_FORCEMINIMIZE = 11;

        }

    }




}
