using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ItemsReport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (ApplicationRunningHelper.AlreadyRunning())
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            AppDomain.CurrentDomain.AssemblyResolve += (Object sender, ResolveEventArgs args) =>
            {
                String thisExe = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                System.Reflection.AssemblyName embeddedAssembly = new System.Reflection.AssemblyName(args.Name);
                String resourceName = ""; // = thisExe + "." + embeddedAssembly.Name + ".dll";
                if (embeddedAssembly.Name.Contains("Bunifu"))
                {
                    resourceName = "WindowsFormsApplication1.Resources.Bunifu_UI_v1.5.3.dll";
                }
                else if (embeddedAssembly.Name.Contains("EPPlus"))
                {
                    resourceName = "WindowsFormsApplication1.Resources.EPPlus.dll";
                }

                using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    Byte[] assemblyData = new Byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return System.Reflection.Assembly.Load(assemblyData);
                }
            };



            Application.Run(new ItemsReport());
        }

        public static class ApplicationRunningHelper
        {
            [DllImport("user32.dll")]
            private static extern
                bool SetForegroundWindow(IntPtr hWnd);
            [DllImport("user32.dll")]
            private static extern
                bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
            [DllImport("user32.dll")]
            private static extern
                bool IsIconic(IntPtr hWnd);

            /// -------------------------------------------------------------------------------------------------
            /// <summary> check if current process already running. if running, set focus to existing process and 
            ///           returns <see langword="true"/> otherwise returns <see langword="false"/>. </summary>
            /// <returns> <see langword="true"/> if it succeeds, <see langword="false"/> if it fails. </returns>
            /// -------------------------------------------------------------------------------------------------
            public static bool AlreadyRunning()
            {
                /*
                const int SW_HIDE = 0;
                const int SW_SHOWNORMAL = 1;
                const int SW_SHOWMINIMIZED = 2;
                const int SW_SHOWMAXIMIZED = 3;
                const int SW_SHOWNOACTIVATE = 4;
                const int SW_RESTORE = 9;
                const int SW_SHOWDEFAULT = 10;
                */
                const int swRestore = 9;

                var me = Process.GetCurrentProcess();
                var arrProcesses = Process.GetProcessesByName(me.ProcessName);

                if (arrProcesses.Length > 1)
                {
                    MessageBox.Show("App is already running", "Message");

                    for (var i = 0; i < arrProcesses.Length; i++)
                    {
                        if (arrProcesses[i].Id != me.Id)
                        {
                            // get the window handle
                            IntPtr hWnd = arrProcesses[i].MainWindowHandle;

                            // if iconic, we need to restore the window
                            if (IsIconic(hWnd))
                            {
                                ShowWindowAsync(hWnd, swRestore);
                            }

                            // bring it to the foreground
                            SetForegroundWindow(hWnd);
                            break;
                        }
                    }
                    return true;
                }

                return false;
            }
        }
    }
}
