using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParShellExtension
{
    public class ParShellCommon
    {
        private static string parToolPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            Assembly.GetExecutingAssembly().GetName().Name);

        public static bool ShowConsole = false;

        public static async Task ParExtract(string parPath, string outPath, bool recursive)
        {
            string args = $"extract \"{Path.GetFullPath(parPath)}\" \"{Path.GetFullPath(outPath)}\"";

            if (recursive)
            {
                args += " -r";
            }

            RunParTool(args);
        }

        public static async Task ParRepack(string inPath, string parPath, bool compress)
        {
            string args = $"create \"{Path.GetFullPath(inPath)}\" \"{Path.GetFullPath(parPath)}\"";

            // Default is compression
            if (!compress)
            {
                args += " -c 0";
            }

            RunParTool(args);
        }

        public static async Task ParAdd(string inPath, List<string> selected, string parPath, bool compress)
        {
            string args = $"add \"{Path.GetFullPath(parPath)}\" \"{Path.GetFullPath(inPath)}\" \"{Path.GetFullPath(parPath)}\"";

            args += GetItemsArg(selected, "--files", File.Exists);
            args += GetItemsArg(selected, "--folders", Directory.Exists);

            // Default is compression
            if (!compress)
            {
                args += " -c 0";
            }

            RunParTool(args);
        }

        public static async Task ParCreate(string inPath, List<string> selected, string parPath, bool compress)
        {
            string args = $"create \"{Path.GetFullPath(inPath)}\" \"{Path.GetFullPath(parPath)}\"";

            args += GetItemsArg(selected, "--files", File.Exists);
            args += GetItemsArg(selected, "--folders", Directory.Exists);

            // Default is compression
            if (!compress)
            {
                args += " -c 0";
            }

            RunParTool(args);
        }

        private static async Task RunParTool(string args)
        {
            Task task = Task.Run(() => RunParToolProcess(args));
            if (!ShowConsole && !task.Wait(2500))
            {
                string caption = "ParTool Shell";
                Task.Run(() => MessageBox.Show("ParTool is running...", caption, MessageBoxButtons.OK, MessageBoxIcon.Information));

                await task.ConfigureAwait(false);
                CloseMessageBox(caption);

                MessageBox.Show("Finished!", caption + " Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static async Task RunParToolProcess(string args)
        {
            Process process = new Process();

            try
            {
                if (!ShowConsole)
                {
                    // Stop the process from opening a new window
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                }

                process.StartInfo.UseShellExecute = true;

                // Setup executable and parameters
                process.StartInfo.WorkingDirectory = parToolPath;
                process.StartInfo.FileName = "ParTool.exe";
                process.StartInfo.Arguments = args;

                // Go
                process.Start();
                process.WaitForExit();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ParTool Shell Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                process.Dispose();
            }
        }

        private static string GetItemsArg(IEnumerable<string> items, string arg, Func<string, bool> existsFunc)
        {
            string itemsArg = string.Empty;

            foreach (string item in items.Where(existsFunc))
            {
                itemsArg += $" \"{Path.GetFileName(item)}\"";
            }

            return (itemsArg.Length == 0) ? string.Empty : $" {arg} {itemsArg}";
        }

        private static void CloseMessageBox(string caption)
        {
            // lpClassName is #32770 for MessageBox
            IntPtr mbWnd = FindWindow("#32770", caption);

            if (mbWnd != IntPtr.Zero)
            {
                SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private const int WM_CLOSE = 0x0010;
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    }
}
