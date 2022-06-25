using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell.Diagnostics;

namespace ParShellExtension
{
    public class ParShellCommon
    {
        private static string parToolPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            Assembly.GetExecutingAssembly().GetName().Name);

        public static bool ShowConsole = false;

        public static bool AlternativeMode = false;

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

            if (AlternativeMode)
            {
                args += " --alternative-mode";
            }

            RunParTool(args);
        }

        public static async Task ParAdd(string inPath, IEnumerable<string> selected, string parPath, bool compress)
        {
            string args = $"add \"{Path.GetFullPath(parPath)}\" \"{Path.GetFullPath(inPath)}\" \"{Path.GetFullPath(parPath)}\"";

            args += GetItemsArg(selected, "--files", File.Exists);
            args += GetItemsArg(selected, "--folders", Directory.Exists);

            // Default is compression
            if (!compress)
            {
                args += " -c 0";
            }

            if (AlternativeMode)
            {
                args += " --alternative-mode";
            }

            RunParTool(args);
        }

        public static async Task ParCreate(string inPath, IEnumerable<string> selected, string parPath, bool compress)
        {
            string args = $"create \"{Path.GetFullPath(inPath)}\" \"{Path.GetFullPath(parPath)}\"";

            args += GetItemsArg(selected, "--files", File.Exists);
            args += GetItemsArg(selected, "--folders", Directory.Exists);

            // Default is compression
            if (!compress)
            {
                args += " -c 0";
            }

            if (AlternativeMode)
            {
                args += " --alternative-mode";
            }

            RunParTool(args);
        }

        private static async Task RunParTool(string args)
        {
            Logging.Log($"Running ParTool with arguments: {args}");

            try
            {
                using Process process = new Process();

                Task parToolTask = Task.Run(() => RunParToolProcess(process, args));

                if (!ShowConsole && !parToolTask.Wait(2000))
                {
                    string caption = "ParTool Shell Running";

                    Task<DialogResult> abortMessageTask = Task.Run(() => ShowMessageBox("ParTool is running...", caption, MessageBoxButtons.OK, MessageBoxIcon.Information, "Abort"));

                    int index = Task.WaitAny(abortMessageTask, parToolTask);

                    if (index == 0 && abortMessageTask.Result == DialogResult.OK)
                    {
                        process.Kill();
                        MessageBox.Show("Process aborted.", "ParTool Shell", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        await parToolTask.ConfigureAwait(false);

                        if (!abortMessageTask.IsCompleted)
                        {
                            CloseMessageBox(caption);
                        }

                        MessageBox.Show("Finished!", "ParTool Shell", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    await parToolTask.ConfigureAwait(false);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ParTool Shell Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static async Task RunParToolProcess(Process process, string args)
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

        private static string GetItemsArg(IEnumerable<string> items, string arg, Func<string, bool> existsFunc)
        {
            string itemsArg = string.Empty;

            foreach (string item in items.Where(existsFunc))
            {
                itemsArg += $" \"{Path.GetFileName(item)}\"";
            }

            return $" {arg}{(itemsArg.Length == 0 ? " \"\"" : itemsArg)}";
        }

        /// <summary>
        /// This method should be ran in a new thread as a Task.
        /// </summary>
        private static DialogResult ShowMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, params string[] buttonsText)
        {
            DialogResult result = DialogResult.None;

            try
            {
                switch (buttons)
                {
                    case MessageBoxButtons.OK:
                        MessageBoxManager.OK = buttonsText[0];
                        break;
                    default:
                        break;
                }

                MessageBoxManager.Register();
                result = MessageBox.Show(text, caption, buttons, icon);
                MessageBoxManager.Unregister();
            }
            catch
            {

            }

            return result;
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
