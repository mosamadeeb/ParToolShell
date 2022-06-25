using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

using static ParShellExtension.ParShellCommon;

namespace ParShellExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class ParShellContextMenu : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return SelectedItemPaths.Any();
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem parToolItem = new ToolStripMenuItem("ParTool", MenuImage);

            List<string> selected = SelectedItemPaths.ToList();

            if (selected.Count == 1 && File.Exists(selected[0]) && Path.GetExtension(selected[0]).ToLowerInvariant().StartsWith(".par"))
            {
                string outPath = selected[0] + ".unpack";
                string extractText = string.Format("Extract to \"{0}\\\"", Path.GetFileName(outPath));

                if (Directory.Exists(outPath))
                {
                    extractText = "[Overwrite] " + extractText;
                }

                ToolStripMenuItem extractItem = new ToolStripMenuItem(extractText);
                extractItem.Click += (sender, args) => Task.Run(() => ParExtract(selected[0], outPath, false));
                parToolItem.DropDownItems.Add(extractItem);

                ToolStripMenuItem extractItemRecursive = new ToolStripMenuItem(extractText + " recursively");
                extractItemRecursive.Click += (sender, args) => Task.Run(() => ParExtract(selected[0], outPath, true));
                parToolItem.DropDownItems.Add(extractItemRecursive);
            }

            if (selected.Count == 1
                && Directory.Exists(selected[0]) && selected[0].ToLowerInvariant().EndsWith(".unpack")
                && Path.GetExtension(Path.GetFileNameWithoutExtension(selected[0])).ToLowerInvariant().StartsWith(".par"))
            {
                string parRepackPath = selected[0].Substring(0, selected[0].Length - ".unpack".Length);
                string repackText = string.Format("Repack into \"{0}\"", Path.GetFileName(parRepackPath));

                if (File.Exists(parRepackPath))
                {
                    repackText = "[Overwrite] " + repackText;
                }

                ToolStripMenuItem repackItem = new ToolStripMenuItem(repackText);
                repackItem.Click += (sender, args) => Task.Run(() => ParRepack(selected[0], parRepackPath, false));
                parToolItem.DropDownItems.Add(repackItem);

                ToolStripMenuItem repackItemCompression = new ToolStripMenuItem(repackText + " with compression");
                repackItemCompression.Click += (sender, args) => Task.Run(() => ParRepack(selected[0], parRepackPath, true));
                parToolItem.DropDownItems.Add(repackItemCompression);
            }

            // We know that we have at least 1 item selected
            string folderPath = Path.GetDirectoryName(selected[0]);
            string parPath = Path.Combine(folderPath, Path.GetFileName(folderPath));

            if (Path.GetExtension(parPath) == ".unpack"
                && Path.GetExtension(Path.GetFileNameWithoutExtension(parPath)).ToLowerInvariant().StartsWith(".par"))
            {
                parPath = parPath.Substring(0, parPath.Length - ".unpack".Length);
            }
            else
            {
                parPath += ".par";
            }

            string addText = string.Format("Add to \"{0}\"", Path.GetFileName(parPath));

            bool parExists = File.Exists(parPath);
            if (!parExists)
            {
                addText = "Create and " + addText;
            }

            ToolStripMenuItem addItem = new ToolStripMenuItem(addText);
            ToolStripMenuItem addItemCompression = new ToolStripMenuItem(addText + " with compression");

            if (parExists)
            {
                addItem.Click += (sender, args) => Task.Run(() => ParAdd(folderPath, selected, parPath, false));
                addItemCompression.Click += (sender, args) => Task.Run(() => ParAdd(folderPath, selected, parPath, true));
            }
            else
            {
                addItem.Click += (sender, args) => Task.Run(() => ParCreate(folderPath, selected, parPath, false));
                addItemCompression.Click += (sender, args) => Task.Run(() => ParCreate(folderPath, selected, parPath, true));
            }

            if (selected.Contains(parPath))
            {
                // Prevent using ParTool Add with the output par being selected
                parToolItem.DropDownItems.Add(string.Format("Cannot add to \"{0}\" with the selected files!", Path.GetFileName(parPath)));
            }
            else
            {
                parToolItem.DropDownItems.Add(addItem);
                parToolItem.DropDownItems.Add(addItemCompression);
            }

            ToolStripSeparator separator = new ToolStripSeparator();
            parToolItem.DropDownItems.Add(separator);

            ToolStripMenuItem showConsoleItem = new ToolStripMenuItem("Show ParTool Console");
            showConsoleItem.Checked = ShowConsole;
            showConsoleItem.Click += (sender, args) => ShowConsole = !ShowConsole;

            parToolItem.DropDownItems.Add(showConsoleItem);

            ToolStripMenuItem alternativeModeItem = new ToolStripMenuItem("Use Alternative Mode");
            alternativeModeItem.Checked = AlternativeMode;
            alternativeModeItem.Click += (sender, args) => AlternativeMode = !AlternativeMode;

            parToolItem.DropDownItems.Add(alternativeModeItem);

            menu.Items.Add(parToolItem);
            return menu;
        }
    }
}
