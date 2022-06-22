using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ParShellExt
{
    using Options = ParTool.Options;
    using ParTool = ParTool.Program;

    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class ParContextMenu : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return SelectedItemPaths.Any();
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem parToolItem = new ToolStripMenuItem("ParTool");

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
                extractItem.Click += (sender, args) => ParExtract(selected[0], outPath, false);
                parToolItem.DropDownItems.Add(extractItem);

                ToolStripMenuItem extractItemRecursive = new ToolStripMenuItem(extractText + " recursively");
                extractItemRecursive.Click += (sender, args) => ParExtract(selected[0], outPath, true);
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
                repackItem.Click += (sender, args) => ParRepack(selected[0], parRepackPath, false);
                parToolItem.DropDownItems.Add(repackItem);

                ToolStripMenuItem repackItemCompression = new ToolStripMenuItem(repackText + " with compression");
                repackItemCompression.Click += (sender, args) => ParRepack(selected[0], parRepackPath, true);
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
                addItem.Click += (sender, args) => ParAdd(folderPath, selected, parPath, false);
                addItemCompression.Click += (sender, args) => ParAdd(folderPath, selected, parPath, true);
            }
            else
            {
                addItem.Click += (sender, args) => ParCreate(folderPath, selected, parPath, false);
                addItemCompression.Click += (sender, args) => ParCreate(folderPath, selected, parPath, true);
            }

            parToolItem.DropDownItems.Add(addItem);
            parToolItem.DropDownItems.Add(addItemCompression);

            menu.Items.Add(parToolItem);
            return menu;
        }

        private void ParExtract(string parPath, string outPath, bool recursive)
        {
            Options.Extract opt = new Options.Extract
            {
                ParArchivePath = parPath,
                OutputDirectory = outPath,
                Recursive = recursive,
            };

            ParTool.Extract(opt);
        }

        private void ParRepack(string inPath, string parPath, bool compress)
        {
            Options.Create opt = new Options.Create
            {
                InputDirectory = inPath,
                ParArchivePath = parPath,
                Compression = compress ? 1 : 0,
                AlternativeMode = false,
            };

            ParTool.Create(opt);
        }

        private void ParAdd(string inPath, List<string> selected, string parPath, bool compress)
        {
            Options.Add opt = new Options.Add
            {
                InputParArchivePath = parPath,
                AddDirectory = inPath,
                OutputParArchivePath = parPath,
                Compression = compress ? 1 : 0,
            };

            ParTool.Add(opt, selected.Where(f => File.Exists(f)).ToList(), selected.Where(f => Directory.Exists(f)).ToList());
        }

        private void ParCreate(string inPath, List<string> selected, string parPath, bool compress)
        {
            Options.Create opt = new Options.Create
            {
                InputDirectory = inPath,
                ParArchivePath = parPath,
                Compression = compress ? 1 : 0,
                AlternativeMode = false,
            };

            ParTool.Create(opt, selected.Where(f => File.Exists(f)).ToList(), selected.Where(f => Directory.Exists(f)).ToList());
        }
    }
}
