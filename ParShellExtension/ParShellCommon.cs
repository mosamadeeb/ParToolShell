using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParShellExtension
{
    using Options = ParTool.Options;
    using ParTool = ParTool.Program;

    public class ParShellCommon
    {
        public static async Task ParExtract(string parPath, string outPath, bool recursive)
        {
            Options.Extract opt = new Options.Extract
            {
                ParArchivePath = parPath,
                OutputDirectory = outPath,
                Recursive = recursive,
            };

            try
            {
                ParTool.Extract(opt);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static async Task ParRepack(string inPath, string parPath, bool compress)
        {
            Options.Create opt = new Options.Create
            {
                InputDirectory = inPath,
                ParArchivePath = parPath,
                Compression = compress ? 1 : 0,
                AlternativeMode = false,
            };

            try
            {
                ParTool.Create(opt);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static async Task ParAdd(string inPath, List<string> selected, string parPath, bool compress)
        {
            Options.Add opt = new Options.Add
            {
                InputParArchivePath = parPath,
                AddDirectory = inPath,
                OutputParArchivePath = parPath,
                Compression = compress ? 1 : 0,
            };

            try
            {
                ParTool.Add(opt, selected.Where(f => File.Exists(f)).ToList(), selected.Where(f => Directory.Exists(f)).ToList());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static async Task ParCreate(string inPath, List<string> selected, string parPath, bool compress)
        {
            Options.Create opt = new Options.Create
            {
                InputDirectory = inPath,
                ParArchivePath = parPath,
                Compression = compress ? 1 : 0,
                AlternativeMode = false,
            };

            try
            {
                ParTool.Create(opt, selected.Where(f => File.Exists(f)).ToList(), selected.Where(f => Directory.Exists(f)).ToList());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
