using ParShellExtension.NamespaceItems;
using SharpShell.Attributes;
using SharpShell.Diagnostics;
using SharpShell.Interop;
using SharpShell.Pidl;
using SharpShell.SharpNamespaceExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Yarhl.FileSystem;
using ParLibrary.Converter;

namespace ParShellExtension
{
    [ComVisible(true)]
    [NamespaceExtensionJunctionPoint(NamespaceExtensionAvailability.Everyone, VirtualFolder.MyComputer, "ParShellNamespace")]
    public class ParShellNamespace : SharpNamespaceExtension
    {
        //private const string Kernel32Dll = "kernel32.dll";

        //[DllImport(Kernel32Dll)]
        //private static extern string GetCommandLineA();

        private string parPath = @"C:\Users\Admin\Desktop\a14_0050.par.unpack\a14_0050.par";

        private Node CurNode;

        private List<ParItem> Items = null;
        private List<ParFolder> Folders = null;

        private bool createdLists = false;
        private bool createdChildren = false;

        public ParShellNamespace()
        {
            Logging.Log("Creating an instance of the namespace");

            Logging.Log("Loading par from file!!!!");

            //curPar = NodeFactory.FromFile(parPath, Yarhl.IO.FileOpenMode.Read);

            MemoryStream stream = new MemoryStream();
            using (FileStream file = File.OpenRead(parPath))
            {
                file.CopyTo(stream);
                stream.Seek(0, SeekOrigin.Begin);
            }

            CurNode = NodeFactory.FromSubstream(Path.GetFileName(parPath), stream, 0, stream.Length);
            CurNode.TransformWith<ParArchiveReader, ParArchiveReaderParameters>(new ParArchiveReaderParameters() { Recursive = true });

            Logging.Log("Finished loading par!!!");

            //rootPath = GetCommandLineA();
            //File.WriteAllText(@"C:\Users\Admin\Desktop\a14_0050.par.unpack\TESTPATH.txt", rootPath);
        }

        public override NamespaceExtensionRegistrationSettings GetRegistrationSettings()
        {
            return new NamespaceExtensionRegistrationSettings
            {
                ExtensionAttributes = AttributeFlags.IsFolder
                                    //| AttributeFlags.IsBrowsable
                                    //| AttributeFlags.IsDropTarget
                                    | AttributeFlags.MayContainSubFolders,
                HideFolderVerbs = true,
                UseAssemblyIcon = false,
                Tooltip = "Test ParShellNamespace",
            };
        }

        protected override IEnumerable<IShellNamespaceItem> GetChildren(ShellNamespaceEnumerationFlags flags)
        {
            Logging.Log($"Enumerating children for the namespace");

            /*
            if (CurNode is null)
            {
                Logging.Log($"CurNode is null!!!");
                Logging.Log("Loading par from file!!!!");

                //curPar = NodeFactory.FromFile(parPath, Yarhl.IO.FileOpenMode.Read);

                MemoryStream stream = new MemoryStream();
                using (FileStream file = File.OpenRead(parPath))
                {
                    file.CopyTo(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                }

                CurNode = NodeFactory.FromSubstream(Path.GetFileName(parPath), stream, 0, stream.Length);
                CurNode.TransformWith<ParArchiveReader, ParArchiveReaderParameters>(new ParArchiveReaderParameters() { Recursive = true });

                Logging.Log("Finished loading par!!!");
            }
            */

            /*
            if (CurNode is null)
            {
                Logging.Log($"CurNode is null!!!");
                yield break;
            }
            */

            if (!createdChildren)
            {
                if (CurNode.Children.Count == 0)
                {
                    Logging.Log($"no children!!!");
                    yield break;
                }

                if (!createdLists)
                {
                    Logging.Log($"creating lists");
                    Items = new List<ParItem>();
                    Folders = new List<ParFolder>();
                    createdLists = true;
                }

                ParItem item;
                ParFolder folder;

                foreach (Node c in CurNode.Children[0].Children)
                {
                    Logging.Log($"child");
                    if (c.IsContainer)
                    {
                        folder = (c.Name == "." && c.Children.Count == 1 && c.Children[0].IsContainer) ? new ParFolder(c.Children[0]) : new ParFolder(c);
                        Folders.Add(folder);

                        if (flags.HasFlag(ShellNamespaceEnumerationFlags.Folders))
                        {
                            yield return folder;
                        }
                    }
                    else
                    {
                        item = new ParItem(c);
                        Items.Add(item);

                        if (flags.HasFlag(ShellNamespaceEnumerationFlags.Items))
                        {
                            yield return item;
                        }
                    }
                }

                Logging.Log($"finished creating children");
                createdChildren = true;
                yield break;
            }

            Logging.Log($"enumerating now");

            if (flags.HasFlag(ShellNamespaceEnumerationFlags.Folders))
            {
                foreach (ParFolder folder in Folders)
                {
                    yield return folder;
                }
            }

            if (flags.HasFlag(ShellNamespaceEnumerationFlags.Items))
            {
                foreach (ParItem item in Items)
                {
                    yield return item;
                }
            }

            //return Enumerable.Empty<IShellNamespaceItem>();
        }

        protected override ShellNamespaceFolderView GetView()
        {
            Logging.Log("Getting view from the namespace");

            /*
            if (CurNode is null)
            {
                Logging.Log("Loading par from file!!!!");

                //curPar = NodeFactory.FromFile(parPath, Yarhl.IO.FileOpenMode.Read);

                MemoryStream stream = new MemoryStream();
                using (FileStream file = File.OpenRead(parPath))
                {
                    file.CopyTo(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                }

                CurNode = NodeFactory.FromSubstream(Path.GetFileName(parPath), stream, 0, stream.Length);
                CurNode.TransformWith<ParArchiveReader, ParArchiveReaderParameters>(new ParArchiveReaderParameters() { Recursive = true });

                Logging.Log("Finished loading par!!!");
            }
            */

            //rootPath = GetCommandLineA();
            //File.WriteAllText(, );
            //File.WriteAllLines(@"C:\Users\Admin\Desktop\a14_0050.par.unpack\TESTPATH.txt", Environment.GetCommandLineArgs());
            //File.WriteAllText(@"C:\Users\Admin\Desktop\a14_0050.par.unpack\TESTPATH.txt", Environment.CommandLine);


            // change shellFolderImpl to protected, use GetCurFolder to access pidls
            // then use PidlManager.PidlToIdlist to get idlist


            List<ShellDetailColumn> columns = new List<ShellDetailColumn>();

            ShellDetailColumn col = new ShellDetailColumn("Name1", new PropertyKey(StandardPropertyKey.PKEY_ItemNameDisplay));
            columns.Add(col);

            DefaultNamespaceFolderView view = new DefaultNamespaceFolderView(columns, itemDetailProvider);

            return view;
        }

        private object itemDetailProvider(IShellNamespaceItem item, ShellDetailColumn col)
        {
            if (!(item is ParItem))
            {
                return null;
            }

            ParItem parItem = (ParItem)item;

            switch (col.Name)
            {
                case "Name1":
                    return parItem.Name;
                case "Size":
                    return parItem.CurNode.Stream.Length;
                default:
                    return null;
            }
        }
    }
}
