using SharpShell.Diagnostics;
using SharpShell.Interop;
using SharpShell.Pidl;
using SharpShell.SharpNamespaceExtension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yarhl.FileSystem;

namespace ParShellExtension.NamespaceItems
{
    public class ParFolder : ParItem, IShellNamespaceFolder //, IShellNamespaceFolderContextMenuProvider
    {
        /*
        public IContextMenu CreateContextMenu(IdList folderIdList, IdList[] folderItemIdLists)
        {
            throw new NotImplementedException();
        }
        */

        private List<ParItem> Items = null;
        private List<ParFolder> Folders = null;

        private bool createdChildren = false;

        public ParFolder(Node node)
            : base(node)
        {
            Logging.Log($"Creating a folder {Name}");

            Items = new List<ParItem>();
            Folders = new List<ParFolder>();
        }

        public AttributeFlags GetAttributes()
        {
            return AttributeFlags.IsFolder
                //| AttributeFlags.CanBeDeleted
                //| AttributeFlags.CanBeMoved
                //| AttributeFlags.IsBrowsable
                //| AttributeFlags.IsDropTarget
                //| AttributeFlags.IsCompressed
                | AttributeFlags.MayContainSubFolders;
        }

        public IEnumerable<IShellNamespaceItem> GetChildren(ShellNamespaceEnumerationFlags flags)
        {
            Logging.Log($"Enumerating children for folder {Name}");

            if (CurNode is null)
            {
                yield break;
            }

            if (!createdChildren)
            {
                if (CurNode.Children.Count == 0)
                {
                    yield break;
                }

                ParItem item;
                ParFolder folder;

                foreach (Node c in CurNode.Children)
                {
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

                createdChildren = true;
                yield break;
            }

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

        public Icon GetIcon()
        {
            return null;
        }

        public ShellNamespaceFolderView GetView()
        {
            Logging.Log($"Getting view from folder {Name}");

            List<ShellDetailColumn> columns = new List<ShellDetailColumn>();

            ShellDetailColumn col = new ShellDetailColumn("Name", new PropertyKey(StandardPropertyKey.PKEY_ItemNameDisplay));
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
                case "Name":
                    return parItem.Name;
                case "Size":
                    return parItem.CurNode.Stream.Length;
                default:
                    return null;
            }
        }
    }
}
