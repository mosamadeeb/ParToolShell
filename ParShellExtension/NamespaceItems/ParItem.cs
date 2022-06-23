using SharpShell.Diagnostics;
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
    public class ParItem : IShellNamespaceItem
    {
        public string Name;

        public Node CurNode;

        public ParItem(Node node)
        {
            CurNode = node;

            Name = node.Name;
            Logging.Log($"Creating a file {Name}");
        }

        public AttributeFlags GetAttributes()
        {
            return AttributeFlags.IsCompressed
                | AttributeFlags.CanBeDeleted
                | AttributeFlags.CanBeMoved;
        }

        public string GetDisplayName(DisplayNameContext displayNameContext)
        {
            return Name;
        }

        public Icon GetIcon()
        {
            return null;
        }

        public ShellId GetShellId()
        {
            return ShellId.FromString(Name);
        }
    }
}
