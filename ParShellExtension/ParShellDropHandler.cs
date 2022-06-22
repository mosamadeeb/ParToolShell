using SharpShell.Attributes;
using SharpShell.SharpDropHandler;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

using static ParShellExtension.ParShellCommon;

namespace ParShellExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.FileExtension, ".par", ".par_c", ".par_e", ".par_j", ".par_k", ".par_z")]
    public class ParShellDropHandler : SharpDropHandler
    {
        protected override void DragEnter(DragEventArgs dragEventArgs)
        {
            dragEventArgs.Effect = DragDropEffects.Copy;
        }

        protected override void Drop(DragEventArgs dragEventArgs)
        {
            Task.Run(() => ParAdd(Path.GetDirectoryName(SelectedItemPath), DragItems.ToList(), SelectedItemPath, false));
        }
    }
}
