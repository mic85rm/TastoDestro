using EnvDTE80;
using Extensibility;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.SqlStudio.Explorer;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TastoDestro.MenuItems;

namespace TastoDestro
{

  class AvvioAutomatico : IDTExtensibility2
  {

    private static bool IsTableMenuAdded = false;
    private static bool IsColumnMenuAdded = false;
    private HierarchyObject _tableMenu = null;
    private DTE2 applicationObject = null;
    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      try
      {
        ContextService contextService = (ContextService)ServiceCache.ServiceProvider.GetService(typeof(IContextService)) ?? throw new ArgumentNullException(nameof(IContextService));
        contextService.ActionContext.CurrentContextChanged += ActionContextOnCurrentContextChanged;
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
      throw new NotImplementedException();
    }

    public void OnAddInsUpdate(ref Array custom)
    {
      throw new NotImplementedException();
    }

    public void OnStartupComplete(ref Array custom)
    {
      throw new NotImplementedException();
    }

    public void OnBeginShutdown(ref Array custom)
    {
      throw new NotImplementedException();
    }

    private void ActionContextOnCurrentContextChanged(object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      //MessageBox.Show("ciao");
      try
      {
        INodeInformation[] nodes;
        INodeInformation node;
        int nodeCount;
        IObjectExplorerService objectExplorer = (ObjectExplorerService)ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService)) ?? throw new ArgumentNullException(nameof(IObjectExplorerService));

        objectExplorer.GetSelectedNodes(out nodeCount, out nodes);
        node = nodeCount > 0 ? nodes[0] : null;

        if ((node != null) && (node.Parent != null) && (node.InvariantName != "SystemTables"))
        {
          if (node.Parent.InvariantName == "UserTables")
          {
            if (!IsTableMenuAdded)
            {
              _tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler)) ?? throw new ArgumentNullException(nameof(IMenuHandler));
              SqlTableMenuItem item = new SqlTableMenuItem(applicationObject);
              _tableMenu.AddChild(string.Empty, item);
              IsTableMenuAdded = true;
            }
          }
          else if (node.Parent.InvariantName == "Columns")
          {
            if (!IsColumnMenuAdded)
            {
              //_tableMenu = (HierarchyObject)node.GetService(typeof(IMenuHandler));
              //SqlColumnMenuItem item = new SqlColumnMenuItem(applicationObject);
              //_tableMenu.AddChild(string.Empty, item);
              //IsColumnMenuAdded = true;
            }
          }
        }
      }
      catch (Exception ObjectExplorerContextException)
      {
        MessageBox.Show(ObjectExplorerContextException.Message);
      }
    }

  }
}
