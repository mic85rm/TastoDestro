using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.SqlStudio.Explorer;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TastoDestro.Helper;
using TastoDestro.MenuItems;
using TastoDestro.Model;
using static Microsoft.VisualStudio.VSConstants;
using static System.Net.Mime.MediaTypeNames;
using Task = System.Threading.Tasks.Task;

namespace TastoDestro
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  /// </summary>
  /// <remarks>
  /// <para>
  /// The minimum requirement for a class to be considered a valid package for Visual Studio
  /// is to implement the IVsPackage interface and register itself with the shell.
  /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
  /// to do it: it derives from the Package class that provides the implementation of the
  /// IVsPackage interface and uses the registration attributes defined in the framework to
  /// register itself and its components with the shell. These attributes tell the pkgdef creation
  /// utility what data to put into .pkgdef file.
  /// </para>
  /// <para>
  /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
  /// </para>
  /// </remarks>
  /// 
  [ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
  [ProvideService(typeof(Command1), IsAsyncQueryable = true)]
  [ProvideMenuResource("Menus.ctmenu", 1)]
  [Guid(Command1Package.PackageGuidString)]


  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
  public sealed class Command1Package : AsyncPackage
  {
    /// <summary>
    /// Command1Package GUID string.
    /// </summary>
    //public const string PackageGuidString = "2aa4241b-3cd0-49d1-8592-b9c5f593fab1";
    public const string PackageGuidString = "f8750854-1ffd-4804-aa8c-ab8cd791448c";
    //private DTE2 applicationObject = null;
    //private AddIn addInInstance = null;
    private static bool IsTableMenuAdded = false;
    private static bool IsColumnMenuAdded = false;
    private HierarchyObject _tableMenu = null;
    private DTE2 applicationObject = null;
        //private GestioneDocumenti gestioneDocumenti = null;
    /// <summary>
    /// Initializes a new instance of the <see cref="Command1Package"/> class.
    /// </summary>
        public Command1Package()
    {
      // Inside this method you can place any initialization code that does not require
      // any Visual Studio service because at this point the package object is created but
      // not sited yet inside Visual Studio environment. The place to do all the other
      // initialization is the Initialize method.
    }

    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
    /// <param name="progress">A provider for progress updates.</param>
    /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
    protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
    {
      // When initialized asynchronously, the current thread may be a background thread at this point.
      // Do any initialization that requires the UI thread after switching to the UI thread.
      //await base.InitializeAsync(cancellationToken, progress);
      await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);


      //addInInstance = (AddIn)addInInst;
      try
      {
              


                ContextService contextService = (ContextService)ServiceCache.ServiceProvider.GetService(typeof(IContextService)) ?? throw new ArgumentNullException(nameof(IContextService));
        contextService.ActionContext.CurrentContextChanged += ActionContextOnCurrentContextChanged;
                DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
               // DTE2 dte = (DTE2)await GetServiceAsync(typeof(EnvDTE.DTE));
                // dte.Events.DocumentEvents.DocumentClosing += DocumentEvents_DocumentClosing;
                //dte.Events.WindowEvents.WindowClosing += WindowEvents_DocumentClosing;
                GestioneDocumenti gestioneDocumenti = new GestioneDocumenti(this);
                gestioneDocumenti.BeforeSave += GestioneDocumenti_BeforeSave;
            
            }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
      applicationObject = (DTE2)await GetServiceAsync(typeof(DTE));

            var outputWindowEvents = applicationObject.Events.OutputWindowEvents["Output"];
            outputWindowEvents.PaneUpdated += OutputWindowEvents_PaneUpdated;
            Microsoft.VisualStudio.CommandBars.CommandBar sqlQueryGridPane = ((CommandBars)applicationObject.CommandBars)["SQL Results Grid Tab Context"];

            CommandBarControl cmdBarControl2 = sqlQueryGridPane.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);

            var myButton = (CommandBarButton)cmdBarControl2;

            myButton.Visible = true;

            myButton.Enabled = true;

            myButton.Caption = "Salva in Excel";

            //myButton.FaceId = 224;

            myButton.Style = MsoButtonStyle.msoButtonIconAndCaption;

            myButton.Click += new _CommandBarButtonEvents_ClickEventHandler(btnMEssageBoxxResults_Click);


            await Command1.InitializeAsync(this);
 
    }

        private void btnMEssageBoxxResults_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Microsoft.VisualStudio.CommandBars.CommandBar sqlQueryGridPane = ((CommandBars)applicationObject.CommandBars)["SQL Results Grid Tab Context"];

            //IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            //Guid generalPaneGuid = VSConstants.GUID_OutWindowGeneralPane; // P.S. There's also the GUID_OutWindowDebugPane available.
            //IVsOutputWindowPane generalPane;
            //outWindow.GetPane(ref generalPaneGuid, out generalPane);

            //generalPane.OutputString("Hello World!");
            //generalPane.Activate(); // Brings this pane into view
            // vsWindowTypeOutput
            //String output = "ST: 0:0:{ 34e76e81 - ee4a - 11d0 - ae2e - 00a0c90fffc3}";
            DTE dte = (DTE)GetService(typeof(DTE));
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });
            var objType2 = Result.GetType();
            var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
            var SQLResultsControl = field.GetValue(Result);
            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;
            foreach (var gridContainer in gridContainers)
            {
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;
                var data = new DataTable();

                for (long i = 0; i < gridStorage.NumRows(); i++)
                {
                    var rowItems = new List<object>();

                    for (int c = 0; c < schemaTable.Rows.Count; c++)
                    {
                        var columnName = schemaTable.Rows[c][0].ToString();
                        var columnType = schemaTable.Rows[c][12] as Type;

                        if (!data.Columns.Contains(columnName))
                        {
                            data.Columns.Add(columnName, columnType);
                        }

                        var cellData = gridStorage.GetCellDataAsString(i, c + 1);

                        if (cellData == "NULL")
                        {
                            rowItems.Add(null);

                            continue;
                        }

                        if (columnType == typeof(bool))
                        {
                            cellData = cellData == "0" ? "False" : "True";
                        }

                        Console.WriteLine($"Parsing {columnName} with '{cellData}'");
                        var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

                        rowItems.Add(typedValue);
                    }

                    data.Rows.Add(rowItems.ToArray());
                }

                data.AcceptChanges();
            }
            //Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            var collezione = window.Collection;
            var michele = window.Selection;
            var michele2 = window.Object;
            var finestra = collezione.Item(3);
            var testo = finestra.Selection;
            finestra.Activate();
          
        }
        public object GetNonPublicField(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

            return f.GetValue(obj);
        }

        private void OutputWindowEvents_PaneUpdated(OutputWindowPane pPane)
        {
            throw new NotImplementedException();
        }

        private void GestioneDocumenti_BeforeSave(object sender, Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
           
            document.Save(string.Format("C:\\{0}",document.Name));
           //MessageBox.Show("finalmente");
        }

    //    private void WindowEvents_DocumentClosing(Window window)
    //    {
    //        ThreadHelper.ThrowIfNotOnUIThread();
    //        var michele=window.DocumentData;
    //        //MessageBox.Show("ciao");
    //    }


    //    private void DocumentEvents_DocumentClosing(Document window)
    //{
    //        ThreadHelper.ThrowIfNotOnUIThread();
          
    //  //MessageBox.Show("ciao");
    //}


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



    protected override int QueryClose(out bool pfCanClose)
    {
      pfCanClose = true;
      return VSConstants.S_OK;
    }

    #endregion
  }
}
