﻿using EnvDTE;
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
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TastoDestro.Controller;
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

            //    [208]   "Menu di scelta rapida editor file SQL" object { Microsoft.VisualStudio.PlatformUI.Automation.CommandBar._Marshaler}
            Microsoft.VisualStudio.CommandBars.CommandBar sqlQueryPanel = ((CommandBars)applicationObject.CommandBars)[208];
            CommandBarPopup commandBarPopup = (CommandBarPopup)sqlQueryPanel.Controls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, 6, true);
            CommandBarControl cmdBarControl3 = commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, 1, true);

      
            //CommandBarControl cmdBarControl3 = sqlQueryPanel.Controls.Add(MsoControlType.msoControlButton,Missing.Value, Missing.Value, 6, true);

            commandBarPopup.Caption = "Inserisci Snippet";


  
            

            var myButton2 = (CommandBarButton)cmdBarControl3;
          
            myButton2.Visible = true;
            myButton2.Style = MsoButtonStyle.msoButtonIconAndCaption;
            myButton2.Enabled = true;
            //myButton2.Picture = IconeMenu.GetIPictureDispFromPicture(IconeMenu.LoadBase64(Properties.Resource1.ICONAEXCEL));
            myButton2.Caption = "Biposte..Personale";
            myButton2.Click += MyButton2_Click;


            var id =((CommandBars)applicationObject.CommandBars)["SQL Results Grid Tab Context"].Index;
            Microsoft.VisualStudio.CommandBars.CommandBar sqlQueryGridPane = ((CommandBars)applicationObject.CommandBars)["SQL Results Grid Tab Context"];
            IObjectExplorerService objectExplorer = (ObjectExplorerService)ServiceCache.ServiceProvider.GetService(typeof(IObjectExplorerService)) ?? throw new ArgumentNullException(nameof(IObjectExplorerService));
            CommandBarPopup commandBarPopup1 = (CommandBarPopup)sqlQueryGridPane.Controls.Add(MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            CommandBarControl cmdBarControl2 = commandBarPopup1.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            CommandBarControl cmdBarControlCopiaFormattata = sqlQueryGridPane.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, 4, true);
            cmdBarControlCopiaFormattata.Caption = "Copia per INSERT";
            var btnCopiaFormattata= (CommandBarButton)cmdBarControlCopiaFormattata;
            btnCopiaFormattata.Visible = true;
            btnCopiaFormattata.Enabled = true;
            btnCopiaFormattata.Caption = "Copia per INSERT";
            btnCopiaFormattata.Style= MsoButtonStyle.msoButtonIconAndCaption;
            // btnCopiaFormattata.Picture = IconeMenu.GetIPictureDispFromPicture(IconeMenu.LoadBase64(Properties.Resource1.ICONAEXCEL));
            btnCopiaFormattata.Click += BtnCopiaFormattata_Click;


            commandBarPopup1.Caption = "Salva nel formato";
            var myButton = (CommandBarButton)cmdBarControl2;

            myButton.Visible = true;

            myButton.Enabled = true;

            myButton.Caption = "Salva in Excel";
            myButton.Style = MsoButtonStyle.msoButtonIconAndCaption;

            myButton.Picture = IconeMenu.GetIPictureDispFromPicture(IconeMenu.LoadBase64(Properties.Resource1.ICONAEXCEL));

            CommandBarControl cmdBarControl4 = commandBarPopup1.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);

            myButton.Click += new _CommandBarButtonEvents_ClickEventHandler(btnMEssageBoxxResults_Click);

            var myButton3 = (CommandBarButton)cmdBarControl4;

            myButton3.Visible = true;

            myButton3.Enabled = true;

            myButton3.Caption = "Salva in JSON";
            myButton3.Style = MsoButtonStyle.msoButtonIconAndCaption;

            myButton3.Picture = IconeMenu.GetIPictureDispFromPicture(IconeMenu.LoadBase64(Properties.Resource1.ICONAJSON));



            myButton3.Click += new _CommandBarButtonEvents_ClickEventHandler(MyButton3_Click);
           

            await Command1.InitializeAsync(this);
 
    }

        private void BtnCopiaFormattata_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            CopiaColonne();
        }

        private void MyButton2_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            DTEApplicationController dteController = new DTEApplicationController();
     
            dteController.WriteToWindow(applicationObject,"ciao");
    }

        private void MyButton3_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileOk += SaveFileDialog_FileOk1;
            saveFileDialog.DefaultExt = "json";
            saveFileDialog.Filter = "JSON files (*.json)|*.json";
            saveFileDialog.ShowDialog();
        }

        private void SaveFileDialog_FileOk1(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string json=DataTableToJSON(SalvaDatatable());
            //var file = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Test_" + DateTime.Now.ToString("M-dd-yyyy-HH.mm.ss") + ".json"));
            var file = new FileInfo(((SaveFileDialog)sender).FileName);
            System.IO.File.WriteAllText(file.ToString(), json);
        }

        private void btnMEssageBoxxResults_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //  Microsoft.VisualStudio.CommandBars.CommandBar sqlQueryGridPane = ((CommandBars)applicationObject.CommandBars)["SQL Results Grid Tab Context"];

            //IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            //Guid generalPaneGuid = VSConstants.GUID_OutWindowGeneralPane; // P.S. There's also the GUID_OutWindowDebugPane available.
            //IVsOutputWindowPane generalPane;
            //outWindow.GetPane(ref generalPaneGuid, out generalPane);

            //generalPane.OutputString("Hello World!");
            //generalPane.Activate(); // Brings this pane into view
            // vsWindowTypeOutput
            //String output = "ST: 0:0:{ 34e76e81 - ee4a - 11d0 - ae2e - 00a0c90fffc3}";

            //Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            //Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            //var collezione = window.Collection;
            //var michele = window.Selection;
            //var michele2 = window.Object;
            //var finestra = collezione.Item(3);
            //var testo = finestra.Selection;
            //finestra.Activate();

         //   var ciao=Ctrl.Control;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileOk += SaveFileDialog_FileOk;
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.ShowDialog();
        }



        public class ColonneDaCopiare
        {
            private int _IDColonna;
            private int _IDUltimaRiga;
         
            public int IDColonna { get { return _IDColonna; } set { _IDColonna = value; } }
            public int IDUltimaRiga { get { return _IDUltimaRiga; } set { _IDUltimaRiga = value; } }

            public ColonneDaCopiare(int IDColonna,int IDUltimaRiga) {
                this.IDColonna = IDColonna;
                this.IDUltimaRiga = IDUltimaRiga;
            }

          

            public bool  GetColonne(int colonna) { if (IDColonna == colonna) { return true; } else { return false; } }
            
        }


        public void CopiaColonne()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = (DTE)GetService(typeof(DTE));
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });
            var objType2 = Result.GetType();
            var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
            var SQLResultsControl = field.GetValue(Result);
            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;
            var data = new DataTable();
            bool inserito = false;
            StringBuilder sb = new StringBuilder();
            foreach (var gridContainer in gridContainers)
            {
               
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                if (grid.ContainsFocus == true) { 
                var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;
                var NumeroColonneSelezionata = grid.SelectedCells;
                 Microsoft.SqlServer.Management.UI.Grid.SelectionManager selectionManager= (SelectionManager)GetNonPublicField(grid, "m_selMgr");
                    var scroll = GetNonPublicField(grid,"m_scrollMgr");
                  
                List<ColonneDaCopiare> ColonneRighe = new List<ColonneDaCopiare>();
                for (int indice=0;indice< NumeroColonneSelezionata.Count;indice++)
                {
                        //ColonneDaCopiare colonneDaCopiare = new ColonneDaCopiare(((int)GetNonPublicProperties(NumeroColonneSelezionata[indice], "LastUpdatedCol") - 1), Convert.ToInt32(GetNonPublicProperties(NumeroColonneSelezionata[indice], "LastUpdatedRow")));
                        for (int jj = (int)NumeroColonneSelezionata[indice].OriginalX; jj <= NumeroColonneSelezionata[indice].Right; jj++)
                        {
                            for (int ii =(int) NumeroColonneSelezionata[indice].OriginalY; ii <= NumeroColonneSelezionata[indice].Bottom; ii++)
                            {
                                ColonneDaCopiare colonneDaCopiare = new ColonneDaCopiare(jj-1,ii);
                                ColonneRighe.Add(colonneDaCopiare);
                            }
                        }
                    }
                ////if (ColonneRighe.Count == 1)
                ////{
                //        for (int idc = (int)selectionManager.CurrentColumn; idc < selectionManager.LastUpdatedColumn+1; idc++) { 
                //    for (int idx = (int)selectionManager.CurrentRow; idx < selectionManager.LastUpdatedRow+1; idx++)
                //    {
                //        ColonneDaCopiare _colonneDaCopiare = new ColonneDaCopiare(idc - 1, idx);
                //        ColonneRighe.Add(_colonneDaCopiare);
                //    }
                //        }
                //    //}
                string columnName;
                Type columnType;
                    
                for (long i = 0; i < gridStorage.NumRows(); i++)
                {
                    var rowItems = new List<object>();
                        
                    if ((i > 0) && (ColonneRighe.Count > 0)&&(ColonneRighe.FindAll(x => x.IDUltimaRiga==i).Count()>0)) {
                            if (ColonneRighe.FindAll(x=>x.IDColonna!=x.IDColonna).Count > 1) { 
                            sb.Remove(sb.Length - 1,1);
                            }

                            sb.AppendLine(); }
                    for (int c = 0; c < schemaTable.Rows.Count; c++)
                    {
                      var risultato=ColonneRighe.Find(x=>x.IDColonna==c);
                    
                     
                        if (risultato!=null) { 
                         columnName = schemaTable.Rows[c][0].ToString();
                         columnType = schemaTable.Rows[c][12] as Type;
                       
                        if (!data.Columns.Contains(columnName))
                        {
                            data.Columns.Add(columnName, columnType);
                        }
                              //  ColonneDaCopiare colonnaDaCercare = new ColonneDaCopiare(c, (int)i);
                                var result = ColonneRighe.Find(x => x.IDUltimaRiga == i && x.IDColonna==c);
                            if (result!=null)
                            {
                                var cellData = gridStorage.GetCellDataAsString(i, c+1 );

                                //if ((cellData == "NULL")|| (i!=risultato.IDUltimaRiga))
                                if (cellData == "NULL")
                                {
                                    rowItems.Add(null);
                                        sb.Append("NULL,");
                                    inserito = false;
                                    continue;
                                }

                                    if (columnType == typeof(bool) )
                                    {
                                        cellData = cellData == "0" ? "False" : "True";
                                    }

                                  // Console.WriteLine($"Parsing {columnName} with '{cellData}'");

                                    var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);
                                rowItems.Add(typedValue);
                                if (columnType.Name == "String")
                                {
                                    typedValue = String.Format("'{0}',", typedValue.ToString().Trim().TrimEnd().TrimStart());
                                }
                                if (columnType.Name == "DateTime")
                                {
                                    DateTime? dateTime = (DateTime)typedValue;
                                    string sqlFormattedDate = dateTime.HasValue ? dateTime.Value.ToString("yyyy-MM-dd HH:mm:ss"): "<errore>";
                                    typedValue = String.Format("'{0}',",sqlFormattedDate );
                                }
                                    if ((columnType.Name == "Byte")||(columnType.Name=="Int32") || (columnType.Name == "Double"))
                                    {
                                       
                                        typedValue = String.Format("{0},", typedValue);
                                    }
                                    if(columnType.Name == "Boolean")
                                    {
                                       var valore=Convert.ToBoolean(typedValue) == false ? "0" : "1";
                                        typedValue = String.Format("{0},", valore );
                                    }
                                    sb.Append(typedValue);
                                inserito = true;
                            }
                            else { inserito = false; }
                        }
                            ColonneDaCopiare colonnaDaEliminare = new ColonneDaCopiare(c,(int)i);
                            ColonneRighe.Remove(colonnaDaEliminare);
                        }

                        //data.Rows.Add(rowItems.ToArray());
                        
                    }

                    // data.AcceptChanges();

                }

            }
            Clipboard.SetText(sb.ToString().Remove(sb.Length-1));
       

        }


        public String EstraiColonna(int idcolonna,int numerorighe,int tipocolonna,DataTable dt) {
            StringBuilder sb = new StringBuilder();
            String Appoggio = string.Empty;
            foreach (DataRow row in dt.Rows)
            {

                if (dt.Rows.IndexOf(row) <= numerorighe)
                {
                    if (tipocolonna == (int)ETipiColonna.String)
                    {
                        Appoggio = String.Format("'{0}'",row.ToString());
                    }
                    sb.Append(Appoggio);
                }

            }

            return sb.ToString(); ;
        }

        //{c6b71c93-b881-3346-bbe2-3635c2fbcd6c}
        public DataTable SalvaDatatable() {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = (DTE)GetService(typeof(DTE));
            var objType = ServiceCache.ScriptFactory.GetType();
            var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
            var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });
            var objType2 = Result.GetType();
            var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
            var SQLResultsControl = field.GetValue(Result);
            var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
            CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;
            var ElementoFocusato = GetNonPublicProperties(m_gridResultsPage, "LastFocusedControl");
            var tagFocusato = GetNonPublicProperties(ElementoFocusato, "Tag");
            var data = new DataTable();
            foreach (var gridContainer in gridContainers)
            {
                var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
                if (grid.Tag==tagFocusato)
                {
                    var gridStorage = grid.GridStorage;
                var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;
                var NumeroColonneSelezionata = grid.SelectedCells;
                Microsoft.SqlServer.Management.UI.Grid.SelectionManager selectionManager = (SelectionManager)GetNonPublicField(grid, "m_selMgr");
                var prova=selectionManager.SelectedBlocks.Capacity;
             
                 
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

                            //Console.WriteLine($"Parsing {columnName} with '{cellData}'");
                            var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

                            rowItems.Add(typedValue);
                        }

                        data.Rows.Add(rowItems.ToArray());
                    }
                
                data.AcceptChanges();

                }
            }
           
            return data;
        }

        public string DataTableToJSON(DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }


        private void SaveFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            List<int> listaIndexDate = new List<int>();
            DataTable dt = SalvaDatatable();
            listaIndexDate.AddRange(dt.Columns.Cast<DataColumn>().Where(c => c.DataType == typeof(DateTime)).Select(x=>x.Ordinal));
            if ((dt != null) && (dt.Rows.Count > 0)) { 
            // var file = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Test_" + DateTime.Now.ToString("M-dd-yyyy-HH.mm.ss") + ".xlsx"));
            var file = new FileInfo(((SaveFileDialog)sender).FileName);
            using (var package = new ExcelPackage(file))
            {
                    ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Test");
                 
                    if (ws != null)
                   {package.Workbook.Worksheets.Delete("Test"); }
                    ws = package.Workbook.Worksheets.Add("Test");

                    ws.Cells[1, 1].LoadFromDataTable(dt, true);
 
                    ws.Cells[string.Format("A1:{0}1",GetColumnName(dt.Columns.Count-1))].AutoFilter = true;
                    ws.Cells[string.Format("A1:{0}1", GetColumnName(dt.Columns.Count-1))].Style.Fill.PatternType= ExcelFillStyle.Solid;
                    ws.Cells[string.Format("A1:{0}1", GetColumnName(dt.Columns.Count-1))].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    ws.Cells[string.Format("A1:{0}1", GetColumnName(dt.Columns.Count - 1))].Style.Font.Bold=true;
                    listaIndexDate.ForEach(x=> ws.Column(x+1).Style.Numberformat.Format = "dd-mm-yyyy");
                    ws.Cells.AutoFitColumns();
                    package.Save();
                //MessageBox.Show("Excel correttamente Salvato");
            }
           }
        }
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }


        public object GetNonPublicField(object obj, string field)
        {
            FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance|BindingFlags.Public );

            return f.GetValue(obj);
        }

        public object GetNonPublicProperties(object obj, string field)
        {
            PropertyInfo f = obj.GetType().GetProperty(field, BindingFlags.NonPublic | BindingFlags.Instance|BindingFlags.Public);

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
