using EnvDTE80;
using Microsoft.SqlServer.Management.UI.VSIntegration.ObjectExplorer;
using Microsoft.SqlServer.Management.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using TastoDestro.Controller;
using Microsoft.VisualStudio.Shell;

namespace TastoDestro.MenuItems
{
  class SqlTableMenuItem : ToolsMenuItemBase, IWinformsMenuHandler
  {
    private DTE2 applicationObject;
    private DTEApplicationController dteController = null;
    private Regex tableRegex = null;
    enum DBTBLSCHEMA
    {
      TABELLAeSCHEMA = 2,
      DATABASE = 1
    }
    public SqlTableMenuItem(DTE2 applicationObject)
    {
      this.applicationObject = applicationObject;
      this.dteController = new DTEApplicationController();
      this.tableRegex = new Regex(Properties.Resource1.TableRegEx3);
    }

    public override object Clone()
    {
      throw new NotImplementedException();
    }

    public ToolStripItem[] GetMenuItems()
    {
      ToolStripMenuItem item = new ToolStripMenuItem("Michele");
      ToolStripMenuItem insertItem = new ToolStripMenuItem("Genera Insert con dati");
      //insertItem.Image = "";
      insertItem.Tag = false;
      insertItem.Click += new EventHandler(InsertItem_Click);
      item.DropDownItems.Add(insertItem);
      return new ToolStripItem[] { item };
    }

    protected override void Invoke()
    {
      throw new NotImplementedException();
    }

    private void InsertItem_Click(object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      ToolStripMenuItem item = (ToolStripMenuItem)sender;
      //bool generateColumnNames = (bool)item.Tag;
      string risultato = string.Empty;
      DataTable DT = new DataTable();
      DataTable table = new DataTable();
      DataTable DTtipi = new DataTable();
      DataSet dataSet = new DataSet();

      //Match match = Regex.Match(this.Parent.Context, Properties.Resource1.TableRegEx);
      MatchCollection match = Regex.Matches(this.Parent.Context, Properties.Resource1.TableRegEx3, RegexOptions.IgnoreCase);
      if ((match != null) && (match.Count == 3))
      {
        string tableName = match[(int)DBTBLSCHEMA.TABELLAeSCHEMA].Value.Split(new string[] { "'" }, StringSplitOptions.None)[1];
        string schema = match[(int)DBTBLSCHEMA.TABELLAeSCHEMA].Value.Split(new string[] { "'" }, StringSplitOptions.None)[3];
        string database = match[(int)DBTBLSCHEMA.DATABASE].Value.Split(new string[] { "'" }, StringSplitOptions.None)[1];


        string connectionString = this.Parent.Connection.ConnectionString + ";Database=" + database;

        try
        {
          using (SqlConnection connection = new SqlConnection(connectionString))
          {
            SqlCommand command = new SqlCommand(string.Format(Properties.Resource1.SQLCOMPLETA, tableName, schema), connection);
            SqlDataAdapter adapter = new SqlDataAdapter(command);



            adapter.Fill(dataSet);
          }
        }
        catch (Exception ex)
        {
          MessageBox.Show(ex.Message);
        }

        if (dataSet.Tables[2].Rows.Count > 0)
        {
          StringBuilder buffer = new StringBuilder();

          // generate INSERT prefix
          StringBuilder prefix = new StringBuilder();
          buffer.AppendFormat("USE [{0}] ", database, Environment.NewLine);
          buffer.Append(Environment.NewLine);
          buffer.AppendFormat("GO", Environment.NewLine);
          buffer.Append(Environment.NewLine);
          if (dataSet.Tables[1].Rows[0][0].ToString() == "1")
          {
            buffer.AppendFormat("SET IDENTITY_INSERT [{0}].[{1}] ON ", schema, tableName, Environment.NewLine);
            buffer.Append(Environment.NewLine);
            buffer.AppendFormat("GO", Environment.NewLine);
            buffer.Append(Environment.NewLine);
          }


          prefix.AppendFormat("INSERT [{0}].[{1}] ({2}) VALUES (", schema, tableName, dataSet.Tables[0].Rows[0][0].ToString());
          // generate INSERT statements
          foreach (DataRow row in dataSet.Tables[2].Rows)
          {
            StringBuilder values = new StringBuilder();
            for (int i = 0; i < dataSet.Tables[2].Columns.Count; i++)
            {
              if (i > 0) values.Append(", ");

              if (row.IsNull(i)) values.Append("NULL");
              else if (dataSet.Tables[2].Columns[i].DataType == typeof(int) ||
                  dataSet.Tables[2].Columns[i].DataType == typeof(long) ||
                  dataSet.Tables[2].Columns[i].DataType == typeof(double) ||
                  dataSet.Tables[2].Columns[i].DataType == typeof(float) ||
                  dataSet.Tables[2].Columns[i].DataType == typeof(byte))
                values.Append(row[i].ToString());
              else if (dataSet.Tables[2].Columns[i].DataType == typeof(decimal))
                values.Append(row[i].ToString().Replace(",", "."));
              else
                if (dataSet.Tables[3].Rows[i][1].ToString() == "varchar" || dataSet.Tables[3].Rows[i][1].ToString() == "char" || dataSet.Tables[3].Rows[i][1].ToString() == "text")
              {
                values.AppendFormat("N'{0}'", row[i].ToString().Replace("'", "''"));
              }
              else
              {
                values.AppendFormat("'{0}'", row[i].ToString().Replace("'", "''"));
              }
            }
            values.AppendFormat(")");

            buffer.AppendLine(prefix.ToString() + values.ToString());
            buffer.AppendFormat("GO", Environment.NewLine);
            buffer.Append(Environment.NewLine);
          }
          prefix.AppendFormat("GO", Environment.NewLine);
          prefix.Append(Environment.NewLine);
          if (dataSet.Tables[1].Rows[0][0].ToString() == "1")
          {
            buffer.AppendFormat("SET IDENTITY_INSERT [{0}].[{1}] OFF ", schema, tableName, Environment.NewLine);
            buffer.Append(Environment.NewLine);
            buffer.AppendFormat("GO", Environment.NewLine);
          }

          // create new document

          this.dteController.CreateNewScriptWindow(buffer);
        }
        else
        {
          MessageBox.Show("Non ci sono dati in questa tabella");
        }
      }
    }
  }
}
