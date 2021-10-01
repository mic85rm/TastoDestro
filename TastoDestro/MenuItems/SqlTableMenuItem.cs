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

namespace TastoDestro.MenuItems
{
  class SqlTableMenuItem : ToolsMenuItemBase, IWinformsMenuHandler
  {
    private DTE2 applicationObject;
    private DTEApplicationController dteController = null;
    private Regex tableRegex = null;
    enum DBTBLSCHEMA
    {
      TABELLA = 9,
      SCHEMA = 12,
      DATABASE = 6
    }
    public SqlTableMenuItem(DTE2 applicationObject)
    {
      this.applicationObject = applicationObject;
      this.dteController = new DTEApplicationController();
      this.tableRegex = new Regex(Properties.Resource1.TableRegEx2);
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
      ToolStripMenuItem item = (ToolStripMenuItem)sender;
      bool generateColumnNames = (bool)item.Tag;



      //Match match = Regex.Match(this.Parent.Context, Properties.Resource1.TableRegEx);
      MatchCollection match = Regex.Matches(this.Parent.Context, Properties.Resource1.TableRegEx2, RegexOptions.IgnoreCase);
      if ((match != null) && (match.Count == 13))
      {

        string tableName = match[(int)DBTBLSCHEMA.TABELLA].Value;
        string schema = match[(int)DBTBLSCHEMA.SCHEMA].Value;
        string database = match[(int)DBTBLSCHEMA.DATABASE].Value;


        string connectionString = this.Parent.Connection.ConnectionString + ";Database=" + database;

        SqlCommand command = new SqlCommand(string.Format(Properties.Resource1.SQLSELECT, schema, tableName));
        command.Connection = new SqlConnection(connectionString);
        command.Connection.Open();

        SqlDataAdapter adapter = new SqlDataAdapter(command);
        DataTable table = new DataTable();
        adapter.Fill(table);

        command.Connection.Close();

        StringBuilder buffer = new StringBuilder();

        // generate INSERT prefix
        StringBuilder prefix = new StringBuilder();
        if (generateColumnNames)
        {
          prefix.AppendFormat("INSERT INTO {0} (", tableName);
          for (int i = 0; i < table.Columns.Count; i++)
          {
            if (i > 0) prefix.Append(", ");
            prefix.AppendFormat("[{0}]", table.Columns[i].ColumnName);
          }
          prefix.Append(") VALUES (");
        }
        else
          prefix.AppendFormat("INSERT INTO {0} VALUES (", tableName);

        // generate INSERT statements
        foreach (DataRow row in table.Rows)
        {
          StringBuilder values = new StringBuilder();
          for (int i = 0; i < table.Columns.Count; i++)
          {
            if (i > 0) values.Append(", ");

            if (row.IsNull(i)) values.Append("NULL");
            else if (table.Columns[i].DataType == typeof(int) ||
                table.Columns[i].DataType == typeof(decimal) ||
                table.Columns[i].DataType == typeof(long) ||
                table.Columns[i].DataType == typeof(double) ||
                table.Columns[i].DataType == typeof(float) ||
                table.Columns[i].DataType == typeof(byte))
              values.Append(row[i].ToString());
            else
              values.AppendFormat("'{0}'", row[i].ToString());
          }
          values.AppendFormat(")");

          buffer.AppendLine(prefix.ToString() + values.ToString());
        }

        // create new document
        this.dteController.CreateNewScriptWindow(buffer);
      }
    }
  }
}
