using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Core;

#region Description
//com--> Microsoft Office Object Library
//UnCheck Debug--> Windows--> Exception Setting --> Managed Debuging Assistants --> Context Swith DeadLock
#endregion Description

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(@"c:\test\", "*xls",SearchOption.AllDirectories);
            string ConnectionStr = "Data Source=172.16.107.195;Initial Catalog=Mydb ; Integrated Security = True";

            SqlConnection Sqlcon = new SqlConnection(ConnectionStr);
            AddTable(Sqlcon);

            for (int k = 0; k < files.Count(); k++)
            {
                string FileName = files[k].ToString();
                DataTable dt = new DataTable();
                string FileExtension = Path.GetExtension(FileName);
                string provider = null;
                OleDbConnection oleDbConnection = new OleDbConnection();
                string Query = null;
                string SheetName = null;
                string TableName = null;
                DataTable ColumnTable = new DataTable();

                if (FileExtension.ToLower() == ".xls")
                {
                    oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";" + "Extended Properties ='Excel 8.0;HDR=NO;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0;'";
                    provider = "'Microsoft.ACE.OLEDB.12.0','Excel 8.0;DataBase=" + FileName + ";HDR=NO'";
                }
                else if (FileExtension.ToLower() == ".xlsx")
                {
                    oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";" + "Extended Properties ='Excel 12.0 xml;HDR=NO;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0;'";
                    provider = "'Microsoft.ACE.OLEDB.12.0','Excel 12.0;DataBase=" + FileName + ";HDR=NO'";
                }

                try
                {
                    oleDbConnection.Open();

                    dt = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    DataTable Source_DataTable = new DataTable();
                    string Query2 = null;

                    Query2 = "select cast(trim(name)as nvarchar(1000))as name from sys.all_columns where OBJECT_NAME(object_id)=N'xlsss' ";
                    if (Sqlcon.State != ConnectionState.Open)
                    {
                        Sqlcon.Open();
                    }
                    SqlDataAdapter columnAdapter = new SqlDataAdapter(Query2, Sqlcon);
                    ColumnTable.Reset();
                    columnAdapter.Fill(ColumnTable);
                    Sqlcon.Close();

                    foreach (DataRow row in dt.Rows)
                    {
                        SheetName = row["Table_Name"].ToString();
                        provider = provider.Replace("\\", @"\");
                        TableName = FileName + SheetName;
                        oleDbConnection.Close();

                        OleDbDataAdapter dataAdapter = new OleDbDataAdapter("select * from [" + SheetName + "]", oleDbConnection);
                        Source_DataTable.Reset();
                        dataAdapter.Fill(Source_DataTable);
                        //Source_DataTable = TraverseExcel(Source_DataTable);

                        #region Delete Null Column 
                        //باید بهینه شود به جای لوپ از سلکت استفاده کن 
                        for (int i = 0; i < Source_DataTable.Columns.Count; i++)
                        {
                            int Row_NotNull = 0;
                            for (int j = 0; j < Source_DataTable.Rows.Count; j++)
                            {
                                if (Row_NotNull < 2)
                                {
                                    Source_DataTable.Columns.RemoveAt(i);
                                    return;
                                }
                                else if (Source_DataTable.Rows[j][i].ToString() != null && Source_DataTable.Rows[j][i].ToString().Length > 0)
                                {
                                    Row_NotNull++;
                                }

                            }

                        }

                        #endregion  Delete Null Column 
                        
                        #region Set Column Name
                        int RowColName = 0;
                        for (int i = 0; i < Source_DataTable.Rows.Count; i++)
                        {
                            int RowNull = 0;
                            for (int j = 0; j < Source_DataTable.Columns.Count; j++)
                            {
                                RowNull = 0;
                                if(Source_DataTable.Rows[i][j].ToString()==null|| Source_DataTable.Rows[i][j].ToString().Length<1)
                                RowNull++;
                                break;
                            }
                            if (RowNull < 1)
                            {
                                RowColName = i;
                                break;
                            }
                        }
                        
                        for(int i=0; i<Source_DataTable.Columns.Count;i++)
                        {
                            if(Source_DataTable.Rows[RowColName][i].ToString().Length>0)
                            {
                                Source_DataTable.Columns[i].ColumnName = Clean(Source_DataTable.Rows[RowColName][i].ToString());
                            }

                        }

                        #endregion Set Column Name

                        Source_DataTable.Columns.Add("FileName");
                        foreach (DataRow dr in Source_DataTable.Rows)
                        {
                            dr["FileName"] = TableName;
                        }
                        SqlBulkCopy conBulk = new SqlBulkCopy(Sqlcon);
                        #region Mapping Source Columns  by Desc Column 
                        for (int i = 0; i < Source_DataTable.Columns.Count; i++)
                        {
                            //sSource_DataTable.Columns[i].ColumnName = Clean(Source_DataTable.Columns[i].ColumnName);
                            int ismap = 0;
                            foreach (DataRow dr_columns in ColumnTable.Rows)
                            {
                                dr_columns["name"] = Clean(dr_columns["name"].ToString());
                                if(Source_DataTable.Columns[i].ColumnName==dr_columns["name"].ToString()&& ismap!=1 )
                                {
                                    SqlBulkCopyColumnMapping col = new SqlBulkCopyColumnMapping();
                                    col.DestinationColumn = Source_DataTable.Columns[i].ColumnName;
                                    col.SourceColumn = Source_DataTable.Columns[i].ColumnName;
                                    conBulk.ColumnMappings.Add(col);
                                    ismap = 1;
                                    break;
                                }
                            }
                            if(ismap==0 && ColumnTable.Rows.Count<431)
                            {
                                AddColumn(Source_DataTable.Columns[i].ColumnName, Sqlcon);
                                SqlBulkCopyColumnMapping col = new SqlBulkCopyColumnMapping();
                                col.DestinationColumn = Source_DataTable.Columns[i].ColumnName;
                                col.SourceColumn = Source_DataTable.Columns[i].ColumnName;
                                conBulk.ColumnMappings.Add(col);
                                ColumnTable.Rows.Add(Source_DataTable.Columns[i].ColumnName);
                            }
                        }
                        #endregion Mapping Source Columns  by Desc Column 

                        conBulk.DestinationTableName = "xlsss";
                        try
                        {
                            conBulk.BatchSize = 10000;
                            conBulk.WriteToServer(Source_DataTable);
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                SheetName = SheetName ?? "";
                                FileName = FileName ?? "";
                                Query = "insert ErrorSheet (FileName,SheetName , ErrorDesc)values ("+"N'"+FileName.ToString().Replace("'","")+"'"+", N"+SheetName.ToString().Replace("'","")+" , "+"N'"+ex.Message.ToString().Replace("'","")+"'" + " )";
                                if (Sqlcon.State != ConnectionState.Open)
                                {
                                    Sqlcon.Open();
                                }
                                SqlCommand sqlCommand = new SqlCommand(Query,Sqlcon);
                                sqlCommand.ExecuteNonQuery();
                                Sqlcon.Close();
                            }
                            catch (Exception)
                            {
                                SheetName = SheetName ?? "";
                                FileName = FileName ?? "";
                                Query = "insert ErrorSheet (FileName,SheetName , ErrorDesc)values (" + "N'" + FileName.ToString().Replace("'", "") + "'" + ", N" + SheetName.ToString().Replace("'", "") + "'" + " , " + "N'" + ex.Message.ToString().Replace("'", "") + "'" + " )";
                                if (Sqlcon.State != ConnectionState.Open)
                                {
                                    Sqlcon.Open();
                                }
                                SqlCommand sqlCommand = new SqlCommand(Query, Sqlcon);
                                sqlCommand.ExecuteNonQuery();
                                Sqlcon.Close();

                            }                            
                        }
                    }
                }
                catch (Exception ex )
                {

                    try
                    {
                        SheetName = SheetName ?? "";
                        FileName = FileName ?? "";
                        Query = "insert ErrorSheet (FileName,SheetName , ErrorDesc)values (" + "N'" + FileName.ToString().Replace("'", "") + "'" + ", N" + SheetName.ToString().Replace("'", "")  + " , " + "N'" + ex.Message.ToString().Replace("'", "") + "'" + " )";
                        if (Sqlcon.State != ConnectionState.Open)
                        {
                            Sqlcon.Open();
                        }
                        SqlCommand sqlCommand = new SqlCommand(Query, Sqlcon);
                        sqlCommand.ExecuteNonQuery();
                        Sqlcon.Close();

                    }
                    catch  
                    {

                        SheetName = SheetName ?? "";
                        FileName = FileName ?? "";
                        Query = "insert ErrorSheet (FileName,SheetName , ErrorDesc)values (" + "N'" + FileName.ToString().Replace("'", "") + "'" + ", N" + SheetName.ToString().Replace("'", "") + "'" + " , " + "N'" + ex.Message.ToString().Replace("'", "") + "'" + " )";
                        if (Sqlcon.State != ConnectionState.Open)
                        {
                            Sqlcon.Open();
                        }
                        SqlCommand sqlCommand = new SqlCommand(Query, Sqlcon);
                        sqlCommand.ExecuteNonQuery();
                        Sqlcon.Close();
                    }
                }
            }
            MessageBox.Show("Success");
        }

        public void AddColumn(string ColumnName, SqlConnection sqlConnection)
        {
            string Query = null;

            Query = "Alter Table xlsss Add [" + ColumnName + "] nvarchar(4000)";

            if (sqlConnection.State != ConnectionState.Open)
            {
                sqlConnection.Open();
            }
            SqlCommand sqlCommand = new SqlCommand(Query, sqlConnection);
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }


        public void AddTable(SqlConnection sqlConnection)
        {
            string Query = null;
            string Query2 = null;

            Query = "drop table if exists  xlsss ; create table xlsss (a nvarchar(4000))";
            Query2 = "drop table if exists  ErrorSheet ;create Table ErrorSheet(FileName nvarchar(4000),SheetName nvarchar(4000),ErrorDesc nvarchar(max)) ";

            if (sqlConnection.State != ConnectionState.Open)
            {
                sqlConnection.Open();
            }

            SqlCommand sqlCommand = new SqlCommand(Query, sqlConnection);
            sqlCommand.ExecuteNonQuery();

            SqlCommand sqlCommand2 = new SqlCommand(Query2, sqlConnection);
            sqlCommand2.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public string Clean(String str)
        {
            str = str.Replace('ء', ' ');
            str = str.Replace(" ", "");
            str = str.Replace('آ', 'ا');
            str = str.Replace('أ', 'ا');
            str = str.Replace('ٳ', 'ا');
            str = str.Replace('إ', 'ا');
            str = str.Replace('ٱ', 'ا');
            str = str.Replace('ؤ', 'و');
            str = str.Replace('ي', 'ی');
            str = str.Replace('ك', 'ک');
            str = str.Replace('ى', 'ی');
            str = str.Replace('ئ', 'ی');
            str = str.Replace('ة', 'ه');
            str = str.Replace('ۀ', 'ه');
            str = str.Replace("لله", "له");
            str = str.Replace('۰', '0');
            str = str.Replace('١', '1');
            str = str.Replace('۲', '2');
            str = str.Replace('۳', '3');
            str = str.Replace('۴', '4');
            str = str.Replace('۵', '5');
            str = str.Replace('۶', '6');
            str = str.Replace('۷', '7');
            str = str.Replace('٨', '8');
            str = str.Replace('۹', '9');

            char[] chd = str.ToCharArray();

            foreach (var item in chd)
            {
                int charIndex = char.Parse(item.ToString());

                if ((charIndex > 1574 && charIndex < 1595) || (charIndex > 1600 && charIndex < 1609) || (charIndex > 64 && charIndex < 123) || (charIndex > 47 && charIndex < 58)
                    || (charIndex == 1705) || (charIndex == 1711) || (charIndex == 1740) || (charIndex == 1662) || (charIndex == 1670) || (charIndex == 1688))
                { }
                else { str = str.Replace(item.ToString(), ""); }
            }
            return str;
        }

        private DataTable TraverseExcel(DataTable dt)
        {
            DataTable dt2 = new DataTable();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt2.Columns.Add();
            }
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt2.Rows.Add();
                dt2.Rows[i][0] = dt.Columns[i].ColumnName;
            }

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    dt2.Rows[i][j + 1] = dt.Rows[j][i];
                }
            }
            return dt2;

        }

    }
}
