using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;

namespace BomDataProcessing
{
    partial class ExcelHelper
    {
        public class ReadEvent : EventArgs
        {
            public DataSet newDS;
        }

        public string SavedFileName
        {
            get
            { return savedExcelFileName; }
        }
        static string openedExcelFileName = "";
        static string savedExcelFileName = "";
        public Import ReadIn;
        public Export WriteOut;

        public ExcelHelper()
        {
            ReadIn = new Import();
            WriteOut = new Export();
        }

        #region OLEOB Variables For Excel File Access

        // OleDbConnection : Open BOM Excel file
        static string ConnStrPattern = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX={1};ReadOnly=0'";

        // OleDbDataAdapter : Read BOM Excel content
        static string selectCommandText = @"SELECT * FROM [{0}$A1:QQ1000]";

        enum IMEX
        {
            Export = 0, // Only accept to write data to Excel
            Import = 1, // Only accept read data from Excel
            Linked = 2  // write and read are acceptable 
        }

        #endregion

        public object InOutOLEDB(bool optInOut, DataTable table = null, string fileName = "")
        {
            object value = null;
            if (fileName.Length == 0)
                return value;

            string connStr = string.Format(ConnStrPattern, fileName, optInOut ? IMEX.Import : IMEX.Export);
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                conn.Open();
                switch (optInOut)
                {
                    case true:
                        openedExcelFileName = fileName;
                        value = LoadSheetData(conn, selectCommandText);
                        break;

                    case false:
                        savedExcelFileName = fileName;
                        CreateNewSheet(conn, table);
                        InsertRowData(conn, table);
                        value = true;
                        break;
                }
            }
            return value;
        }

        DataSet LoadSheetData(OleDbConnection conn, string selectCommandText)
        {
            DataSet newDataSet = new DataSet();
            foreach (string sheetName in LoadSheetName(conn))
            {
                using (OleDbDataAdapter dataAdatper = new OleDbDataAdapter(string.Format(selectCommandText, sheetName), conn))
                {
                    DataTable table = new DataTable();
                    dataAdatper.Fill(table);
                    table.TableName = sheetName;
                    newDataSet.Tables.Add(RemoveEmptyRows(table));
                }
            }
            return newDataSet;
        }

        List<string> LoadSheetName(OleDbConnection conn)
        {
            List<string> list = new List<string>();
            using (DataTable table = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
            {
                foreach (DataRow row in table.Rows)
                {
                    string sheetName = row["TABLE_NAME"].ToString().Split('$')[0].TrimStart('\'');
                    if (!list.Contains(sheetName))
                        list.Add(sheetName);
                }
            }
            return list;
        }

        DataTable RemoveEmptyRows(DataTable table)
        {
            DataTable newTable = table.Clone();
            bool isRowEmpty = false;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                isRowEmpty = true;
                for (int j = 0; j < table.Columns.Count && isRowEmpty; j++)
                {
                    if (table.Rows[i][j].ToString().Length > 0)
                    {
                        newTable.Rows.Add(table.Rows[i].ItemArray);
                        isRowEmpty = false;
                    }
                }
            }
            return newTable;
        }

        void CreateNewSheet(OleDbConnection conn, DataTable table)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("CREATE TABLE [{0}] (", table.TableName);
            for (int i = 0; i < table.Columns.Count; i++)
                sb.AppendFormat("[{0}] {1}{2}", table.Columns[i].ColumnName, "NVARCHAR(200)", i + 1 == table.Columns.Count ? ")" : ",");
            OleDbCommand(conn, sb.ToString());
        }

        void InsertRowData(OleDbConnection conn, DataTable table)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                sb.Clear();
                sb.AppendFormat(@"INSERT INTO [{0}] VALUES (", table.TableName);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    string cellValue = Regex.Replace(table.Rows[i][j].ToString(), "'", "-");
                    sb.AppendFormat(@"'{0}'{1}", cellValue, j + 1 == table.Columns.Count ? ")" : ",");
                }
                OleDbCommand(conn, sb.ToString());
            }
        }

        void OleDbCommand(System.Data.OleDb.OleDbConnection conn, string command)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = command;
            cmd.ExecuteNonQuery();
        }
    }
}
