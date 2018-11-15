using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows;

namespace BomDataProcessing
{
    class BomDataSet : DataSet
    {
        public BomDataSet(DataSet dataSet)
        {
            this._tableNames = new List<string>();
            this._columnVisibility = new Dictionary<string, Dictionary<string, bool>>();
            this.Tables.CollectionChanged += TablesCollectionChanged;
            this.Merge(dataSet);
        }

        List<string> _tableNames;
        internal List<string> TableNames
        {
            get
            {
                return _tableNames;
            }
        }

        Dictionary<string, Dictionary<string, bool>> _columnVisibility;

        internal void SetColumnVisibility(string sheetName, string columnName, bool value)
        {
            _columnVisibility[sheetName][columnName] = value;
        }

        internal bool IsColumnVisible(string sheetName, string columnName)
        {
            return _columnVisibility[sheetName][columnName];
        }

        internal Visibility GetColumnVisibility(string sheetName, string columnName)
        {
            return _columnVisibility[sheetName][columnName] ? Visibility.Visible : Visibility.Hidden;
        }

        internal List<string> ColumnNames(string tableName)
        {
            List<string> columnNames = new List<string>();
            for (int i = 0; i < Tables[tableName].Columns.Count; i++)
                columnNames.Add(Tables[tableName].Columns[i].ColumnName);
            return columnNames;
        }

        internal DataTable GetTableWithVisibleColumns(string tableName)
        {
            DataTable table = Tables[tableName].Copy();
            int columnCount = _columnVisibility[tableName].Count;
            for (int i = columnCount - 1; i >= 0; i--)
            {
                if (_columnVisibility[tableName][table.Columns[i].ColumnName] == false)
                {
                    table.Columns.RemoveAt(i);
                }
            }
            return table;
        }

        void TablesCollectionChanged(object sender, CollectionChangeEventArgs e)
        {
            _tableNames.Clear();
            _columnVisibility.Clear();

            for (int i = 0; i < Tables.Count; i++)
            {
                _tableNames.Add(Tables[i].TableName);
                _columnVisibility.Add(Tables[i].TableName, new Dictionary<string, bool>());
                for (int j = 0; j < Tables[i].Columns.Count; j++)
                {
                    _columnVisibility[Tables[i].TableName].Add(Tables[i].Columns[j].ColumnName, true);
                }
            }
        }
    }
}
