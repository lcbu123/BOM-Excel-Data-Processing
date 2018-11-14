using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace BomDataProcessing
{
    class BomDataSet : DataSet
    {
        List<string> _tableNames = new List<string>();
        internal List<string> TableNames
        {
            get
            {
                return _tableNames;
            }
        }

        internal Dictionary<string, Dictionary<string, bool>> ColumnVisibility = new Dictionary<string, Dictionary<string, bool>>();

        public BomDataSet(DataSet dataSet)
        {
            this.Tables.CollectionChanged += TablesCollectionChanged;
            this.Merge(dataSet);
        }

        internal List<string> ColumnNames(string tableName)
        {
            List<string> columnNames = new List<string>();
            for (int i = 0; i < Tables[tableName].Columns.Count; i++)
                columnNames.Add(Tables[tableName].Columns[i].ColumnName);
            return columnNames;
        }

        internal DataTable SelectVisibleTable(string tableName)
        {
            DataTable table = Tables[tableName].Copy();
            int columnCount = ColumnVisibility[tableName].Count;
            for (int i = columnCount - 1; i >= 0; i--)
            {
                if (ColumnVisibility[tableName][table.Columns[i].ColumnName] == false)
                {
                    table.Columns.RemoveAt(i);
                }
            }
            return table;
        }

        void TablesCollectionChanged(object sender, CollectionChangeEventArgs e)
        {
            _tableNames.Clear();
            ColumnVisibility.Clear();

            for (int i = 0; i < Tables.Count; i++)
            {
                _tableNames.Add(Tables[i].TableName);
                ColumnVisibility.Add(Tables[i].TableName, new Dictionary<string, bool>());
                for (int j = 0; j < Tables[i].Columns.Count; j++)
                {
                    ColumnVisibility[Tables[i].TableName].Add(Tables[i].Columns[j].ColumnName, true);
                }
            }
        }
    }
}
