using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace BomDataProcessing
{
    public partial class MainWindow : Window
    {
        #region Create CheckBoxes Dynamically for mapping Excel Column Name and Index

        public ObservableCollection<ExcelColumnsInfo> CheckBoxesForExcelColumnList { get; set; }

        public class ExcelColumnsInfo
        {
            public string ExcelColumnName { get; set; }
            public int ExcelColumnIndex { get; set; }
            public bool CheckBoxIsChecked { get; set; }
        }

        public void Initial_ExcelColumnList()
        {
            if (CheckBoxesForExcelColumnList != null)
                CheckBoxesForExcelColumnList.Clear();
            else
            {
                CheckBoxesForExcelColumnList = new ObservableCollection<ExcelColumnsInfo>();
                this.DataContext = this;
            }
        }

        public void CreateCheckBoxes(string sheetName)
        {
            Initial_ExcelColumnList();

            List<string> columnNames = bomDataSet.ColumnNames(sheetName);
            for (int i = 0; i < columnNames.Count; i++)
            {
                CheckBoxesForExcelColumnList.Add(new ExcelColumnsInfo
                {
                    ExcelColumnName = columnNames[i],
                    ExcelColumnIndex = i,
                    CheckBoxIsChecked = bomDataSet.IsColumnVisible(sheetName, columnNames[i])
                });
            }
        }

        #endregion
    }
}
