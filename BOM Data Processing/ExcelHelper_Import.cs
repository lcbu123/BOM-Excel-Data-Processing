using Microsoft.Win32;
using System.Data;

namespace BomDataProcessing
{
    partial class ExcelHelper
    {
        public class Import
        {
            public DataSet Data
            {
                get
                {
                    return (DataSet)new ExcelHelper().InOutOLEDB(true, null, OpenExcelFileDialog());
                }
            }

            string OpenExcelFileDialog()
            {
                OpenFileDialog openFile = new OpenFileDialog()
                {
                    DefaultExt = "xlsx",
                    Filter = "Excel 2010 Workbook|*.xlsx",
                    Title = "Open Excel File"
                };
                return openFile.ShowDialog().Value ? openFile.FileName : "";
            }
        }       
    }
}
