using Microsoft.Win32;
using System;
using System.Data;

namespace BomDataProcessing
{
    partial class ExcelHelper
    {
        public class Import
        {
            public event EventHandler LoadFileCompleted;

            public void Data()
            {
                var result = (DataSet)new ExcelHelper().InOutOLEDB(true, null, OpenExcelFileDialog());
                if (result != null)
                {
                    ReadEvent e = new ReadEvent();
                    e.newDS = result;
                    LoadFileCompleted(this, e);
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
