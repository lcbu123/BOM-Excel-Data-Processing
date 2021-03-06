﻿using Microsoft.Win32;
using System;
using System.Data;

namespace BomDataProcessing
{
    partial class ExcelHelper
    {
        public class Export
        {
            public event EventHandler SaveFileCompleted;

            public Nullable<bool> Data(BomDataSet bomDataSet, string sheetName)
            {
                DataTable table = bomDataSet.GetTableWithVisibleColumns(sheetName);
                var result = (Nullable<bool>)new ExcelHelper().InOutOLEDB(false, table, OpenExcelFileDialog());
                if (result.Value)
                {
                    SaveFileCompleted(this, new EventArgs());
                }
                return result;
            }

        string OpenExcelFileDialog()
            {
                SaveFileDialog saveFile = new SaveFileDialog()
                {
                    DefaultExt = "xlsx",
                    Filter = "Excel 2010 Workbook|*.xlsx",
                    Title = "Please enter a new file name to save data of current table...",
                    FileName = string.Format("{0}_{1}{2:00}{3:00}-{4:00}{5:00}{6:00}.xlsx", openedExcelFileName.Replace(".xlsx", ""), DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second)
                };
                return saveFile.ShowDialog().Value ? saveFile.FileName : "";
            }
        }
    }
}
