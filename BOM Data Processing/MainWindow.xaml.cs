using System;
using System.Windows;
using System.Windows.Controls;

namespace BomDataProcessing
{
    public partial class MainWindow : Window
    {
        BomDataSet bomDataSet = null;
        ExcelHelper excel = new ExcelHelper();
        
        public MainWindow()
        {
            InitializeComponent();
            excel.ReadIn.LoadFileCompleted += Excel_LoadFileCompleted;
            excel.WriteOut.SaveFileCompleted += Excel_SaveFileCompleted;
            UnLockUI(false);
        }

        private void UnLockUI(bool value)
        {
            btnExcel_Export.IsEnabled = value;
            comBoxExcelSheet.IsEnabled = value;
            listBoxZone.IsEnabled = value;
            dgExcelContent.IsEnabled = value;
        }

        private void Excel_LoadFileCompleted(object sender, EventArgs e)
        {
            bomDataSet = new BomDataSet(((ExcelHelper.ReadEvent)e).newDS);
            comBoxExcelSheet.ItemsSource = null;
            comBoxExcelSheet.ItemsSource = bomDataSet.TableNames;
            comBoxExcelSheet.SelectedIndex = 0;
            UnLockUI(true);
        }

        private void Excel_SaveFileCompleted(object sender, EventArgs e)
        {
            MessageBox.Show("Data has been stored to " + excel.SavedFileName, "File Saving is OK.");
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            switch (((Button)sender).Name.Split('_')[1])
            {
                case "Load":
                    excel.ReadIn.Data();
                    break;
                case "Export":
                    string sheetName = comBoxExcelSheet.SelectedItem.ToString();
                    excel.WriteOut.Data(bomDataSet, sheetName);
                    break;
            }
        }

        private void comBoxExcelSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0)
                return;
            string sheetName = e.AddedItems[0].ToString();
            CreateCheckBoxes(sheetName);
            dgExcelContent.ItemsSource = bomDataSet.Tables[sheetName].DefaultView;
            for (int i = 0; i < dgExcelContent.Columns.Count; i++)
            {
                string columnName = bomDataSet.Tables[sheetName].Columns[i].ColumnName;
                dgExcelContent.Columns[i].Visibility = bomDataSet.GetColumnVisibility(sheetName, columnName);
            }
        }

        private void CheckBoxes_CheckChanged(object sender, RoutedEventArgs e)
        {
            CheckBox cBox = (CheckBox)sender;
            int excelColIndx = Convert.ToInt32(cBox.Tag);
            string sheetName = comBoxExcelSheet.SelectedItem.ToString();
            string columName = cBox.Content.ToString();
            SetVisibility_DataGrid(dgExcelContent.Columns[excelColIndx], cBox.IsChecked.Value);
            bomDataSet.SetColumnVisibility(sheetName, columName, cBox.IsChecked.Value);
        }

        void SetVisibility_DataGrid(DataGridColumn column, bool value)
        {
            column.Visibility = value ? Visibility.Visible : Visibility.Hidden;
        }
    }
}
