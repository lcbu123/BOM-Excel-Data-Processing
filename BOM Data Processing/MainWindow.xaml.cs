using System;
using System.Data;
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
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            switch (((Button)sender).Name.Split('_')[1])
            {
                case "Load":
                    DataSet newSet = excel.ReadIn.Data;
                    if (newSet != null)
                    {
                        bomDataSet = new BomDataSet(newSet);
                        comBoxExcelSheet.ItemsSource = null;
                        comBoxExcelSheet.ItemsSource = bomDataSet.TableNames;
                        comBoxExcelSheet.SelectedIndex = 0;
                    }
                    break;
                case "Export":
                    if (bomDataSet != null && excel.WriteOut.Data(bomDataSet.SelectVisibleTable(comBoxExcelSheet.SelectedItem.ToString())) != null)
                    {
                        MessageBox.Show("Data has been stored to " + excel.SavedFileName, "File Saving is OK.");
                    }
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
                dgExcelContent.Columns[i].Visibility = bomDataSet.ColumnVisibility[sheetName][columnName] ? Visibility.Visible : Visibility.Hidden;
            }
        }

        private void CheckBoxes_CheckChanged(object sender, RoutedEventArgs e)
        {
            CheckBox cBox = (CheckBox)sender;
            int excelColIndx = Convert.ToInt32(cBox.Tag);
            string sheetName = comBoxExcelSheet.SelectedItem.ToString();
            string columName = cBox.Content.ToString(); ;
            if (cBox.IsChecked == true)
            {
                dgExcelContent.Columns[excelColIndx].Visibility = Visibility.Visible;
                bomDataSet.ColumnVisibility[sheetName][columName] = true;
            }
            else
            {
                dgExcelContent.Columns[excelColIndx].Visibility = Visibility.Hidden;
                bomDataSet.ColumnVisibility[sheetName][columName] = false;
            }
        }
    }
}
