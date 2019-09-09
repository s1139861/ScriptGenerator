using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ScriptGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseInput_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = "*.xls";
            dlg.Filter = "Excel Worksheets|*.xls;*.xlsx;*.csv";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                InputFilePath.Text = filename;
            }
        }

        private void BrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = "*.xls";
            dlg.Filter = "Excel Worksheets|*.xls;*.xlsx;*.csv";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                OutputFilePath.Text = filename;
            }
        }

        private void GenerateScript_Click(object sender, RoutedEventArgs e)
        {
            DataSet ds = new DataSet();
            string inputfilepath = InputFilePath.Text;
            if (File.Exists(inputfilepath))
            {
                var extension = System.IO.Path.GetExtension(inputfilepath).ToLower();
                using (var stream = new FileStream(inputfilepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //判斷格式套用讀取方法
                    IExcelDataReader reader = null;
                    if (extension == ".xls")
                    {
                        UpdateResultPreviewBlock(" => XLS格式");
                        reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }
                    else if (extension == ".xlsx")
                    {
                        UpdateResultPreviewBlock(" => XLSX格式");
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else if (extension == ".csv")
                    {
                        UpdateResultPreviewBlock(" => CSV格式");
                        reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }

                    //沒有對應產生任何格式
                    if (reader == null)
                    {
                        UpdateResultPreviewBlock("未知的處理檔案：" + extension);
                    }
                    Console.WriteLine(" => 轉換中");
                    using (reader)
                    {

                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            UseColumnDataType = false,
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                //設定讀取資料時是否忽略標題
                                UseHeaderRow = false
                            }
                        });


                        //把 DataSet 顯示出來
                        var table = ds.Tables[0];
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            for (var col = 0; col < table.Columns.Count; col++)
                            {
                                string data = table.Rows[row][col].ToString();
                                UpdateResultPreviewBlock(data + ",");
                            }
                            UpdateResultPreviewBlock("");
                        }
                    }
                }
            }     
        }

        private void UpdateResultPreviewBlock(string text)
        {
            OutputResult.Text = OutputResult.Text + System.Environment.NewLine + text;
        }
    }
}
