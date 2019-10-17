using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = "*.json";
            dlg.Filter = "JSON|*.json|All|*.*";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                TemplateFilePath.Text = filename;
            }
        }

        private void GenerateScript_Click(object sender, RoutedEventArgs e)
        {
            DataSet ds = new DataSet();
            string templatepath = TemplateFilePath.Text;
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
                        //UpdateResultPreviewBlockWithNewLine(" => XLS格式");
                        reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }
                    else if (extension == ".xlsx")
                    {
                        //UpdateResultPreviewBlockWithNewLine(" => XLSX格式");
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else if (extension == ".csv")
                    {
                        //UpdateResultPreviewBlockWithNewLine(" => CSV格式");
                        reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }

                    //沒有對應產生任何格式
                    if (reader == null)
                    {
                        UpdateResultPreviewBlockWithNewLine("未知的處理檔案：" + extension);
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

                        UpdateResultPreviewBlockWithNewLine("");
                        //把 DataSet 顯示出來
                        /*
                        var table = ds.Tables["abc"];
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            for (var col = 0; col < table.Columns.Count; col++)
                            {
                                string data = table.Rows[row][col].ToString();
                                UpdateResultPreviewBlock(data + ",");
                            }
                            UpdateResultPreviewBlockWithNewLine("");
                        }
                        */
                    }
                }
            }
            else
            {
                UpdateResultPreviewBlockWithNewLine("Input file not found");
                return;
            }
            if (File.Exists(templatepath))
            {
                //load template
                dynamic template = JsonConvert.DeserializeObject(ReadTemplate(templatepath));
                //get excel worksheet name (after first "_" in template filename)
                string filename = System.IO.Path.GetFileNameWithoutExtension(templatepath);
                string sheetName = Regex.Match(filename, @".*?_(.*)").Groups[1].Value;
                var table = ds.Tables[sheetName];
                UpdateResultPreviewBlockWithNewLine("" + template.start);
                for (int row = 1; row < table.Rows.Count; row++)
                {
                    string action = table.Rows[row][0].ToString();
                    for (var col = 1; col < table.Columns.Count; col++)
                    {
                        string paramName = table.Rows[0][col].ToString();
                        string data = table.Rows[row][col].ToString().Trim();
                        if (data == "")
                            continue;
                        string cmdLine;
                        cmdLine = "" + template.action[action][paramName];

                        if(cmdLine=="")
                            continue;

                        UpdateResultPreviewBlockWithNewLine(cmdLine.Trim().Replace("$"+paramName,data));
                    }
                    if((row+1) < table.Rows.Count)
                        UpdateResultPreviewBlockWithNewLine(""+ template.next);
                }
                UpdateResultPreviewBlockWithNewLine("" + template.end);
            }
            else
            {
                UpdateResultPreviewBlockWithNewLine("Template file not found");
                return;
            }

        }

        private string ReadTemplate(string path)
        {
            string json;
            using (StreamReader r = new StreamReader(path))
            {
                json = r.ReadToEnd();
            }
            return json;
        }

        private void UpdateResultPreviewBlockWithNewLine(string text)
        {
            OutputResult.Text = OutputResult.Text + text + System.Environment.NewLine;
        }

        private void UpdateResultPreviewBlock(string text)
        {
            OutputResult.Text = OutputResult.Text + text;
        }
    }
}
