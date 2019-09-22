#region Usings

using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows;
using Extend;
using Microsoft.Win32;
using Newtonsoft.Json;

#endregion

namespace ExcelParserExt
{
    //http://www.microsoft.com/en-us/download/confirmation.aspx?id=13255
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Ctor

        public MainWindow()
        {
            InitializeComponent();
        }

        #endregion

        private void OpenFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                FilePath.Text = ofd.FileName;
                try
                {
                    ReadFile(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), $"Fuck!: {ex.Message}", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private DataSet Parse(string fileName, string sheetName)
        {
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName};Extended Properties=Excel 12.0;";

            var data = new DataSet();
            using (var con = new OleDbConnection(connectionString))
            {
                var dataTable = new DataTable();
                var query = "SELECT * FROM [{0}]".F(sheetName);
                con.Open();
                var adapter = new OleDbDataAdapter(query, con);
                adapter.Fill(dataTable);
                data.Tables.Add(dataTable);
            }

            return data;
        }

        private void ReadFile(string path)
        {
            var data = Parse(path, SheetName.Text);
            var result = (from DataRow row in data.Tables[0].Rows
                    select new
                    {
                        url = row[0],
                        url_min = row[1],
                        size = row[8],
                        sold = row[9] as String == "true",
                        de = new
                        {
                            title = row[2],
                            desc = row[3],
                            caption = row[4]
                        },
                        en = new
                        {
                            title = row[5],
                            desc = row[6],
                            caption = row[7]
                        }
                    }).Cast<object>()
                .ToList();
            var res = JsonConvert.SerializeObject(result, Formatting.Indented);
            Result.Text = res;
        }
    }
}