#region Usings

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Documents;
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

        void OpenFileButton_OnClick( Object sender, RoutedEventArgs e )
        {
            var ofd = new OpenFileDialog();
            if ( ofd.ShowDialog() == true )
            {
                FilePath.Text = ofd.FileName;
                ReadFile( ofd.FileName );
            }
        }

        private void ReadFile( String path )
        {
            var data = Helper.Parse( path, SheetName.Text );
            var result = new List<Object>();
            foreach (DataRow row in data.Tables[0].Rows )
            {
                var dataObj = new
                {
                    url = row[0],
                    url_min = row[1],
                    size = row[8],
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
                    },
                };
                result.Add( dataObj );
            }
            var res = JsonConvert.SerializeObject(result, Formatting.Indented);
            Result.Text = res;
        }
        }

    public static class Helper
    {
        public static String[] GetExcelSheetNames( String connectionString )
        {
            OleDbConnection con = null;
            DataTable dt = null;
            con = new OleDbConnection( connectionString );
            con.Open();
            dt = con.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );

            if ( dt == null )
                return null;

            String[] excelSheetNames = new String[dt.Rows.Count];
            Int32 i = 0;

            foreach ( DataRow row in dt.Rows )
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            return excelSheetNames;
        }

        public static DataSet Parse( String fileName, String sheetName )
        {
            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;";

            var data = new DataSet();
            using ( var con = new OleDbConnection( connectionString ) )
            {
                var dataTable = new DataTable();
                var query = "SELECT * FROM [{0}]".F( sheetName );
                con.Open();
                var adapter = new OleDbDataAdapter( query, con );
                adapter.Fill( dataTable );
                data.Tables.Add( dataTable );
            }

            return data;
        }
    }
}