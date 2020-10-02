using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System;
using System.Web;

namespace ScrapeLinkedinData
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

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var str = SearchBox.Text;
            var extractedData = GoogleExtractor.ExtractCustomSearchData(str);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Creating an instance 
            // of ExcelPackage 
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet 
            var workSheet = excel.Workbook.Worksheets.Add("$Sheet1");

            // setting the properties 
            // of the work sheet  
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties 
            // of the first row 
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "S.No";
            workSheet.Cells[1, 2].Value = "Title";
            workSheet.Cells[1, 3].Value = "Link";
            workSheet.Cells[1, 4].Value = "Snippet";

            // Inserting the article data into excel 
            // sheet by using the for each loop 
            // As we have values to the first row  
            // we will start with second row 
            int recordIndex = 2;

            foreach (var data in extractedData)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = data.Title;
                workSheet.Cells[recordIndex, 3].Value = data.Link;
                workSheet.Cells[recordIndex, 4].Value = data.Content;
                recordIndex++;
            }

            // By default, the column width is not  
            // set to auto fit for the content 
            // of the range, so we are using 
            // AutoFit() method here.  
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();
            workSheet.Column(4).AutoFit();

            // file name with .xlsx extension  
            //string p_strPath = "C:/Excel/LeadGenExtractedData.xlsx";
            var curDir = Directory.GetCurrentDirectory().Replace("\\bin\\Debug", "") + "\\resource\\LeadGenExtractData.xlsx";
            if (File.Exists(curDir))
                File.Delete(curDir);

            // Create excel file on physical disk  

            FileStream objFileStrm = File.Create(curDir);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(curDir, excel.GetAsByteArray());
            //Close Excel package 
            excel.Dispose();
            //excel.Load();

            MessageBox.Show($@"Successfully Exported to your Drive {curDir}");

        }
    }
}
