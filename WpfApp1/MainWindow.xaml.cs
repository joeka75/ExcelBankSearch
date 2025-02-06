using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Text.RegularExpressions;

namespace WpfApp1
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

        private void OnBrowseClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel and PDF Files|*.xls;*.xlsx;*.xlsm;*.pdf"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FilePath.Text = openFileDialog.FileName;
            }
        }

        private void OnSearchClick(object sender, RoutedEventArgs e)
        {
            string filePath = FilePath.Text;
            string searchValue = SearchValue.Text;

            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(searchValue))
            {
                MessageBox.Show("Please provide both the file path and the search value.");
                return;
            }

            try
            {
                string extension = System.IO.Path.GetExtension(filePath).ToLower();
                if (extension == ".xls" || extension == ".xlsx" || extension == ".xlsm")
                {
                    SearchInExcel(filePath, searchValue);
                }
                else if (extension == ".pdf")
                {
                    SearchInPdf(filePath, searchValue);
                }
                else
                {
                    MessageBox.Show("Unsupported file format.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void SearchInExcel(string filePath, string searchValue)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    ResultTextBox.Text = "No worksheet found in the Excel file.";
                    return;
                }

                bool found = false;
                var results = new StringBuilder();

                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (worksheet.Cells[row, col].Text == searchValue)
                        {
                            var rowContent = string.Join(", ", worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Select(cell => cell.Text));
                            results.AppendLine($"Found in row {row}: {rowContent}");
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    ResultTextBox.Text = results.ToString();
                }
                else
                {
                    ResultTextBox.Text = $"Value '{searchValue}' not found in the Excel file.";
                }
            }
        }
        private void SearchInPdf(string filePath, string searchValue)
        {
            using (PdfReader pdfReader = new PdfReader(filePath))
            using (PdfDocument pdfDoc = new PdfDocument(pdfReader))
            {
                bool found = false;
                var results = new StringBuilder();
                decimal totalSum = 0;

                // Regex a HUF összegekhez
                var hufRegex = new Regex(@"-\d+(\.\d{1,2})?(?=\sHUF|\s|$)");

                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var text = PdfTextExtractor.GetTextFromPage(page);

                    var lines = text.Split('\n');

                    foreach (var (index, line) in lines.Select((line, index) => (index, line)))
                    {
                        if (line.Contains(searchValue.ToUpper()))
                        {
                            int linesNum = index;
                            string prev = lines[linesNum - 1];
                            string fullText = prev + "\n" + line.Trim();

                            // HUF összegek keresése az adott sorban
                            Match match = hufRegex.Match(fullText);
                            if (match.Success)
                            {
                                string numberStr = match.Value.Split(' ')[0].TrimStart('-').Replace(".", ",");
                                if (decimal.TryParse(numberStr, out decimal amount))
                                {
                                    totalSum += amount;
                                }
                            }
                            //foreach (Match match in matches)
                            //{
                            //    // A HUF előtti szám kinyerése és összegzése
                            //    string numberStr = match.Value.Split(' ')[0].Remove(0,1).Replace(".",",");
                            //    if (decimal.TryParse(numberStr, out decimal amount))
                            //    {
                            //        totalSum += amount;
                            //    }
                            //}

                            results.AppendLine($"Found on page {i}: {fullText}");
                            found = true;
                        }
                    }
                }

                if (found)
                {
                    results.AppendLine($"\nÖsszesen: {totalSum} HUF");
                    ResultTextBox.Text = results.ToString();
                }
                else
                {
                    ResultTextBox.Text = $"Value '{searchValue}' not found in the PDF file.";
                }
            }
        }
    }
}
        //    private void SearchInPdf(string filePath, string searchValue)
        //    {
        //        using (PdfReader pdfReader = new PdfReader(filePath))
        //        using (PdfDocument pdfDoc = new PdfDocument(pdfReader))
        //        {
        //            bool found = false;
        //            var results = new StringBuilder();

        //            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
        //            {
        //                var page = pdfDoc.GetPage(i);
        //                var text = PdfTextExtractor.GetTextFromPage(page);

        //                var lines = text.Split('\n');

        //                foreach (var (index, line) in lines.Select((line, index) => (index, line)))
        //                {

        //                    if (line.Contains(searchValue))
        //                    {
        //                        int linesNum = index;
        //                        string prev = lines[linesNum - 1];
        //                        results.AppendLine($"Found on page {i}: {prev + "\n" + line.Trim()}");
        //                        found = true;
        //                    }
        //                }
        //            }

        //            if (found)
        //            {
        //                ResultTextBox.Text = results.ToString();
        //            }
        //            else
        //            {
        //                ResultTextBox.Text = $"Value '{searchValue}' not found in the PDF file.";
        //            }
        //        }
        //    }
        //}
    
