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
                    // Automatikus formátumfelismerés
                    if (IsNewFormat(filePath))
                    {
                        SearchInPdf_V2(filePath, searchValue);
                    }
                    else
                    {
                        SearchInPdf(filePath, searchValue);
                    }
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

                // Regex az összegek keresésére CSAK a sor végéről (pl. -3 528 vagy -3,528.00)
                var hufRegex = new Regex(@"(-\s?\d{1,3}(?:[.\s]\d{3})*(?:,\d{1,2})?)\s*$", RegexOptions.Compiled);

                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var text = PdfTextExtractor.GetTextFromPage(page);

                    var lines = text.Split('\n');

                    foreach (var (index, line) in lines.Select((line, index) => (index, line)))
                    {
                        if (line.ToUpper().Contains(searchValue.ToUpper()))
                        {
                            int linesNum = index;
                            string prev = lines[Math.Max(0, linesNum - 1)];
                            string transactionBlock = prev + "\n" + line.Trim();

                            // Keresünk egy összeget, de CSAK a sor végéről
                            Match match = hufRegex.Match(line);
                            if (match.Success)
                            {
                                string numberStr = match.Groups[1].Value.Replace(" ", "").Replace(".", "").Replace(",", ".");

                                if (decimal.TryParse(numberStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal amount))
                                {
                                    totalSum += amount; // Megőrizzük az előjelet
                                }
                            }

                            results.AppendLine($"Found on page {i}: {transactionBlock}");
                            found = true;
                        }
                    }
                }

                // Összeg kiírása CSAK EGYSZER a ciklus végén
                if (found)
                {
                    results.AppendLine($"\nÖsszesen: {totalSum} HUF");
                }
                else
                {
                    results.AppendLine($"Value '{searchValue}' not found in the PDF file.");
                }

                ResultTextBox.Text = results.ToString();
            }
        }




        // **Új verzió a jelenlegi fájlformátum kezelésére**
        private void SearchInPdf_V2(string filePath, string searchValue)
        {
            using (PdfReader pdfReader = new PdfReader(filePath))
            using (PdfDocument pdfDoc = new PdfDocument(pdfReader))
            {
                bool found = false;
                var results = new StringBuilder();
                decimal totalSum = 0;

                // Regex a HUF összegek keresésére
                var hufRegex = new Regex(@"-\s?\d{1,3}(?:[.\s]\d{3})*(?:,\d{1,2})?", RegexOptions.Compiled);

                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var text = PdfTextExtractor.GetTextFromPage(page);

                    var lines = text.Split('\n');

                    foreach (var (index, line) in lines.Select((line, index) => (index, line)))
                    {
                        if (line.ToUpper().Contains(searchValue.ToUpper()))
                        {
                            int linesNum = index;
                            StringBuilder transactionBlock = new StringBuilder();
                            transactionBlock.AppendLine(line.Trim().ToUpper());

                            bool foundAmount = false;
                            string transactionAmount = "";

                            // A vásárlás fő sorának keresése (ahol az összeg szerepel)
                            int startLineIndex = linesNum;
                            while (startLineIndex > 0 && !lines[startLineIndex].Contains("Vásárlás belföldi"))
                            {
                                startLineIndex--; // Visszafelé keresünk a tranzakció fejlécéig
                            }

                            // Ellenőrizzük a forgalom oszlop értékét a fő tranzakciós sorban
                            if (startLineIndex >= 0)
                            {
                                string forgalomLine = lines[startLineIndex].Trim();

                                // Új regex az egész sor feldolgozására
                                Match match = Regex.Match(forgalomLine, @"-\s?\d{1,3}(?:[.\s]\d{3})*(?:,\d{1,2})?");

                                if (match.Success)
                                {
                                    transactionAmount = match.Value.Trim();
                                    foundAmount = true;
                                }
                            }

                            // Ha találtunk forgalmi összeget, akkor feldolgozzuk
                            if (foundAmount)
                            {
                                string numberStr = transactionAmount.Replace(" ", "").Replace(".", "").Replace(",", ".").TrimStart('-');
                                if (decimal.TryParse(numberStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal amount))
                                {
                                    totalSum += amount;
                                }
                            }

                            // Kimenet formázása, ha nincs HUF jelzés, akkor hozzáadjuk
                            string formattedAmount = foundAmount ? $"{transactionAmount} HUF" : "N/A";

                            transactionBlock.AppendLine($"Forgalom: {formattedAmount}");
                            results.AppendLine($"Found on page {i}: {transactionBlock.ToString().Trim()}");

                            found = true;
                        }

                       
                    }
                   
                }
                if (found)
                {
                    results.AppendLine($"\nÖsszesen: {totalSum} HUF");
                }
                else
                {
                    results.AppendLine($"Value '{searchValue}' not found in the PDF file.");
                }

                ResultTextBox.Text = results.ToString();
            }
        }

        // **Automatikus PDF formátum felismerés**
        private bool IsNewFormat(string filePath)
        {
            using (PdfReader pdfReader = new PdfReader(filePath))
            using (PdfDocument pdfDoc = new PdfDocument(pdfReader))
            {
                var firstPageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(1));

                // Ha van olyan minta, amely a jelenlegi formátumra utal
                return firstPageText.Contains("ÖSSZESÍTETT BANKSZÁMLAKIVONAT") || firstPageText.Contains("BANKSZÁMLAKIVONAT");
            }
        }
    }
}
