using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace CSR_ToExcel
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<Test> testList = new List<Test>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenTXTButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            string txtPath = openFileDialog.FileName;

            if (txtPath != null && txtPath != "")
                using (FileStream fs = new FileStream(txtPath, FileMode.Open))
                {
                    LoadTests(txtPath, fs);
                }
            else { MessageBox.Show("Nie wybrano pliku!"); }

            foreach (Test test in testList)
                infoTextBox.Text += $"\n {test.Day} {test.OriginalFileName} \n" +
                    $"f1 = {test.FrictionResults[0]}\n" +
                    $"f2 = {test.FrictionResults[1]}\n" +
                    $"f3 = {test.FrictionResults[2]}\n" +
                    $"Avarage = {test.FrictionAvarage}";

            string pathToSave = txtPath.Substring(0, txtPath.LastIndexOf(System.IO.Path.DirectorySeparatorChar)+1);
            SaveInExcel(pathToSave, testList);

        }

        private void LoadTests(string txtPath, FileStream fs)
        {
            StreamReader reader = new StreamReader(fs, Encoding.GetEncoding("Windows-1250"));
            string dateOfTest = txtPath.Substring(txtPath.LastIndexOf(System.IO.Path.DirectorySeparatorChar) + 1);
            dateOfTest = dateOfTest.Remove(dateOfTest.IndexOf('.'));
            string tempLine = "";
            Test test;
            Regex timeRegex = new Regex(@"\d{2}:\d{2}:\d{2}");

            while (!reader.EndOfStream)
            {
                tempLine = reader.ReadLine();

                if (tempLine.Contains("****"))
                {
                    test = new Test();
                    do
                    {
                        tempLine = reader.ReadLine();

                        if (tempLine == null || tempLine == "")
                            continue;
                        if (timeRegex.IsMatch(tempLine))
                        {
                            test.Day = $"{dateOfTest}";
                            test.Time = tempLine.Substring(0, 8);
                            test.OriginalFileName = tempLine.Substring(8, tempLine.Length - 8).Trim();
                        }
                        else if (tempLine.Contains("Temperatura"))
                            test.TireTemperature = float.Parse(tempLine.Substring(tempLine.IndexOf(':') + 1).Trim());
                        else if (tempLine.Contains("Cisnienie"))
                            try { test.TirePressure = float.Parse(tempLine.Substring(tempLine.IndexOf(':') + 1).Trim()); }
                            catch { test.TirePressure = 7; }
                        else if (tempLine.Contains("RODZAJ"))
                            test.IsWaterOn = tempLine.Contains("mokro") ? true : false;
                        else if (tempLine.Contains("km/h"))
                            test.Speed = tempLine.Contains("65") ? 65 : 95;
                        else if (tempLine.Contains("ROZBIEG"))
                            test.AccelerateDistance = short.Parse(tempLine.Substring(tempLine.LastIndexOf('-') + 1).TrimEnd('m').Trim());
                        else if (tempLine.Contains("DYSTANS"))
                            test.TestDistance = short.Parse(tempLine.Substring(tempLine.LastIndexOf('-') + 1).TrimEnd('m').Trim());
                        else if (tempLine.Contains("Tercja"))
                        {
                            double[] friction = new double[3];
                            string[] tempLineSplit = tempLine.Split('-');

                            friction[0] = double.Parse(tempLineSplit[2].Substring(0, tempLineSplit[2].IndexOf('T') - 1).Trim());
                            friction[1] = double.Parse(tempLineSplit[3].Substring(0, tempLineSplit[3].IndexOf('T') - 1).Trim());
                            friction[2] = double.Parse(tempLineSplit[4].Trim());

                            test.FrictionResults = friction;
                        }

                    } while (!tempLine.Contains("----------------------------------------------------------------------"));

                    testList.Add(test);
                }
                else { continue; }

            }
        }

        private void SaveInExcel(string path, List<Test> testList)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets.Add();

            int currentRow = 1;
            int currentColumn = 1;
            
            worksheet.Cells[currentRow, currentColumn++] = "Data pomiaru";
            worksheet.Cells[currentRow, currentColumn++] = "Godzina pomiaru";
            worksheet.Cells[currentRow, currentColumn++] = "Nazwa pliku";
            worksheet.Cells[currentRow, currentColumn++] = "Temperatura koła";
            worksheet.Cells[currentRow, currentColumn++] = "Ciśnienie w kole";
            worksheet.Cells[currentRow, currentColumn++] = "Nawierzchnia";
            worksheet.Cells[currentRow, currentColumn++] = "Woda";
            worksheet.Cells[currentRow, currentColumn++] = "Prędkość";
            worksheet.Cells[currentRow, currentColumn++] = "Rozbieg";
            worksheet.Cells[currentRow, currentColumn++] = "Dystans";
            worksheet.Cells[currentRow, currentColumn++] = "CSR1";
            worksheet.Cells[currentRow, currentColumn++] = "CSR2";
            worksheet.Cells[currentRow, currentColumn++] = "CSR3";
            worksheet.Cells[currentRow, currentColumn] = "CSR średnia";

            foreach(Test test in testList)
            {
                currentRow++;
                currentColumn = 1;

                worksheet.Cells[currentRow, currentColumn++] = test.Day;
                worksheet.Cells[currentRow, currentColumn++] = test.Time;
                worksheet.Cells[currentRow, currentColumn++] = test.OriginalFileName;
                worksheet.Cells[currentRow, currentColumn++] = test.TireTemperature;
                worksheet.Cells[currentRow, currentColumn++] = test.TirePressure;
                worksheet.Cells[currentRow, currentColumn++] = test.OriginalFileName.ToLower().Contains("bc") ? "BC" : "AC";
                worksheet.Cells[currentRow, currentColumn++] = test.IsWaterOn ? "M" : "S";
                worksheet.Cells[currentRow, currentColumn++] = test.Speed;
                worksheet.Cells[currentRow, currentColumn++] = test.AccelerateDistance;
                worksheet.Cells[currentRow, currentColumn++] = test.TestDistance;
                worksheet.Cells[currentRow, currentColumn++] = test.FrictionResults[0];
                worksheet.Cells[currentRow, currentColumn++] = test.FrictionResults[1];
                worksheet.Cells[currentRow, currentColumn++] = test.FrictionResults[2];
                worksheet.Cells[currentRow, currentColumn] = test.FrictionAvarage;
            }

            path = $"{path}CSR.xlsx";
            workbook.SaveAs(path);
            workbook.Close();
            excelApp.Quit();
        }

        private void infoTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
        }
    }
}
