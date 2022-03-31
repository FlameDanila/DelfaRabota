using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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

namespace DelfaRabota
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buildButton_Click(object sender, RoutedEventArgs e)
        {
            string name = "";
            string adress = "";
            string route = "";
            string phone = "";
            string coordinate = "";
            string site = "";
            string category = "";
            int Counter = 1;
            int startRowIndex = 1;
            string numberFirms = "";
            string sity = "";
            string modifyLink = "";
            string typeOfSearch = "";
            int allCounter = 1;

            var application = new Excel.Application();

            Excel.Workbook wb = application.Workbooks.Add(Type.Missing);
            linkText.Text += "/";

            while (allCounter == Counter)
            {
                WebClient massivSsilok = new WebClient();

                if (Counter == 1)
                {
                    modifyLink = linkText.Text;
                    allCounter++;
                }
                else
                {
                    try
                    {
                        Match matches = Regex.Match(linkText.Text, @"https://2gis.ru/(\w*?)/search/(.*?)/");

                        sity = matches.Groups[1].Value;
                        typeOfSearch = matches.Groups[2].Value;

                        modifyLink = $@"https://2gis.ru/{sity}/search/{typeOfSearch}/page/{Counter}";

                        allCounter++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        allCounter--;
                    }
                }
                Counter++;

                Byte[] massivData = massivSsilok.DownloadData(modifyLink);

                using (FileStream file = new FileStream(@"Category.txt", FileMode.Create))
                {
                    Byte[] vs = massivData;

                    file.Write(vs, 0, vs.Length);
                }

                string massTextData = File.ReadAllText(@"Category.txt");

                MatchCollection match1 = Regex.Matches(massTextData, @"class=""_1h3cgic""><a href=""/(\w*?)/firm/(\d*?)"" class=""_1rehek"">");

                for (int j = 0; j < match1.Count; j++)
                {
                    Match Firms = Regex.Match(match1[j].ToString(), @"class=""_1h3cgic""><a href=""/(\w*?)/firm/(\d*?)"" class=""_1rehek"">");
                    sity = Firms.Groups[1].Value;
                    numberFirms = Firms.Groups[2].Value;

                    WebClient web = new WebClient();

                    Byte[] Data = web.DownloadData($"https://2gis.ru/{sity}/firm/{numberFirms}");

                    using (FileStream file = new FileStream(@"t.txt", FileMode.Create))
                    {
                        Byte[] vs = Data;

                        file.Write(vs, 0, vs.Length);
                    }

                    string allData = File.ReadAllText(@"t.txt");
                     

                    if (All.IsChecked == true)
                    {
                        Match match = Regex.Match(allData.ToString(), @"<h1 class=""_d9xbeh""><span class=""""><span class=""_oqoid"">(.*?)</span></span></h1>");
                        name = match.Groups[1].Value;

                        match = Regex.Match(allData.ToString(), @"address_name"":""(.*?),\s(.*?)"",""adm_div""");
                        adress = match.Groups[1].Value.ToString() + " " + match.Groups[2].Value.ToString();

                        match = Regex.Match(allData.ToString(), @"href=""tel:(.*?)"" class=");
                        phone = match.Groups[1].Value.ToString();

                        match = Regex.Match(allData.ToString(), @"<link href=""(.*?)""");
                        route = match.Groups[1].Value.ToString();

                        match = Regex.Match(allData.ToString(), @"data-divider-shifted="".*""><div class=""_14uxmys""><span><a href=""(.*?)"".* aria-label=""ВКонтакте""");
                        site = match.Groups[1].Value;
                        Match matchWA = Regex.Match(allData.ToString(), @"</span></div><div class=""_14uxmys""><span><a href=""(.*?)""(.*?) aria-label=""WhatsApp""");
                        site += "\n" + matchWA.Groups[1].Value;
                        Match matchViber = Regex.Match(allData.ToString(), @"_14uxmys""><span><a href=""(.*?)""(.*?) aria-label=""Viber""");
                        site += "\n" + matchViber.Groups[1].Value;

                        match = Regex.Match(allData.ToString(), @"default_pos"":{""lat"":(.*?),""lon"":(.*?),""zoom"":");
                        coordinate = "Долгота: " + match.Groups[1].Value + "\nШирота:" + match.Groups[2].Value;

                        match = Regex.Match(allData.ToString(), @"<div class=""_11eqcnu""><span class=""_oqoid"">(.*?)</span></div>");
                        category = match.Groups[1].Value;

                        Excel.Worksheet worksheet = application.Worksheets.Item[1];

                        worksheet.Cells[1][startRowIndex] = "Название";
                        worksheet.Cells[2][startRowIndex] = "Ссылка";
                        worksheet.Cells[3][startRowIndex] = "Телефон";
                        worksheet.Cells[4][startRowIndex] = "Адрес";
                        worksheet.Cells[5][startRowIndex] = "Сайты";
                        worksheet.Cells[6][startRowIndex] = "Категория";

                        worksheet.Cells[1][startRowIndex + 1] = name;
                        worksheet.Cells[2][startRowIndex + 1] = route;
                        worksheet.Cells[3][startRowIndex + 1] = phone;
                        worksheet.Cells[4][startRowIndex + 1] = adress;
                        worksheet.Cells[5][startRowIndex + 1] = site;
                        worksheet.Cells[6][startRowIndex + 1] = category;

                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[6][startRowIndex + 1]];

                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                        worksheet.Columns.AutoFit();

                        startRowIndex += 2;
                    }
                    else
                    {
                        Match match = Regex.Match(allData.ToString(), @"<h1 class=""_d9xbeh""><span class=""""><span class=""_oqoid"">(.*?)</span></span></h1>");
                        name = match.Groups[1].Value;

                        match = Regex.Match(allData.ToString(), @"{""lon"":(.*?),""lat"":(.*?)}");

                        Excel.Worksheet worksheet = application.Worksheets.Item[1];

                        worksheet.Cells[1][startRowIndex] = "Долгота";
                        worksheet.Cells[2][startRowIndex] = "Широта";
                        worksheet.Cells[3][startRowIndex] = "Радиус";
                        worksheet.Cells[4][startRowIndex] = "Название";

                        worksheet.Cells[1][startRowIndex + 1] = match.Groups[2].Value + ",";
                        worksheet.Cells[2][startRowIndex + 1] = match.Groups[1].Value + ",";
                        worksheet.Cells[3][startRowIndex + 1] = "500,";
                        worksheet.Cells[4][startRowIndex + 1] = name;

                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex + 1]];

                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                        worksheet.Columns.AutoFit();

                        startRowIndex += 2;
                    }
                }
                application.Visible = true;
            }
        }
    }
}
