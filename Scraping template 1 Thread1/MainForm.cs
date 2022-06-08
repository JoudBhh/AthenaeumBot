using ExcelHelperExe;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Scraping_template_1_Thread1.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Formatting = Newtonsoft.Json.Formatting;

namespace Scraping_template_1_Thread1
{
    public partial class MainForm : MetroForm
    {
        private bool _logToUi = true;
        private bool _logToFile = true;
        Regex _regex = new Regex("[^a-zA-Z0-9]");
        private readonly string _path = Application.StartupPath;
        private int _maxConcurrency;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        private ChromeDriver _driver;
        public MainForm()
        {
            InitializeComponent();
        }

        private async Task MainWork()
        {
            Console.WriteLine("work Start");
            //await GetAllPaints();
            //var paintsList = new List<Paint>();
            //paintsList = JsonConvert.DeserializeObject<List<Paint>>(File.ReadAllText("paints001"));
            //var paintDetail = paintsList.SelectMany(x => x.Details).ToList();

            //int j = 0;
            //foreach (var paint in paintsList)
            //{
            //    paint.Id = j;

            //    foreach (var detail in paint.Details)
            //    {
            //        if (detail.Key == "Owner/Location")
            //            paint.Location = detail.Value;
            //        if (detail.Key == "Dates")
            //            paint.Dates = detail.Value;
            //        if (detail.Key == "Dimensions")
            //            paint.Dimensions = detail.Value;
            //        if (detail.Key == "Medium")
            //            paint.Medium = detail.Value;
            //        if (detail.Key == "Enteredby")
            //            paint.Enteredby = detail.Value;
            //        if (detail.Key == "Artistage")
            //            paint.Artistage = detail.Value;
            //    }
            //    j++;
            //}

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //using (ExcelPackage pck = new ExcelPackage())
            //{
            //    pck.Workbook.Worksheets.Add("Paints").Cells[1, 1].LoadFromCollection(paintsList, true);
            //    pck.SaveAs(new FileInfo(@"D:\PaintsListFinal.xlsx"));
            //}

            //GetUniqueInfo();


            Console.WriteLine("End Work");
            //SuccessLog("Work completed");
        }

        void GetUniqueInfo()
        {
            var paints = new List<Paint>();
            var unique = new Dictionary<string, int>();
            var paintsList = new List<Paint>();

            paintsList = JsonConvert.DeserializeObject<List<Paint>>(File.ReadAllText("paints001"));
            var paintDetail = paintsList.SelectMany(x => x.Details).ToList();

                foreach (var info in paintDetail)
                {
                    if (!unique.ContainsKey(info.Key))
                        unique.Add(info.Key, 1);
                    else
                        unique[info.Key]++;
                }
            foreach (var item in unique)
            {
                Console.WriteLine(item);
            }
        }

        private async Task GetAllPaints()
        {
            var artistList = new List<Artist>();
            var paintsList = new List<Paint>();
            var artworksUrlsList = new List<List<string>>();
            if (File.Exists("paints001"))
                paintsList = JsonConvert.DeserializeObject<List<Paint>>(File.ReadAllText("paints001"));
            artistList = JsonConvert.DeserializeObject<List<Artist>>(File.ReadAllText("ArtistInfo007"));
            var urlPaintingList = artistList.SelectMany(x => x.AllArrtworksUrls).ToList();
            var savedUrls = paintsList.Select(x => x.Url).ToHashSet();
            var remainingUrls = urlPaintingList.ToHashSet();
            remainingUrls.RemoveWhere(x => savedUrls.Contains(x));
            urlPaintingList = remainingUrls.ToList();

            var tasks = new List<Task<Paint>>();
            var taskUrls = new Dictionary<int, string>();
            var i = 0;
            do
            {
                //while there are still unprocessed items
                if (i < urlPaintingList.Count)
                {
                    var url = urlPaintingList[i];
                    var task = GetPaint(url);
                    tasks.Add(task);
                    taskUrls.Add(task.Id, url);
                    i++;
                    Display($"working on {i} / {urlPaintingList.Count}");
                    SetProgress((i) * 100 / urlPaintingList.Count);
                }

                if (urlPaintingList.Count % 1000 == 0)
                    File.WriteAllText("paints001", JsonConvert.SerializeObject(paintsList));
                //paintsList.Save();

                if (tasks.Count == 10 || i == urlPaintingList.Count)
                {
                    var t = await Task.WhenAny(tasks);
                    try
                    {
                        paintsList.Add(await t);
                    }
                    catch (Exception ex)
                    {
                        ErrorLog($"{taskUrls[t.Id]}\n{ex}");
                        File.WriteAllText("paints001", JsonConvert.SerializeObject(paintsList));
                    }

                    tasks.Remove(t);
                }
                if (tasks.Count == 0) break;
            } while (true);
            File.WriteAllText("paints001", JsonConvert.SerializeObject(paintsList));

            //for (int i = 0; i < remainingUrls.Count; i++)
            //{                
            //        string url = urlPaintingList[i];
            //        Display($"Working on {i + 1} / {urlPaintingList.Count} , total scraped {paintsList.Count}");
            //        try
            //        {
            //        //var t1 = new Task(async () => await GetPaint(url, j));
            //        paintsList.Add(await GetPaint(url, j));
            //        if (i % 1000 == 0)
            //                File.WriteAllText("paints001", JsonConvert.SerializeObject(paintsList));
            //            j++;

            //        }
            //        catch (Exception ex)
            //        {
            //            ErrorLog($"{url}\n{ex}\n{j}");
            //            File.WriteAllText("paints001", JsonConvert.SerializeObject(paintsList));
            //        }     
            //}                  
        }
        private async Task<Paint> GetPaint(string url)
        {
            string _url = url;
            string _artistUrl = "";
            string _urlImages = "";
            var _details = new Dictionary<string, string>();
            var docResp = await HttpCaller.GetDoc(url, 5);
            var infoNodes = docResp.DocumentNode.SelectNodes("//div[@id='generalInfo']//tr");
            string _title = docResp.DocumentNode?.SelectSingleNode("//div[@id='title']")?.InnerText;
            string _artistName = docResp.DocumentNode?.SelectSingleNode("//div[@id='hdrbox']//a")?.InnerText;
            string artistUrlNode = docResp.DocumentNode?.SelectSingleNode("//div[@id='hdrbox']//a")?.GetAttributeValue("href", "");
            if (artistUrlNode != null)
            { _artistUrl = "http://www.the-athenaeum.org/people/" + artistUrlNode.Replace("../people/", ""); }
            else _artistUrl = "Empty";
            string _copyright = docResp.DocumentNode?.SelectSingleNode("//div[@id='copyright']//div[1]//strong")?.InnerText;
            string imageUrlNodes = docResp.DocumentNode?.SelectSingleNode("//table[@width='100%']//a/img")?.GetAttributeValue("src", "");
            if (imageUrlNodes != null)
                _urlImages = "http://www.the-athenaeum.org/art/" + imageUrlNodes;
            else _urlImages = "Empty";
            string _tags = "";
            var tagsNodes = docResp.DocumentNode?.SelectSingleNode("//div[@id='tagsExisting']//div")?.InnerText;
            if (tagsNodes != null)
                _tags = tagsNodes;
            else _tags = "Empty";

            foreach (var node in infoNodes)
            {
                var info = node?.InnerText.Replace("\n", "").Replace(" ", String.Empty).Replace("\r", "");
                if (info.Contains(":"))
                {
                    var key = info.Substring(0, info.IndexOf(":"));
                    var value = info.Substring(info.LastIndexOf(':') + 1);
                    _details.Add(key, value);
                }
            }
            return new Paint
            {
                Title = _title,
                ArtistName = _artistName,
                ArtistUrl = _artistUrl,
                Copyright = _copyright,
                Tags = _tags,
                Url = _url,
                UrlImages = _urlImages,
                Details = _details
            };
        }

        private async Task<List<string>> GetAllUrlPaint(string artworksUrl)
        {
            var listUrls = new List<string>();
            var paint = new Paint();
            var docResp = await HttpCaller.GetDoc(artworksUrl);
            var titleNodes = docResp.DocumentNode?.SelectSingleNode("//table[@cellpadding='4']");
            if (titleNodes != null)
            {
                foreach (var node in titleNodes.SelectNodes(".//td//div/a"))
                {
                    var urlPaint = node?.GetAttributeValue("href", "").Replace("\n", "").Replace(" ", String.Empty).Replace("\r", "");
                    listUrls.Add(($"http://www.the-athenaeum.org/art/{urlPaint}").Replace(" ", ""));

                }
            }
            else
            {
                var urlNodes = docResp.DocumentNode?.SelectSingleNode("//div[@id='generalInfo']/div/a")?.GetAttributeValue("href", "");
                var urlRedirect = "http://www.the-athenaeum.org/art/detail.php?ID=" + urlNodes.Substring(urlNodes.LastIndexOf("ID") + 3);
                listUrls.Add(urlRedirect);
                Console.WriteLine("hello : " + urlRedirect);
            }
            return listUrls;
        }
        private async Task<List<Artist>> GetartistUrl(string url)
        {
            var path = @"artisturldoc";
            string fileName = path;
            string text = File.ReadAllText(fileName);
            Console.WriteLine(text);
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(text);
            var docResp = await HttpCaller.GetDoc(url, 1);
            var artistList = new List<Artist>();

            //var urlNodes = docResp.DocumentNode.SelectSingleNode("//table[@cellpadding='4']");
            var nodeNumber = docResp.DocumentNode.SelectNodes("//table[@cellpadding='4']//tr");
            var urlNodes = docResp.DocumentNode.SelectNodes("//table[@cellpadding='4']//tr//a");
            docResp.Save(@"artisturldoc");

            int _id = 0;
            foreach (var row in htmlDoc.DocumentNode.SelectNodes("//table[@cellpadding='4']//tr"))
            {

                string _name = "";
                string artistInfo = "";
                string _dateBirth = "";
                string _nationality = "";
                string _url = "";
                string _artworksUrl = "";
                string _artworksNum = "";
                var nameArt = "";
                for (int i = 1; i < row.SelectNodes(".//td").Count + 1; i++)
                {

                    switch (i)
                    {
                        case 1:
                            nameArt = row?.SelectSingleNode($".//td[{i}]")?.InnerText.Replace(" ", String.Empty); ;
                            _url = "http://www.the-athenaeum.org" + row?.SelectSingleNode($".//td[{i}]/a")?.GetAttributeValue("href", "").Replace(" ", String.Empty);
                            break;
                        case 2:
                            artistInfo = row?.SelectSingleNode($".//td[{i}]")?.InnerText;
                            break;
                        case 3:
                            _artworksNum = row?.SelectSingleNode($".//td[{i}]")?.InnerText;
                            _artworksUrl = "http://www.the-athenaeum.org" + row?.SelectSingleNode($".//td[{i}]/a")?.GetAttributeValue("href", "");
                            break;
                        default:
                            Console.WriteLine("??");
                            break;
                    }
                    if (artistInfo != "" && artistInfo.Contains(","))
                    {
                        _nationality = artistInfo.Substring(0, artistInfo.IndexOf(','));
                        _dateBirth = artistInfo.Substring(artistInfo.LastIndexOf(',') + 2);
                    }

                    if (nameArt != "" && nameArt.Contains(','))
                    {
                        var nom = nameArt.Substring(0, nameArt.IndexOf(',')).Trim();
                        var prenom = nameArt.Substring(nameArt.LastIndexOf(',')).Trim();
                        _name = (prenom + " " + nom).Replace(",", "");
                    }

                }
                if (_artworksUrl != "")
                {
                    artistList.Add(
                                   new Artist
                                   {
                                       Id = _id,
                                       Name = _name,
                                       DateBirth = _dateBirth,
                                       Nationality = _nationality,
                                       ArtworksNum = _artworksNum,
                                       ArtworksUrl = _artworksUrl,
                                       Url = _url
                                   });
                    _id++;
                }

            }

            File.WriteAllText("artistInfo2", JsonConvert.SerializeObject(artistList));
            return artistList;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            Directory.CreateDirectory("data");
            outputI.Text = _path + @"\output.xlsx";
            LoadConfig();
        }
        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_config[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_config[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _config[x.Name];
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_config[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _config.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _config.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _config = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_config, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _config = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (_logToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (_logToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
                Display(s);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
        }

        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
            _driver?.Quit();
        }
        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(outputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"xlsx file|*.xlsx",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                outputI.Text = saveFileDialog1.FileName;
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            SaveConfig();
            _logToUi = logToUII.Checked;
            _logToFile = logToFileI.Checked;
            await Task.Run(MainWork);
        }
    }
}
