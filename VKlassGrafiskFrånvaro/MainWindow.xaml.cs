using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
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
using System.Windows.Threading;
using VKlassAbsence;

namespace VKlassGrafiskFrånvaro
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        VKlassChartCreator vklass;
        Stopwatch sw = new Stopwatch();

        DispatcherTimer dt = new();

        string lastFilePath = "";
        static string currentVersionTag = "v1.3.0";

        public MainWindow()
        {

            InitializeComponent();

            var capWriter = new StringCaptureWriter();
            //Console.SetOut(capWriter);
            capWriter.StringWrittenTo += ConsoleUpdated;

            vklass = new VKlassChartCreator();

            SetDates();
            
        }

        private void ConsoleUpdated(object sender, string e)
        {
            InfoBlock.Text += "\n" + e;
        }

        private void SetDates()
        {
            endDate.SelectedDate = DateTime.Now;

            // Find out date of the monday before the first sunday in the previous month
            endDate.DisplayDateEnd = DateTime.Now;
            startDate.DisplayDateEnd = ISOWeek.ToDateTime(DateTime.Now.Year ,ISOWeek.GetWeekOfYear(DateTime.Now), DayOfWeek.Monday);
            endDate.DisplayDateStart = DateTime.Now - TimeSpan.FromDays(365);
            startDate.DisplayDateStart = DateTime.Now - TimeSpan.FromDays(365);

            // Selected start date
            var startSelected = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            startSelected = startSelected.AddMonths(-1);
            while (startSelected.DayOfWeek != DayOfWeek.Monday)
            {
                startSelected = startSelected.AddDays(-1);
            }
            startDate.SelectedDate = startSelected;
        }

        private async void OnLoad(object sender, RoutedEventArgs e)
        {
            infoBlockBorder.Visibility = Visibility.Collapsed;

            // Update ChromeDriver if needed
            ChromeDriverInstaller installer = new ChromeDriverInstaller();
            Progress<double> progress = new Progress<double>();


            progress.ProgressChanged += HandleChromedriverInstallProgress;

            try
            {
                var installProcess = installer.Install(null, true, progress);
                await installProcess;
                InfoBlock.Text += "\nChromeDriver uppdaterad";
                section2.IsEnabled = true;
                section2.BorderBrush = Brushes.ForestGreen;
                section1.IsEnabled = false;
                section1.BorderBrush = Brushes.Black;
            }
            catch (Exception exc)
            {
                installerLabel.Content = "Fel vid uppdatering av ChromeDriver:\n" + exc.Message;
                installerLabel.Foreground = Brushes.Red;
            }

            ChromeDriverInstaller.v115httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("vklass-attendance", "1"));
            var latestTag = await ChromeDriverInstaller.v115httpClient.GetStringAsync("https://api.github.com/repos/simoneddelandltg/vklass-attendance/releases/latest");
            dynamic latestReleaseTag = JsonConvert.DeserializeObject(latestTag);

            if (latestReleaseTag.tag_name != currentVersionTag)
            {
                newVersionBlock.Visibility = Visibility.Visible;
            }

            
 
        }

        private void HandleChromedriverInstallProgress(object? sender, double e)
        {
            InfoBlock.Text += "\n" + e;

            switch (e)
            {
                case -1:
                    installerLabel.Content = "Inga verktyg behövde installeras";
                    installerLabel.Foreground = Brushes.Green;
                    break;

                case 1:
                    installerLabel.Content = "Laddar ner verktyg...";
                    installerLabel.Foreground = Brushes.Orange;
                    break;

                case 2:
                    installerLabel.Content = "Alla verktyg redo";
                    installerLabel.Foreground = Brushes.Green;
                    break;
                default:
                    break;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            vklass.StartChromeWindow();
            InfoBlock.Text += "\nFönster startat";
            section2.IsEnabled = false;
            section2.BorderBrush = Brushes.Black;
            section4.IsEnabled = false;
            section4.BorderBrush = Brushes.Black;
            openBrowserButton.IsEnabled = false;
            section3.IsEnabled = true;
            section3.BorderBrush = Brushes.ForestGreen;
            progressTextBlock.Text = "Frånvaroöversikten har inte börjat hämtas än.";
            progressTextBlock.Foreground = Brushes.Black;
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            
            section3.IsEnabled = false;
            section3.BorderBrush = Brushes.Black;
            section4.IsEnabled = true;
            section4.BorderBrush = Brushes.ForestGreen;

            Progress<AbsenceProgress> progress = new Progress<AbsenceProgress>();
            progress.ProgressChanged += HandleAbsenceProgress;

            var start = startDate.SelectedDate;
            var end = endDate.SelectedDate;

            sw.Restart();
            dt.Interval = TimeSpan.FromMilliseconds(1000);
            dt.Tick += Dt_Tick;
            dt.Start();

            try
            {
                await Task.Run(() => vklass.GetAbsenceDataFromClassByListOverview(progress, start, end));
            }
            catch (Exception f)
            {
                MessageBox.Show("Ett fel inträffade\n" + f.Message);
                throw;
            }

            
            //await vklass.GetAbsenceDataFromClass(progress);

        }

        private void Dt_Tick(object? sender, EventArgs e)
        {
            timeElapsed.Content = "Tid: " + sw.Elapsed.ToString(@"mm\:ss");
        }

        private void HandleAbsenceProgress(object? sender, AbsenceProgress e)
        {
            InfoBlock.Text += $"\n{e.FinishedStudents}/{e.TotalStudents}";

            if (e.FinishedStudents > e.TotalStudents)
            {
                progressTextBlock.Text = "Frånvaroöversikten är färdighämtad och ligger i mappen VKlass-frånvaro på ditt skrivbord!";
                progressTextBlock.Foreground = Brushes.ForestGreen;
                section2.IsEnabled = true;
                section2.BorderBrush = Brushes.ForestGreen;
                openBrowserButton.Visibility = Visibility.Visible;
                openBrowserButton.IsEnabled = true;
                lastFilePath = e.PathToOverview;
                sw.Stop();
                dt.Stop();
                estimatedTime.Content = "Uppskattad körtid: " + sw.Elapsed.ToString(@"mm\:ss");
                timeElapsed.Content = "Tid: " + sw.Elapsed.ToString(@"mm\:ss");
                openBrowserButton.IsEnabled = true;

            }
            else
            {
                pbStatus.Value = Math.Round(100 * ((double)e.FinishedStudents) / e.TotalStudents);
                progressTextBlock.Text = $"{e.FinishedStudents}/{e.TotalStudents} elever klara.";
                if (e.FinishedStudents > 0)
                {
                    estimatedTime.Content = "Uppskattad körtid: " + ((e.TotalStudents / (double)e.FinishedStudents) * sw.Elapsed).ToString(@"mm\:ss");
                }
                
            }

        }

        private void openBrowserButton_Click(object sender, RoutedEventArgs e)
        {
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(lastFilePath)
            {
                UseShellExecute = true
            };
            p.Start();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {


            Process p = new Process()
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                }
            };
            p.Start();
            e.Handled = true;
        }
    }
}
