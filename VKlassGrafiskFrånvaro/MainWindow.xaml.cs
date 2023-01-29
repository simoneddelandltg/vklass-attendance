using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
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
using VKlassAbsence;

namespace VKlassGrafiskFrånvaro
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        VKlassChartCreator vklass;

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
            // Update ChromeDriver if needed
            ChromeDriverInstaller installer = new ChromeDriverInstaller();
            Progress<double> progress = new Progress<double>();

            progress.ProgressChanged += HandleChromedriverInstallProgress;
            var installProcess = installer.Install(null, false, progress);
            await installProcess;
            InfoBlock.Text += "\nChromeDriver uppdaterad";
            startSeleniumButton.IsEnabled = true;
        }

        private void HandleChromedriverInstallProgress(object? sender, double e)
        {
            InfoBlock.Text += "\n" + e;

            switch (e)
            {
                case -1:
                    installerLabel.Content = "Inga uppdateringar behövdes";
                    installerLabel.Foreground = Brushes.Green;
                    break;

                case 1:
                    installerLabel.Content = "Laddar ner uppdaterad version...";
                    installerLabel.Foreground = Brushes.Yellow;
                    break;

                case 2:
                    installerLabel.Content = "Uppdatering klar";
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
            startSeleniumButton.IsEnabled = false;
            selectDateAndGetAbsenceGrid.IsEnabled = true;
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //await vklass.GetAbsenceDataFromClass();
            selectDateAndGetAbsenceGrid.IsEnabled = false;

            Progress<AbsenceProgress> progress = new Progress<AbsenceProgress>();
            progress.ProgressChanged += HandleAbsenceProgress;

            await Task.Run(() => vklass.GetAbsenceDataFromClass(progress));
        }

        private void HandleAbsenceProgress(object? sender, AbsenceProgress e)
        {
            InfoBlock.Text += $"\n{e.FinishedStudents}/{e.TotalStudents}";
            pbStatus.Value = Math.Round(100 * ((double)e.FinishedStudents) / e.TotalStudents);
        }
    }
}
