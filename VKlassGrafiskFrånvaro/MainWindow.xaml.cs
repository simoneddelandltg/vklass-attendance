using System;
using System.Collections.Generic;
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
            vklass = new VKlassChartCreator();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            vklass.StartChromeWindow();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            vklass.GetAbsenceDataFromClass();
        }
    }
}
