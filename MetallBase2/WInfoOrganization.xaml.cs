using MetallBase2.Classes;
using System.Windows;
using System.Windows.Input;

namespace MetallBase2
{
    /// <summary>
    /// Логика взаимодействия для WInfoOrganization.xaml
    /// </summary>
    public partial class WInfoOrganization : Window
    {
        public WInfoOrganization(CProdDetails prodDetails)
        {
            InitializeComponent();
            Details = prodDetails;
        }

        public CProdDetails Details { get; set; } = new CProdDetails();

        private void Label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            gridMain.DataContext = Details;
        }
    }
}
