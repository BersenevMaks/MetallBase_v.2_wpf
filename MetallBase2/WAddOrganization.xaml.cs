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
using System.Windows.Shapes;
using MetallBase2.ViewModels;

namespace MetallBase2
{
    /// <summary>
    /// Логика взаимодействия для WAddOrganization.xaml
    /// </summary>
    public partial class WAddOrganization : Window
    {
        public WAddOrganization(string sqlConnectionString)
        {
            InitializeComponent();
            AddOrganizationVM.connectionString = sqlConnectionString;
        }
        AddOrganizationVM addOrganizationVM = new AddOrganizationVM();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(addOrganizationVM.organizationsViewModel.OrgAddress);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
