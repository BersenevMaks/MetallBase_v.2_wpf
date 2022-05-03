using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
using MetallBase2.Classes;

namespace MetallBase2
{
    /// <summary>
    /// Логика взаимодействия для WOrganizations.xaml
    /// </summary>
    public partial class WOrganizations : Window
    {
        private string sqlConnectionString = "";

        public WOrganizations(string connString)
        {
            InitializeComponent();
            sqlConnectionString = connString;
            Refresh();
        }

        OrganizationsViewModel organizationsViewModel = new OrganizationsViewModel();

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Refresh()
        {
            DataTable dtOrgname = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            List<string> orgs = new List<string>();
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"select [Name] from Organization";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtOrgname);
                foreach (DataRow row in dtOrgname.Rows)
                {
                    orgs.Add(row["Name"].ToString());
                }
                orgs.Sort();
                organizationsViewModel.OrgsCount = "Всего: " + orgs.Count;
                organizationsViewModel.Organizations = orgs;
                if (conn.State == ConnectionState.Open) conn.Close();
                SetDataContextes();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TreeView_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            gridSecond.DataContext = null;
            StackPanelButtons.DataContext = null;
            var name = "";
            if (sender is TreeView treeView)
            {
                name = treeView.SelectedItem as String;
            }
            if (!string.IsNullOrEmpty(name))
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                var dtOrg = new DataTable();
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();

                    string query = @"select ID_Organization as 'id', [name] as 'org', adress as 'adr', telefon as 'tel', 
Email as 'mail', [site] as 'site', INNKPP as 'inn', RasSchet as 'rs', KorSchet as 'ks', BIK as 'bik', [datePriceList] as 'priceDate'
from organization where [name]='" + name + "'";
                    SqlDataAdapter adapter;
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtOrg);
                    foreach (DataRow row in dtOrg.Rows)
                    {
                        organizationsViewModel.OrgID = row["id"].ToString();
                        organizationsViewModel.OrgName = row["org"].ToString();
                        organizationsViewModel.OrgAddress = row["adr"].ToString();
                        organizationsViewModel.OrgTel = row["tel"].ToString();
                        organizationsViewModel.OrgEmail = row["mail"].ToString();
                        organizationsViewModel.OrgSite = row["site"].ToString();
                        organizationsViewModel.OrgINN = row["inn"].ToString();
                        organizationsViewModel.OrgRSchet = row["rs"].ToString();
                        organizationsViewModel.OrgKorSchet = row["ks"].ToString();
                        organizationsViewModel.OrgBIK = row["bik"].ToString();
                        //organizationsViewModel = row["id"].ToString();
                        organizationsViewModel.OrgDatePrice = row["priceDate"].ToString();
                    }
                    organizationsViewModel.IsEnabledDelButton = true;
                    organizationsViewModel.IsEnabledSaveButton = true;
                    SetDataContextes();
                    if (conn.State == ConnectionState.Open) conn.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void BtnDelOrg_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("При удалении организии произойдет удаление всей связанной\nс данной организацией продукции\nВы действительно хотите удалить \"" +
                organizationsViewModel.OrgName + "\"?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.OK)
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    var sqlCmd = new SqlCommand("dbo.delOrg", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;

                    sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ organizationsViewModel.OrgName);

                    sqlCmd.ExecuteNonQuery();
                    Clearing();
                    Refresh();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void BtnSaveChanges_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                SqlCommand sqlCmd;

                sqlCmd = new SqlCommand("dbo.updateOrganization", conn);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("@ID_Organization", organizationsViewModel.OrgID);

                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ organizationsViewModel.OrgName);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgAddress)) organizationsViewModel.OrgAddress = " ";
                sqlCmd.Parameters.AddWithValue("@Adress", /* Значение параметра */ organizationsViewModel.OrgAddress);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgTel)) organizationsViewModel.OrgTel = " ";
                sqlCmd.Parameters.AddWithValue("@Telefon", /* Значение параметра */ organizationsViewModel.OrgTel);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgEmail)) organizationsViewModel.OrgEmail = " ";
                sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ organizationsViewModel.OrgEmail);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgSite)) organizationsViewModel.OrgSite = " ";
                sqlCmd.Parameters.AddWithValue("@Site", /* Значение параметра */ organizationsViewModel.OrgSite);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgINN)) organizationsViewModel.OrgINN = " ";
                sqlCmd.Parameters.AddWithValue("@INNKPP", /* Значение параметра */ organizationsViewModel.OrgINN);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgRSchet)) organizationsViewModel.OrgRSchet = " ";
                sqlCmd.Parameters.AddWithValue("@RasSchet", /* Значение параметра */ organizationsViewModel.OrgRSchet);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgKorSchet)) organizationsViewModel.OrgKorSchet = " ";
                sqlCmd.Parameters.AddWithValue("@KorSchet", /* Значение параметра */ organizationsViewModel.OrgKorSchet);

                if (string.IsNullOrEmpty(organizationsViewModel.OrgBIK)) organizationsViewModel.OrgBIK = " ";
                sqlCmd.Parameters.AddWithValue("@BIK", /* Значение параметра */ organizationsViewModel.OrgBIK);

                sqlCmd.Parameters.AddWithValue("@datePrice", organizationsViewModel.OrgDatePrice);

                sqlCmd.ExecuteNonQuery();
                MessageBox.Show("Измения сохранены в базе", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
                Clearing();
                SetDataContextes();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            Clearing();
            SetDataContextes();
        }
        
        private void SetDataContextes()
        {
            MainGrid.DataContext = null;
            StackPanelButtons.DataContext = null;
            gridSecond.DataContext = null;

            MainGrid.DataContext = organizationsViewModel;
            StackPanelButtons.DataContext = MainGrid.DataContext;
            gridSecond.DataContext = MainGrid.DataContext;
        }

        private void Clearing()
        {
            organizationsViewModel.OrgID = "";
            organizationsViewModel.OrgName = "";
            organizationsViewModel.OrgAddress = "";
            organizationsViewModel.OrgTel = "";
            organizationsViewModel.OrgEmail = "";
            organizationsViewModel.OrgSite = "";
            organizationsViewModel.OrgINN = "";
            organizationsViewModel.OrgRSchet = "";
            organizationsViewModel.OrgKorSchet = "";
            organizationsViewModel.OrgBIK = "";
            organizationsViewModel.OrgDatePrice = "";
            organizationsViewModel.IsEnabledDelButton = false;
            organizationsViewModel.IsEnabledSaveButton = false;
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            WAddOrganization wAddOrganization = new WAddOrganization(sqlConnectionString);
            wAddOrganization.ShowDialog();
            Refresh();
        }
    }
}
