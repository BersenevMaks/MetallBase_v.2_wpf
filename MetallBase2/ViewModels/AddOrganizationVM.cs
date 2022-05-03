using MetallBase2.Classes;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace MetallBase2.ViewModels
{
    public class AddOrganizationVM : INotifyPropertyChanged
    {
        public static string connectionString = "";
        
        public event PropertyChangedEventHandler PropertyChanged;

        public OrganizationsViewModel organizationsViewModel = new OrganizationsViewModel();
        
        private string orgID;
        private string orgName;
        private string orgAddress;
        private string orgTel;
        private string orgEmail;
        private string orgSite;
        private string orgINN;
        private string orgRSchet;
        private string orgKorSchet;
        private string orgBIK;
        private string orgDatePrice;
        private string orgCity;

        public string OrgID
        {
            get
            { return orgID; }
            set
            {
                orgID = value;
                OnPropertyChanged();
            }
        }

        public string OrgName
        {
            get
            { return orgName; }
            set
            {
                orgName = value;
                OnPropertyChanged();
            }
        }
        public string OrgAddress
        {
            get
            { return orgAddress; }
            set
            {
                orgAddress = value;
                OnPropertyChanged();
            }
        }
        public string OrgTel
        {
            get
            { return orgTel; }
            set
            {
                orgTel = value;
                OnPropertyChanged();
            }
        }
        public string OrgEmail
        {
            get
            { return orgEmail; }
            set
            {
                orgEmail = value;
                OnPropertyChanged();
            }
        }
        public string OrgSite
        {
            get
            { return orgSite; }
            set
            {
                orgSite = value;
                OnPropertyChanged();
            }
        }
        public string OrgINN
        {
            get
            { return orgINN; }
            set
            {
                orgINN = value;
                OnPropertyChanged();
            }
        }
        public string OrgRSchet
        {
            get
            { return orgRSchet; }
            set
            {
                orgRSchet = value;
                OnPropertyChanged();
            }
        }
        public string OrgKorSchet
        {
            get
            { return orgKorSchet; }
            set
            {
                orgKorSchet = value;
                OnPropertyChanged();
            }
        }
        public string OrgBIK
        {
            get
            { return orgBIK; }
            set
            {
                orgBIK = value;
                OnPropertyChanged();
            }
        }
        public string OrgDatePrice
        {
            get
            { return orgDatePrice; }
            set
            {
                orgDatePrice = value;
                OnPropertyChanged();
            }
        }
        public string OrgCity
        {
            get
            { return orgCity; }
            set
            {
                orgCity = value;
                OnPropertyChanged();
            }
        }

        private ICommand addCommand;
        public ICommand AddCommand
        {
            get
            {
                if (addCommand == null)
                {
                    addCommand = new RelayCommand(
                        o => { AddOrganization(); }
                        );
                }
                return addCommand;
            }
        }

        private void AddOrganization()
        {
            SqlConnection conn = new SqlConnection(connectionString);
            try
            {


                if (conn.State == ConnectionState.Closed) conn.Open();
                var sqlCmd = new SqlCommand("dbo.insOrg", conn);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ OrgName);

                if (OrgAddress == "") OrgAddress = " ";
                sqlCmd.Parameters.AddWithValue("@Adress", /* Значение параметра */ OrgAddress);

                if (OrgTel == "") OrgTel = " ";
                sqlCmd.Parameters.AddWithValue("@Telefon", /* Значение параметра */ OrgTel);

                if (OrgEmail == "") OrgEmail = " ";
                sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ OrgEmail);

                if (OrgSite == "") OrgSite = " ";
                sqlCmd.Parameters.AddWithValue("@Site", /* Значение параметра */ OrgSite);

                if (OrgINN == "") OrgINN = " ";
                sqlCmd.Parameters.AddWithValue("@INNKPP", /* Значение параметра */ OrgINN);

                if (OrgRSchet == "") OrgRSchet = " ";
                sqlCmd.Parameters.AddWithValue("@RasSchet", /* Значение параметра */ OrgRSchet);

                if (OrgKorSchet == "") OrgKorSchet = " ";
                sqlCmd.Parameters.AddWithValue("@KorSchet", /* Значение параметра */ OrgKorSchet);

                if (OrgBIK == "") OrgBIK = " ";
                sqlCmd.Parameters.AddWithValue("@BIK", /* Значение параметра */ OrgBIK);

                if (OrgCity == "") OrgCity = " ";
                sqlCmd.Parameters.AddWithValue("@CityName", /* Значение параметра */ OrgCity);

                sqlCmd.Parameters.AddWithValue("@datePrice", /* Значение параметра */ DateTime.Now.Year.ToString() + "." +
                    DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString());

                //sqlCmd.Parameters.AddWithValue("@CityName", /* Значение параметра */ organizationsViewModel.SelectedItem);

                sqlCmd.ExecuteNonQuery();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
                if (conn.State == ConnectionState.Open) conn.Close(); }
        }

        public void OnPropertyChanged([CallerMemberName] string param = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(param));
        }
    }
}
