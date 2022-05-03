using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Data.SqlClient;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Collections.ObjectModel;
using MetallBase2.Classes;
using System.Data;
using System.Text.RegularExpressions;
using SWF = System.Windows.Forms;
using SWI = System.Windows.Interop;
//using MetallBase2;

namespace MetallBase2
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

        string sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser; Pooling=true;";
        int NumbOfThreads = 1; //количество потоков для загрузки
        TreeViewItem lastSelectedTreeViewItem;
        TabItem lastSelectedTabItem;
        SqlConnection conn;

        public List<TabItem> _tabItems { get; set; } = new List<TabItem>();
        
        public ObservableCollection<CProdItem> Prods { get; set; } = new ObservableCollection<CProdItem>();
        //TabItem _tabAdd;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Settings();
            rbProd.IsChecked = true;

            // initialize tabItem array
            _tabItems = new List<TabItem>();

            // add a tabItem with + in header 
            //_tabAdd = new TabItem();
            //_tabAdd.Header = "+";

            //_tabItems.Add(_tabAdd);

            // add first tab
            //this.AddTabItem("Приветствие", "");

            // bind tab control
            //MainTabControl.DataContext = _tabItems;

            //MainTabControl.SelectedIndex = 0;
        }

        private TabItem AddTabItem(string Name, string Type)
        {
            int count = _tabItems.Count;

            // create new tab item
            TabItem tab = new TabItem();
            tab.Header = string.Format("{0}", Name);
            tab.Name = string.Format("tab{0}", count);
            tab.HeaderTemplate = MainTabControl.FindResource("TabHeader") as DataTemplate;

            // add controls to tab item, this case I added just a textbox
            var tivm = new TabItemViewModel();
            tivm.Name = Name;
            tivm.Type = Type;
            DataTable dtl = GetDataFromSQL(Name, Type);
            tivm.Prods = dtl;
            tivm.ComboBoxMarks = GetMarks(Name);
            tivm.ComboBoxOrgs = GetOrgs(Name);
            tivm.ComboBoxGosts = GetGosts(Name);
            tab.Content = tivm;

            // insert tab item right before the last (+) tab item
            if (count > 0)
                _tabItems.Insert(count, tab);
            else
                _tabItems.Insert(0, tab);
            return tab;
        }

        private void TabDynamic_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabItem tab = MainTabControl.SelectedItem as TabItem;

            if (tab == null) return;

            //if (tab.Equals(_tabAdd))
            //{
            //    // clear tab control binding
            //    MainTabControl.DataContext = null;

            //    TabItem newTab = this.AddTabItem("Новая вкладка","");

            //    // bind tab control
            //    MainTabControl.DataContext = _tabItems;

            //    // select newly added tab item
            //    MainTabControl.SelectedItem = newTab;
            //}
            //else
            //{
                lastSelectedTabItem = tab;
            //}

        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            string tabName = (sender as Button).CommandParameter.ToString();
            try
            {
                var item = MainTabControl.Items.Cast<TabItem>().Where(i => i.Name.Equals(tabName)).SingleOrDefault();


                if (item is TabItem tab)
                {
                    if (_tabItems.Count < 1)
                    {
                        MessageBox.Show("Нельзя закрыть последнюю вкладку.");
                    }
                    else if (MessageBox.Show(string.Format("Закрыть вкладку '{0}'?", tab.Header.ToString()),
                        "Закрыть вкладку", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        // get selected tab
                        TabItem selectedTab = MainTabControl.SelectedItem as TabItem;

                        // clear tab control binding
                        MainTabControl.DataContext = null;

                        _tabItems.Remove(tab);

                        // bind tab control
                        MainTabControl.DataContext = _tabItems;

                        // select previously selected tab. if that is removed then select first tab
                        if (selectedTab == null || selectedTab.Equals(tab))
                        {
                            if (_tabItems.Count > 0)
                                selectedTab = _tabItems[_tabItems.Count - 1];
                            else
                                selectedTab = null;
                        }

                        if (selectedTab != null) MainTabControl.SelectedItem = selectedTab;
                    }
                }
            }
            catch { }
        }

        public class TabItemViewModel
        {
            public string Name { get; set; }
            public string Type { get; set; }
            public DataTable Prods { get; set; }
            public List<ComboBoxItem> ComboBoxMarks { get; set; }
            public List<ComboBoxItem> ComboBoxOrgs { get; set; }
            public List<ComboBoxItem> ComboBoxGosts { get; set; }

            public string TxtDiamFilter { get; set; }
            public string txtTolshFilter { get; set; }
            public string combMarkFilter { get; set; }
            public string combOrganizarionFilter { get; set; }
            public string combGosts { get; set; }

            public bool ExpanderState { get; set; }

            private bool isSelected;

            public bool IsSelected
            {
                get { return isSelected; }
                set
                {
                    isSelected = value;
                    DoSomethingWhenSelected();
                }
            }
            private void DoSomethingWhenSelected()
            {
                if (isSelected) { }
                    //Debug.WriteLine("You selected " + Name);
            }

        }
        
        private DataTable GetDataFromSQL(string Name, string Type, string Diam = "", string Tolsh = "", string Mark = "", string OrgName = "", string Gost = "")
        {
            DataTable dt = new DataTable();
            string Query = "";

            if (!string.IsNullOrEmpty(Name) && string.IsNullOrEmpty(Type))
            {
                if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(Name))
                {
                    Query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + Name + @"'";
                }
                else
                {
                    Query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Размер', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + Name + @"'";
                }
            }
            else if (!string.IsNullOrEmpty(Name) && !string.IsNullOrEmpty(Type))
            {
                if (Type.ToLower().Contains("нерж")) Type = "нерж";
                if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(Name))
                {

                    Query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + Name + "'";
                    if(Type=="нерж")
                        Query += " and p.Type like '%" + Type + "%'";
                    else Query +=" and p.Type = '" + Type + "'";
                }
                else
                {
                    Query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Размер', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + Name + "'";
                    if (Type == "нерж")
                        Query += " and p.Type like '%" + Type + "%'";
                    else Query += " and p.Type = '" + Type + "'";
                }
            }
            if(!string.IsNullOrEmpty(Query))
            {
                if (!string.IsNullOrEmpty(Diam))
                    Query += " and p.diametr='" + Diam + "'";
                if (!string.IsNullOrEmpty(Tolsh))
                    Query += " and p.tolshina='" + Tolsh + "'";
                if (!string.IsNullOrEmpty(Mark))
                    Query += " and p.Marka like '%" + Mark + "%'";
                if (!string.IsNullOrEmpty(OrgName))
                    Query += " and o.name like '%" + OrgName + "%'";
                if (!string.IsNullOrEmpty(Gost))
                    Query += " and p.standart like '%" + Gost + "%'";

                Query += " order by p.diametr, p.Tolshina";
            }
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();

                SqlCommand comm;

                comm = new SqlCommand(Query, conn);

                SqlDataAdapter reader = new SqlDataAdapter(comm);
                reader.Fill(dt);
            }
            catch (Exception ex) { ex.ToString(); }
            return dt;
        }

        private void Settings()
        {
            string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\settings.set";                //пишем полный путь к файлу
            fileName = fileName.Replace("file:\\", "");
            if (File.Exists(fileName) == true)
            {
                string[] allText = { "", "", "", "", "", "" };
                try
                {                                  //чтение файла
                    allText = File.ReadAllLines(fileName);         //чтение всех строк файла в массив строк
                }
                catch (FileNotFoundException e)
                {
                    Console.WriteLine(e.Message);
                }
                if (allText.Length == 7)
                {
                    sqlConnString = "Server=tcp:" + allText[0] + "," + allText[2] + "; Database=" + allText[3] + "; User ID=" + allText[4] + "; Password=" + allText[5] + "; Pooling=true;";
                    if (!string.IsNullOrEmpty(allText[6]))
                        int.TryParse(allText[6], out NumbOfThreads);
                }
                //else
                //{
                //    settingForm sf = new settingForm(new settingDelegate(GetSettings));
                //    sf.ShowDialog();
                //    if (set.Count == 7)
                //    {
                //        sqlConnString = "Server=tcp:" + set[0] + "," + set[2] + "; Database=" + set[3] + "; User ID=" + set[4] + "; Password=" + set[5] + "; Pooling=true;";
                //        int.TryParse(set[6], out NumbOfThreads);
                //    }
                //    else sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser";
                //}
            }
            else sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser";
            //Console.ReadKey();
        }

        private void ConnectAndUpdate()
        {
            MainLabel.Content = "Выполняется подключение";

            if (rbProd.IsChecked == true)
            {
                try
                {
                    conn = new SqlConnection(sqlConnString);
                    conn.Open();

                    SqlCommand comm;

                    comm = new SqlCommand("Select distinct [Name] from dbo.Product order by [Name]", conn);

                    SqlDataReader reader = comm.ExecuteReader();

                    List<string> prods = new List<string>();
                    while (reader.Read())
                    { prods.Add(reader.GetString(0)); }
                    if (!reader.IsClosed) reader.Close();

                    ObservableCollection<CProductTreeView> products = new ObservableCollection<CProductTreeView>();
                    foreach (string p in prods)
                    {
                        ObservableCollection<CProductTypeTreeView> types = new ObservableCollection<CProductTypeTreeView>();
                        comm = new SqlCommand("Select distinct [Type] from dbo.Product where [Name]='" + p + "' order by [Type]", conn);
                        reader = comm.ExecuteReader();
                        while (reader.Read())
                        {
                            string str = reader.GetString(0);
                            str = str.Trim();
                            if (str != "")
                                types.Add(
                                    new CProductTypeTreeView
                                    {
                                        ParentName = p,
                                        TypeName = reader.GetString(0)
                                    }

                                    );
                        }
                        products.Add(
                            new CProductTreeView
                            {
                                Name = p,
                                Types = types
                            }
                            );
                        if (!reader.IsClosed) reader.Close();
                    }
                    if (!reader.IsClosed) reader.Close();
                    if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                    TreeViewMain.ItemsSource = products;

                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                    MainLabel.Content = "Не подключено";
                }
                if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                MainLabel.Content = "Подключено";
            }
            else if (rbType.IsChecked == true)
            {
                try
                {
                    conn = new SqlConnection(sqlConnString);
                    conn.Open();

                    SqlCommand comm;

                    List<string> typs = new List<string>()
                    {
                        "Нержавейка",
                        "Алюминий",
                        "Бронза"
                    };
                    ObservableCollection<CProductTreeView> products = new ObservableCollection<CProductTreeView>();
                    foreach (string t in typs)
                    {
                        var queryType = t;
                        if (t.ToLower().Contains("нерж")) queryType = "нерж";
                        ObservableCollection<CProductTypeTreeView> types = new ObservableCollection<CProductTypeTreeView>();
                        comm = new SqlCommand("Select distinct[Name] from dbo.Product where [Type] like '%" + queryType +"%' order by[Name]", conn);
                        SqlDataReader reader = comm.ExecuteReader();
                        while(reader.Read())
                        {
                            string str = reader.GetString(0);
                            str = str.Trim();
                            if (str != "")
                                types.Add(
                                    new CProductTypeTreeView
                                    {
                                        ParentName = t,
                                        TypeName = str
                                    }
                                    );
                        }
                        if (!reader.IsClosed) reader.Close();
                        products.Add(
                            new CProductTreeView
                            {
                                Name = t,
                                Types = types
                            }
                            );
                    }
                    if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                    TreeViewMain.ItemsSource = products;

                    conn.Close();
                    MainLabel.Content = "Подключено";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    if (conn.State == System.Data.ConnectionState.Open) conn.Close();
                    MainLabel.Content = "Не подключено";
                }
                MainLabel.Content = "Подключено";
            }
        }

        private void TreeViewMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            
        }

        private void TreeViewItemSelected(object sender, RoutedEventArgs e)
        {
            TreeViewItem tvi = e.OriginalSource as TreeViewItem;
            this.lastSelectedTreeViewItem = tvi;

        }

        private void TreeViewMain_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            //отмена встроенного алгоритма обработки двойных кликов
            //чтобы не разворачивалась ветка, а выполнялся мой код
            e.Handled = true; 
            string Name = "";
            string Type = "";

            if (rbProd.IsChecked == true)
            {
                if (((sender as TreeView).SelectedItem as CProductTreeView) != null)
                {
                    Name = ((sender as TreeView).SelectedItem as CProductTreeView).Name;
                    Type = "";
                }
                else if (((sender as TreeView).SelectedItem as CProductTypeTreeView) != null)
                {
                    Name = ((sender as TreeView).SelectedItem as CProductTypeTreeView).ParentName;
                    Type = ((sender as TreeView).SelectedItem as CProductTypeTreeView).TypeName;
                }
            }
            else
            {
                if (((sender as TreeView).SelectedItem as CProductTypeTreeView) != null)
                {
                    Name = ((sender as TreeView).SelectedItem as CProductTypeTreeView).TypeName;
                    Type = ((sender as TreeView).SelectedItem as CProductTypeTreeView).ParentName;
                }
                else if (((sender as TreeView).SelectedItem as CProductTreeView) != null)
                {
                    Name = "";
                    Type = ((sender as TreeView).SelectedItem as CProductTreeView).Name;
                }
            }

            if (!string.IsNullOrEmpty(Name))
            {
                // clear tab control binding
                MainTabControl.DataContext = null;

                TabItem newTab = this.AddTabItem(Name, Type);

                // bind tab control
                MainTabControl.DataContext = _tabItems;

                // select newly added tab item
                MainTabControl.SelectedItem = newTab;

            }
            this.Cursor = Cursors.Arrow;
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            ConnectAndUpdate();
        }

        private void WindowCloseButton(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Row_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            DataGridRow row = sender as DataGridRow;
            WInfoOrganization wInfo = new WInfoOrganization(GetDetails((row.Item as DataRowView)[1].ToString()));
            wInfo.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            wInfo.ShowDialog();
        }

        private CProdDetails GetDetails(string NameOrganization)
        {
            CProdDetails cProdDetails = new CProdDetails();
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();
                SqlCommand comm;
                comm = new SqlCommand("Select City,Email,Telefon from Organization where [Name]='" + NameOrganization + "'  order by [City]", conn);
                SqlDataReader reader = comm.ExecuteReader();
                    cProdDetails.OrgName = NameOrganization;
                reader.Read();
                cProdDetails.City = reader.GetString(0).ToString() ?? string.Empty;
                cProdDetails.Email = reader.GetString(1).ToString() ?? string.Empty;
                cProdDetails.Telephone = reader.GetString(2).ToString() ?? string.Empty;
                if (!reader.IsClosed) reader.Close();
                comm = new SqlCommand("Select m.[Name] as 'Имя', m.Telefon as 'Телефон' from Manager m where ID_Organization = (select ID_Organization from Organization where [Name] = '" + NameOrganization + "');", conn);
                SqlDataAdapter sda = new SqlDataAdapter(comm);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                cProdDetails.Managers = dt;
                if (!reader.IsClosed) reader.Close();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

            return cProdDetails;
        }

        private List<ComboBoxItem> GetMarks(string Name)
        {
            List<ComboBoxItem> lcbi = new List<ComboBoxItem>();
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();
                SqlCommand comm;
                comm = new SqlCommand("Select distinct [Marka] from dbo.Product where [Name]='" + Name + "'  order by [Marka]", conn);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    lcbi.Add(new ComboBoxItem() { Content = reader.GetString(0) });
                }
                if (!reader.IsClosed) reader.Close();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch(Exception ex) { MessageBox.Show(ex.ToString()); }

            return lcbi;
        }

        private List<ComboBoxItem> GetOrgs(string Name)
        {
            List<ComboBoxItem> lcbi = new List<ComboBoxItem>();
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();
                SqlCommand comm;
                comm = new SqlCommand("Select distinct [Name] from dbo.Organization where id_Organization in (select id_Organization from dbo.Product where [Name]='" + Name + "')  order by [Name]", conn);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    lcbi.Add(new ComboBoxItem() { Content = reader.GetString(0) });
                }
                if (!reader.IsClosed) reader.Close();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

            return lcbi;
        }

        private List<ComboBoxItem> GetGosts(string Name)
        {
            List<ComboBoxItem> lcbi = new List<ComboBoxItem>();
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();
                SqlCommand comm;
                comm = new SqlCommand("Select distinct standart from dbo.Product where [Name]='" + Name + "' order by [standart]", conn);
                SqlDataReader reader = comm.ExecuteReader();
                while (reader.Read())
                {
                    lcbi.Add(new ComboBoxItem() { Content = reader.GetString(0) });
                }
                if (!reader.IsClosed) reader.Close();
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

            return lcbi;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            ImportMetallBase form = new ImportMetallBase(sqlConnString);
            SWI.WindowInteropHelper wih = new SWI.WindowInteropHelper(this);
            //wih.Owner = form.Handle;
            form.ShowDialog();
        }

        private void BtnApplayFilter_Click(object sender, RoutedEventArgs e)
        {
            // clear tab control binding
            
            int index = Convert.ToInt32(new Regex(@"(?<=tab)\d+", RegexOptions.IgnoreCase).Match(lastSelectedTabItem.Name).Value);
            if (index < 0) index = 0;
            TabItemViewModel tivm = _tabItems[index].Content as TabItemViewModel;
            TabItem tab = _tabItems[index];
            MainTabControl.DataContext = null;
            int count = _tabItems.Count;

            // create new tab item
            
            //tab.Header = string.Format("{0}", tivm.Name);
            //tab.Name = string.Format("tab{0}", count);
            //tab.HeaderTemplate = MainTabControl.FindResource("TabHeader") as DataTemplate;

            tivm.Prods = GetDataFromSQL(tivm.Name, tivm.Type, tivm.TxtDiamFilter, tivm.txtTolshFilter, tivm.combMarkFilter, tivm.combOrganizarionFilter,
                tivm.combGosts);
            //tivm.ComboBoxMarks = GetMarks(tivm.Name);
            //tivm.ComboBoxOrgs = GetOrgs(tivm.Name);
            //tivm.ComboBoxGosts = GetGosts(tivm.Name);
            tivm.ExpanderState = true;
            tab.Content = tivm;

            // insert tab item right before the last (+) tab item
            index = Convert.ToInt32(new Regex(@"(?<=tab)\d+", RegexOptions.IgnoreCase).Match(lastSelectedTabItem.Name).Value);
            if (index < 0) index = 0;
            _tabItems[index] = tab;

            // bind tab control
            MainTabControl.DataContext = _tabItems;

            // select newly added tab item
            MainTabControl.SelectedItem = tab;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //TabItemViewModel tivm = lastSelectedTabItem.Content as TabItemViewModel;
            //tivm.TxtDiamFilter = (sender as TextBox).Text;
            //int index = Convert.ToInt32(new Regex(@"(?<=tab)\d+", RegexOptions.IgnoreCase).Match(lastSelectedTabItem.Name).Value)-1;
            //if (index < 0) index = 0;
            //_tabItems[index].Content = tivm;
            ////MainTabControl.DataContext = _tabItems;
        }

        private void BtnClearFilter_Click(object sender, RoutedEventArgs e)
        {
            int index = Convert.ToInt32(new Regex(@"(?<=tab)\d+", RegexOptions.IgnoreCase).Match(lastSelectedTabItem.Name).Value);
            if (index < 0) index = 0;
            TabItemViewModel tivm = _tabItems[index].Content as TabItemViewModel;
            TabItem tab = _tabItems[index];
            MainTabControl.DataContext = null;
            
            tivm.TxtDiamFilter = string.Empty;
            tivm.txtTolshFilter = string.Empty;
            tivm.combGosts = string.Empty;
            tivm.combMarkFilter = string.Empty;
            tivm.combOrganizarionFilter = string.Empty;
            tivm.ExpanderState = true;
            tivm.Prods = tivm.Prods = GetDataFromSQL(tivm.Name, tivm.Type);
            tab.Content = tivm;
            _tabItems[index] = tab;
            MainTabControl.DataContext = _tabItems;
            MainTabControl.SelectedItem = tab;
        }

        private void MISettingsClick(object sender, RoutedEventArgs e)
        {
            WSettings wSettings = new WSettings();
            wSettings.ShowDialog();
            Settings();
        }

        private void MIOrganizations_Click(object sender, RoutedEventArgs e)
        {
            WOrganizations wOrganizations = new WOrganizations(sqlConnString);
            wOrganizations.ShowDialog();
        }

        private void MINerg_Click(object sender, RoutedEventArgs e)
        {
            RedactorMarks redactorMarks = new RedactorMarks(sqlConnString, "Nerj");
            redactorMarks.ShowDialog();
        }

        private void MIAlumin_Click(object sender, RoutedEventArgs e)
        {
            RedactorMarks redactorMarks = new RedactorMarks(sqlConnString, "Alumin");
            redactorMarks.ShowDialog();
        }

        private void MIMed_Click(object sender, RoutedEventArgs e)
        {
            RedactorMarks redactorMarks = new RedactorMarks(sqlConnString, "Med");
            redactorMarks.ShowDialog();
        }

        private void MILatun_Click(object sender, RoutedEventArgs e)
        {
            RedactorMarks redactorMarks = new RedactorMarks(sqlConnString, "Latun");
            redactorMarks.ShowDialog();
        }

        private void MIListErrors_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
