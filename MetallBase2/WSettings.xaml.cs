using System.Windows;
using System.IO;
using System.Reflection;
using System.Windows.Input;

namespace MetallBase2
{
    /// <summary>
    /// Логика взаимодействия для WSettings.xaml
    /// </summary>
    public partial class WSettings : Window
    {
        public WSettings()
        {
            InitializeComponent();
            mainGridView = new MainGridView();
            fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\settings.set";                //пишем полный путь к файлу
            GetParams();
            MainGrid.DataContext = mainGridView;
        }

        string fileName = "";
        MainGridView mainGridView;

        internal class MainGridView
        {
            public string ServName { get; set; }
            public string InstName { get; set; }
            public string PortNumb { get; set; }
            public string DataBase { get; set; }
            public string UserID { get; set; }
            public string Password { get; set; }
            public string NumbStreamReader { get; set; }
            public ICommand comCancel { get; set; }
        }

        private void GetParams()
        {
            fileName = fileName.Replace("file:\\", "");
            if (File.Exists(fileName) == true)
            {
                try
                {                                  //чтение файла
                    string[] allText = File.ReadAllLines(fileName);         //чтение всех строк файла в массив строк
                    if (allText.Length > 0)
                    {
                        mainGridView.ServName = allText[0];
                        mainGridView.InstName = allText[1];
                        mainGridView.PortNumb = allText[2];
                        mainGridView.DataBase = allText[3];
                        mainGridView.UserID = allText[4];
                        mainGridView.Password = allText[5];
                        if (allText.Length > 6)
                            mainGridView.NumbStreamReader = allText[6];
                    }
                    MainGrid.DataContext = mainGridView;
                }
                catch (FileNotFoundException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void SetParams()
        {
            using (StreamWriter sw = new StreamWriter(new FileStream(fileName, FileMode.Create, FileAccess.Write)))
            {
                sw.WriteLine(mainGridView.ServName);
                sw.WriteLine(mainGridView.InstName);
                sw.WriteLine(mainGridView.PortNumb);
                sw.WriteLine(mainGridView.DataBase);
                sw.WriteLine(mainGridView.UserID);
                sw.WriteLine(mainGridView.Password);
                sw.WriteLine(mainGridView.NumbStreamReader);
            }
        }

        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ButtonOK_Click(object sender, RoutedEventArgs e)
        {
            SetParams();
            this.Close();
        }
    }
}
