using MetallBase2.Classes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MetallBase2
{
    /// <summary>
    /// Логика взаимодействия для RedactorMarks.xaml
    /// </summary>
    public partial class RedactorMarks : Window
    {
        public RedactorMarks(string connString, string TypeMark)
        {
            InitializeComponent();
            sqlConnectionString = connString;
            typeMark = TypeMark;
            redactorMarksViewModel = new RedactorMarksViewModel();
            switch (typeMark)
            {
                case "Nerj":
                    this.Title += "Нержавейка";
                    break;
                case "Alumin":
                    this.Title += "Алюминий";
                    break;
                case "Med":
                    this.Title += "Медь";
                    break;
                case "Latun":
                    this.Title += "Латунь";
                    break;
            }
            Refresh();
        }

        string sqlConnectionString = "";
        string typeMark = "";
        public RedactorMarksViewModel redactorMarksViewModel;

        private void Refresh()
        {
            DataTable dtMarks = new DataTable();

            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"select Mark from Marks where MarkType = '" + typeMark + "'";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtMarks);
                redactorMarksViewModel.Marks = new List<string>();
                foreach (DataRow row in dtMarks.Rows)
                {
                    redactorMarksViewModel.Marks.Add(row["Mark"].ToString());
                }

                if (conn.State == ConnectionState.Open) conn.Close();
                SetDataContextes();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                SetDataContextes();
            }
        }

        private void SetDataContextes()
        {
            MainGrid.DataContext = null;
            MainGrid.DataContext = redactorMarksViewModel;
        }

        private void ListView_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListView lv)
            {
                redactorMarksViewModel.NewMark = lv.SelectedItem.ToString();
                SetDataContextes();
            }
        }

        private void BtnAddMark_Click(object sender, RoutedEventArgs e)
        {

            SqlConnection conn = new SqlConnection(sqlConnectionString);

            if (!string.IsNullOrEmpty(redactorMarksViewModel.NewMark))
            {
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    var sqlCmd = new SqlCommand("dbo.insMarks", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ typeMark);
                    sqlCmd.Parameters.AddWithValue("@Mark", /* Значение параметра */ redactorMarksViewModel.NewMark);
                    int k = sqlCmd.ExecuteNonQuery();
                    MessageBox.Show("Марка " + redactorMarksViewModel.NewMark + " успешно добавлена");
                    redactorMarksViewModel.NewMark = "";
                }
                catch
                {
                    {
                        MessageBox.Show("Ошибка, марка не добавлена");
                        redactorMarksViewModel.NewMark = "";
                    }
                }
                Refresh();
            }
        }

        private void BtnDelCurMark_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                var sqlCommand = new SqlCommand("delete from dbo.Marks where MarkType='" + typeMark + "' and mark='" + redactorMarksViewModel.SelectedMark + "'", conn);
                int result = sqlCommand.ExecuteNonQuery();
                if (result > -1)
                    MessageBox.Show("Выбранная марка удалена");
                else MessageBox.Show("Ошибка при удалении/nВыбранная марка не удалена");
                Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ListView_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListView lv)
                if (lv.SelectedItem is string s)
                    redactorMarksViewModel.SelectedMark = s;
        }

        private void BtnAddMarkFromFile_Click(object sender, RoutedEventArgs e)
        {
            List<string> result = new List<string>();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            SqlCommand comm;
            string[] str;
            string query;
            int count = 0;
            string message = "Будут добавлены следующие марки:\n";
            string PathFile = "";
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                if (ofd.ShowDialog() == true)
                    PathFile = ofd.FileName;
                if (!string.IsNullOrEmpty(PathFile))
                {
                    using (StreamReader sr = new StreamReader(PathFile, System.Text.Encoding.Default))
                    {
                        string st = sr.ReadToEnd();
                        str = st.Split(';');

                        foreach (string s in str)
                        {
                            result.Add(s.Replace("\r", "").Replace("\n", ""));
                            message += result[result.Count - 1] + "\n";
                        }
                    }
                }
                message += "\nДобавить их в базу?";
                if (MessageBox.Show(message, "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    foreach (string AddingMark in result)
                    {
                        if (conn.State == ConnectionState.Closed) conn.Open();
                        var sqlCmd = new SqlCommand("dbo.insMarks", conn);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ typeMark);
                        sqlCmd.Parameters.AddWithValue("@Mark", /* Значение параметра */ AddingMark);
                        int k = sqlCmd.ExecuteNonQuery();
                        count = k;
                    }
                    Refresh();
                    if (count == result.Count)
                    {
                        MessageBox.Show("Успешно добавлено");
                    }
                    //else MessageBox.Show("Не все марки были добавлены, проверьте список.");
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при чтении файла\n\n" + ex.ToString()); }
        }
    }
}
