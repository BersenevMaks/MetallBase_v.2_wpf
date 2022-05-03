﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetallBase2
{
    class Class_Skat
    {
        private string filePath;

        public void Set(string Path)
        {
            filePath = Path;
        }

        public void GetTableFromExcel()
        {
            dtProduct.Columns.Add("Название");
            dtProduct.Columns.Add("Тип");
            dtProduct.Columns.Add("Диаметр (высота), мм");
            dtProduct.Columns.Add("Толщина (ширина), мм");
            dtProduct.Columns.Add("Метраж, м (длина, мм)");
            dtProduct.Columns.Add("Мерность (т, м, мм)");
            dtProduct.Columns.Add("Марка");
            dtProduct.Columns.Add("Стандарт");
            dtProduct.Columns.Add("Класс");
            dtProduct.Columns.Add("Цена");
            dtProduct.Columns.Add("Примечание");
            ReadExcel();
            //return dtProduct;
        }

        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        bool isExcelOpen = false;
        DataTable dtProduct = new DataTable();

        string orgname = "";

        public string NameOrg() { return orgname; }

        private void ReadExcel()
        {
            InfoOrganization infoOrg = new InfoOrganization
            {
                SkladAdr = new List<string>(),
                Manager = new List<string[]>()
            };

            excelapp = new Excel.Application();
            try
            {
                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                infoOrg.OrgName = orgname;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                    isExcelOpen = true;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла " + orgname + "\n\n" + ex.ToString()); isExcelOpen = false; }

                string temp = "", tmp = "", standart = "", mark = "", name = "", type = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "";
                var regexParam = new C_RegexParamProduct();
                List<double> dDiam;
                List<double> dTolsh;
                List<double> dMetraj;

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    //if (excelworksheet.Index != 9) continue;
                    //MessageBox.Show(excelsheets.Count.ToString());
                    var tab = new C_InfoTable();
                    var naaame = excelworksheet.Name;
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 25;
                    //cCelCol = 10;

                    int ColDlina = 0, ColMera = 0, ColPrice = 0, ColGost = 0, ColSoder = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;

                    for (int j = 4; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^содержание", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab = new C_InfoTable
                                    {
                                        StartCol = i,
                                        StartRow = jj
                                    };
                                    tabs.Add(tab);
                                    j = cCelRow + 1;
                                    break;
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }
                    }

                    ProcessChanged(0);
                    Max = tabs.Count;
                    SetMaxValProgressBar(Max);
                    progress = 0;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        ColPrice = 0; ColDlina = 0; ColMera = 0; ColSoder = 0; ColGost = 0;
                        name = ""; type = "";
                        Excel.Range cellRange;
                        tab = tabs[k];
                        int endRow = cCelRow;
                        if (k < tabs.Count - 1)   // определение последней строки в текущей минитаблице
                            if (tab.StartCol == tabs[k + 1].StartCol)
                                endRow = tabs[k + 1].StartRow - 1;
                            else if (k < tabs.Count - 2)
                                if (tab.StartCol == tabs[k + 2].StartCol)
                                    endRow = tabs[k + 2].StartRow - 1;
                                else if (k < tabs.Count - 3)
                                    if (tab.StartCol == tabs[k + 3].StartCol)
                                        endRow = tabs[k + 3].StartRow - 1;

                        // // // поиск имени продукции
                        ProcessChanged(0); //установить прогрессбар на 0
                        SetMaxValProgressBar(cCelCol); Max = cCelCol; //установить максимум для прогрессбара
                        progress = 0;
                        int jjj = tab.StartRow;
                        for (int i = tab.StartCol; i <= cCelCol; i++) //столбцы
                        {
                            cellRange = (Excel.Range)excelworksheet.Cells[jjj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"\bсодержание\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColSoder = i;
                                }
                                if (new Regex(@"\bгост\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColGost = i;
                                }
                                if (new Regex(@"\bкол[\w\d\-\.,]во\b(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMera = i;
                                }
                                if (new Regex(@"\bпараметры\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDlina = i;
                                }
                                if (new Regex(@"\bцена\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }

                        ProcessChanged(0);//установить прогрессбар на 0
                        progress = 0;
                        SetMaxValProgressBar(endRow * (k + 1)); //установить максимум для прогрессбара
                        Max = endRow * (k + 1);
                        for (int j = tab.StartRow + 1; j <= endRow; j++)
                        {
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; price = ""; standart = ""; mark = ""; name = "";
                            if (ColSoder > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColSoder];
                                if (cellRange.MergeArea.Columns.Count > 2) continue;
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    name = regexParam.RegName.Match(temp).Value;
                                    type = new Regex(@"оцин|эл.св|рыж|г/к|х/к|х/д", RegexOptions.IgnoreCase).Match(temp).Value;
                                    switch (type)
                                    {
                                        case "оцин":
                                            type = "оцинкованная";
                                            break;
                                        case "эл.св":
                                            type = "электросварная";
                                            break;
                                        case "рыж":
                                            type = "рыжая";
                                            break;
                                        case "г/к":
                                            type = "горячекатаная";
                                            break;
                                        case "х/к":
                                            type = "холоднокатаная";
                                            break;
                                        case "х/д":
                                            type = "холоднодеформированная";
                                            break;
                                        default:
                                            type = "";
                                            break;
                                    }

                                    if (string.IsNullOrEmpty(type))
                                    {
                                        type = regexParam.RegType.Match(temp).Value;
                                    }

                                    if (ColDlina > 0)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColDlina];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            metraj = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(name))
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColSoder];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            prim = temp;
                                            temp = temp.Replace(".", ",");
                                            temp = temp.Trim();
                                            //diamDxDxD
                                            if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                            }
                                            //diamDxD
                                            else if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                try
                                                {
                                                    if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show("Ошибка преобразования\n diam = " + diam + ", tolsh = " + tolsh + "\n" + ex.ToString());
                                                }

                                            }
                                            //diam D
                                            else if (new Regex(@"\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                            }

                                            standart = regexParam.RegTU.Match(temp).Value;
                                            mark = regexParam.RegMark.Match(temp).Value;
                                            if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(mark)) mark = "";
                                        }
                                    }
                                }

                                if (ColMera > 0)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[j, ColMera];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        mera = temp;
                                    }
                                }

                                
                                if (ColGost > 0)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[j, ColGost];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        standart = temp;
                                    }
                                }

                                if (ColPrice > 0)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        price = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }

                                if (!String.IsNullOrEmpty(diam))
                                {
                                    DataRow row = dtProduct.NewRow();
                                    row["Название"] = name;
                                    if (string.IsNullOrEmpty(type))
                                        row["Тип"] = tab.Type;
                                    else row["Тип"] = type;
                                    if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                    else row["Тип"] = row["Тип"].ToString().ToLower();
                                    row["Диаметр (высота), мм"] = diam;
                                    row["Толщина (ширина), мм"] = tolsh;
                                    row["Метраж, м (длина, мм)"] = metraj;
                                    row["Мерность (т, м, мм)"] = mera;
                                    if (String.IsNullOrEmpty(mark))
                                        row["Марка"] = tab.Mark;
                                    else row["Марка"] = mark;
                                    if (String.IsNullOrEmpty(standart))
                                        row["Стандарт"] = tab.Standart;
                                    else row["Стандарт"] = standart;
                                    row["Класс"] = "";
                                    row["Цена"] = price;
                                    row["Примечание"] = prim;
                                    dtProduct.Rows.Add(row);
                                }
                            }
                            if (j * (k + 1) < Max) ProcessChanged(j * (k + 1));
                            else ProcessChanged(Max);
                        }
                    }

                    //поиск информации об организации ТОЛЬКО на ПЕРВОМ листе
                    if (tabs.Count > 0 && dtProduct.Rows.Count > 0 && excelworksheet.Index == 1)
                    {
                        Max = (tabs[0].StartRow - 1) * cCelCol;
                        SetMaxValProgressBar(Max);
                        Excel.Range cellRange;
                        for (int j = 1; j <= tabs[0].StartRow - 1; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    FillInfoOrg(infoOrg, temp, regexParam, excelworksheet, j, i);
                                }

                                if (i * j < Max) ProcessChanged(i * j);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                }

                if (isExcelOpen)
                {
                    //excelappworkbook.Close();
                    excelapp.Quit();
                }

                SetInfoOrganization(infoOrg);
                WorkCompleted(dtProduct);

            }
            catch (Exception ex) { WorkCompleted(dtProduct); MessageBox.Show("Ошибка в функции ReedExcel() в " + this.ToString() + "\n\n" + ex.ToString()); }
        }

        private static string StringFirstUp(string StringIn)
        {
            string StringOut = "";
            if (!String.IsNullOrEmpty(StringIn))
            {
                if (StringIn.Length > 2)
                    StringOut = StringIn.Substring(0, 1).ToUpper() + StringIn.Substring(1, StringIn.Length - 1).ToLower();
                else StringOut = StringIn;
            }
            return StringOut;
        }

        private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam, Excel.Worksheet worksheet, int j, int i)
        {
            //if (regexParam.OrgAdresFull.IsMatch(temp))
            //{
            //    infoOrg.OrgAdress = regexParam.OrgAdresFull.Match(temp).Value;//new Regex(@"(?<=офис.*:\s*).*(?:\w|\)|\.)", RegexOptions.IgnoreCase).Match(temp).Value;
            //}
            //if (string.IsNullOrEmpty(infoOrg.OrgAdress))
            //    if (new Regex(@"г\.\s*(?:\w+\s?){1,2}", RegexOptions.IgnoreCase).IsMatch(temp))
            //        infoOrg.OrgAdress = new Regex(@"г\.\s*(?:\w+\s?){1,2}", RegexOptions.IgnoreCase).Match(temp).Value;
            if (new Regex(@"(?<=тел./факс\s*)[\(\)\d\-,\s]+[\d,](?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=тел./факс\s*)[\(\)\d\-,\s]+[\d,](?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            //if (regexParam.INN_KPP.IsMatch(temp))
            //{
            //    infoOrg.Inn_Kpp = regexParam.INN_KPP.Match(temp).Value;
            //}
            //if (regexParam.R_S.IsMatch(temp))
            //{
            //    infoOrg.r_s = regexParam.R_S.Match(temp).Value;
            //}
            //if (regexParam.K_S.IsMatch(temp))
            //{
            //    infoOrg.k_s = regexParam.K_S.Match(temp).Value;
            //}
            //if (regexParam.BIK.IsMatch(temp))
            //{
            //    infoOrg.BIK = regexParam.BIK.Match(temp).Value;
            //}

            //if (new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    foreach (Match m in new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).Matches(temp))
            //        if (String.IsNullOrEmpty(infoOrg.Email)) infoOrg.Email = m.Value;
            //        else infoOrg.Email += "; " + m.Value;
            //}
            if (new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Site = new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.SkladAdr.Add(new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
            }
            if (new Regex(@"(?:\w+\s*)+тел.*\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                string tmp = new Regex(@"(?:\w+\s*)+тел.*\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).Match(temp).Value;
                string[] manager = new string[]
                {
                    new Regex(@"(?:\w+\s*)+(?=тел.*\.(?:ru|com|info|рф))", RegexOptions.IgnoreCase).Match(temp).Value, //имя
                    new Regex(@"(?<=(?:\w+\s*)+тел.*)[\d\-]+(?=\s*e-.*(?:ru|com|info|рф))", RegexOptions.IgnoreCase).Match(temp).Value, //телефон
                    new Regex(@"(?<=(?:\w+\s*)+тел.*[\d\-]+\s*e-mail:).*(?:ru|com|info|рф)", RegexOptions.IgnoreCase).Match(temp).Value //email
                };
                infoOrg.Manager.Add(manager);
            }
        }

        private List<double> GetIncrementingMassiv(string[] strParams)
        {
            List<double> Ddiam = new List<double>();
            List<double> ch = new List<double>();
            string str;
            foreach (string s in strParams)
            {
                str = s.Replace('.', ',');
                Ddiam.Add(Convert.ToDouble(s));
            }
            if (strParams.Length > 1)
            {
                double increment = 0;
                if (Ddiam[1] >= 1 && Ddiam[1] < 4) increment = 0.5;
                if (Ddiam[1] >= 4 && Ddiam[1] < 50) increment = 2;
                if (Ddiam[1] >= 50 && Ddiam[1] < 100) increment = 10;
                if (Ddiam[1] >= 100 && Ddiam[1] < 1000) increment = 100;
                if (Ddiam[1] >= 1000 && Ddiam[1] < 5000) increment = 500;
                if (Ddiam[1] >= 5000 && Ddiam[1] < 50000) increment = 1000;

                if (increment > 0)
                {
                    for (double d = Ddiam[0]; d <= Ddiam[1]; d += increment)
                    {
                        if (d != Ddiam[0] && d % 1 == 1)
                            d -= 0.1;
                        ch.Add(d);
                        if (d + increment > Ddiam[1] && d != Ddiam[1]) ch.Add(Ddiam[1]);
                    }
                    if (ch.Count > 0) Ddiam = ch;
                }
            }
            return Ddiam;
        }

        public event Action<int> ProcessChanged; //установить текущее значение прогрессбара

        public event Action<int> SetMaxValProgressBar; //установить максимальное значение для прогрессбара

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
