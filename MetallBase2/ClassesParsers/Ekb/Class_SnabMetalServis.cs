﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;

namespace MetallBase2.ClassesParsers.Ekb
{
	class Class_SnabMetalServis
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
			ReadWord();
			//return dtProduct;
		}

		Word._Application application;
		Word._Document document;
		Object missingObj = System.Reflection.Missing.Value;
		Object trueObj = true;
		Object falseObj = false;
		bool isOpenWord = false;

		DataTable dtProduct = new DataTable();

		string orgname = "";


		public string NameOrg() { return orgname; }

		private void ReadWord()
		{
			InfoOrganization infoOrg = new InfoOrganization
			{
				SkladAdr = new List<string>(),
				Manager = new List<string[]>()
			};

			//создаем обьект приложения word
			application = new Word.Application();
			// создаем путь к файлу
			Object templatePathObj = filePath;
			Word.Tables tables;
			try
			{
				orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
				infoOrg.OrgName = orgname;

				document = application.Documents.Open(ref templatePathObj, ref missingObj,
					ref missingObj, ref missingObj, ref missingObj, ref missingObj,
					ref missingObj, ref missingObj, ref missingObj, ref missingObj,
					ref missingObj, ref missingObj, ref missingObj, ref missingObj,
					ref missingObj, ref missingObj);

				tables = document.Tables;
				isOpenWord = true;

				int lastRow = 0;
				int ColDiam = 0, ColTolsh = 0, ColMera = 0, ColMark = 0, ColPrim = 0, ColTU = 0, ColName = 0, ColPrice = 0;
				C_InfoTable tab;

				string temp = "", tmp = "", price = "", prim = "", name = "", type = "", mark = "";
				string diam = "", tolsh = "", metraj = "", mera = "", standart = "";
				var regexParam = new C_RegexParamProduct();
				List<double> ddiam = new List<double>();
				List<double> dtolsh = new List<double>();
				List<double> dmetraj = new List<double>();

				int progress = 0;

				for (int t = 1; t <= tables.Count; t++)
				{
					tab = new C_InfoTable();
					Word.Table wTab = tables[t];
					int cCelCol = wTab.Columns.Count;
					int cCelRow = wTab.Rows.Count;
					//if (cCelCol < 10) cCelCol = 10;
					//if (cCelCol > 20) cCelCol = 20;

					int Max = cCelCol * cCelRow;
					SetMaxValProgressBar(Max);
					//Поиск заголовков столбцов
					progress = 0;
					for (int j = 1; j <= cCelRow; j++) //строки
					{
						int jj = j;
						for (int i = 1; i <= cCelCol; i++) //столбцы
						{
							if (i <= wTab.Rows[jj].Cells.Count)
							{
								temp = wTab.Cell(jj, i).Range.Text;
								temp = temp.Replace("\r\a", string.Empty).Trim();
								if (new Regex(@"^наименование", RegexOptions.IgnoreCase).IsMatch(temp))
								{ ColName = i; j = cCelRow; tab.StartRow = jj; continue; }
								else if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
								{ ColMera = i; continue; }
								else if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
								{ ColPrice = i; continue; }

								if (progress < Max) ProcessChanged(progress++);
								else ProcessChanged(Max);
							}
						}
					}
					Max = cCelRow;
					SetMaxValProgressBar(Max);
					progress = 0;

					for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) //строки
					{
						if (ColName != 0)
						{
							temp = wTab.Cell(jj, ColName).Range.Text;
							temp = temp.Replace("\r\a", string.Empty).Trim();
							if (temp != "")
							{
								diam = ""; tolsh = ""; metraj = ""; tab.Standart = ""; price = ""; name = "";
								ddiam = new List<double>();
								dtolsh = new List<double>();
								dmetraj = new List<double>();

								if (regexParam.RegName.IsMatch(temp))
								{
									prim = temp;
									name = regexParam.RegName.Match(temp).Value;
									temp = temp.Replace(name, "");
									type = regexParam.RegType.Match(temp).Value;
									if (!string.IsNullOrEmpty(type))
										temp = temp.Replace(type, "");
									standart = regexParam.RegTU.Match(temp).Value;
									if (!string.IsNullOrEmpty(standart))
										temp = temp.Replace(standart, "");
									//mark = new Regex(@"(?<=(?:сталь|ст)\s*\.?\s*)\d+", RegexOptions.IgnoreCase).Match(temp).Value;
									//if (string.IsNullOrEmpty(mark))
									{
										mark = regexParam.RegMark.Match(temp).Value;
									}
									if (!string.IsNullOrEmpty(mark))
										temp = temp.Replace(mark, "");

									if (new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*\d+(?:[,.]\d+)?(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).IsMatch(temp))
									{
										string tstr = new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?(?=[xх]\s*\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*\d+(?:[,.]\d+)?(-\s*\d+(?:[,.]\d+)?\s*)?)", RegexOptions.IgnoreCase).Match(temp).Value;
										ddiam = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
										tstr = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*)\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?(?=[xх]\s*\d+(?:[,.]\d+)?(-\s*\d+(?:[,.]\d+)?\s*)?)", RegexOptions.IgnoreCase).Match(temp).Value;
										dtolsh = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
										tstr = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*)\d+(?:[,.]\d+)?(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).Match(temp).Value;
										dmetraj = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
									}
									if (new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).IsMatch(temp))
									{
										string tstr = new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?(?=[xх]\s*\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?)", RegexOptions.IgnoreCase).Match(temp).Value;
										ddiam = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
										tstr = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?[xх]\s*)\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).Match(temp).Value;
										dtolsh = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
									}
									if (new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).IsMatch(temp))
									{
										string tstr = new Regex(@"\d+(?:[,.]\d+)?\s*(-\s*\d+(?:[,.]\d+)?\s*)?", RegexOptions.IgnoreCase).Match(temp).Value;
										ddiam = GetIncrementingMassiv(tstr.Replace(" ", "").Trim().Split('-'));
									}
									if (ddiam.Count == 0)
									{
										diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
									}
									if (dmetraj.Count == 0)
									{
										metraj = new Regex(@"\d+(?:[,.]\d+)?(?=м\b)", RegexOptions.IgnoreCase).Match(temp).Value;
										if (!string.IsNullOrEmpty(metraj))
											dmetraj = GetIncrementingMassiv(metraj.Replace(" ", "").Trim().Split('-'));
									}
									if (dtolsh.Count == 0) dtolsh.Add(0);
									if (dmetraj.Count == 0) dmetraj.Add(0);

									if (ddiam.Count > 0)
										if (ddiam[0] != 0)
										{
											if (ColMera > 0)
											{
												tmp = "";
												tmp = wTab.Cell(jj, ColMera).Range.Text;
												tmp = tmp.Replace("\r\a", string.Empty).Trim();
												if (tmp != "")
												{
													mera = tmp;
												}
											}
											if (ColPrice > 0)
											{
												tmp = "";
												tmp = wTab.Cell(jj, ColPrice).Range.Text;
												tmp = tmp.Replace("\r\a", string.Empty).Trim();
												if (tmp != "")
												{
													price = tmp;
												}
											}

											for (int di = 0; di < ddiam.Count; di++)
												for (int ti = 0; ti < dtolsh.Count; ti++)
													for (int mi = 0; mi < dmetraj.Count; mi++)
													{
														DataRow row = dtProduct.NewRow();
														if (!string.IsNullOrEmpty(name))
															row["Название"] = name;
														else row["Название"] = tab.Name;
														if (string.IsNullOrEmpty(type))
															row["Тип"] = tab.Type;
														else row["Тип"] = type;
														if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
														else row["Тип"] = row["Тип"].ToString().ToLower();
														if (ddiam[di] == 0) row["Диаметр (высота), мм"] = "";
														else
															row["Диаметр (высота), мм"] = ddiam[di];
														if (dtolsh[ti] == 0) row["Толщина (ширина), мм"] = "";
														else
															row["Толщина (ширина), мм"] = dtolsh[ti];
														if (dmetraj[mi] == 0) row["Метраж, м (длина, мм)"] = "";
														else
															row["Метраж, м (длина, мм)"] = dmetraj[mi];
														row["Мерность (т, м, мм)"] = mera;
														if (String.IsNullOrEmpty(mark))
															row["Марка"] = tab.Mark;
														else row["Марка"] = mark;
														if (String.IsNullOrEmpty(tab.Standart))
															row["Стандарт"] = standart;
														else row["Стандарт"] = tab.Standart;
														row["Класс"] = "";
														row["Цена"] = price;
														row["Примечание"] = prim;
														dtProduct.Rows.Add(row);
													}
										}
								}
							}
							if (progress < Max) ProcessChanged(progress++);
							else ProcessChanged(Max);
						}
					}
				}

				SetMaxValProgressBar(document.Sections.Count);
				ProcessChanged(0);
				progress = 0;
				// поиск информации об организации в первых 10 параграфах
				for (int i = 1; i <= document.Shapes.Count; i++)
				{
					if (document.Shapes[i].TextFrame.HasText == 0) continue;
					temp = document.Shapes[i].TextFrame.TextRange.Text;
					temp = temp.Replace("\r\a", string.Empty).Trim();
					if (temp != "")
					{
						if (new Regex(@"\d{6}\s*г.\w+,?\s*ул.\s*\w+\s*\d+\s*\w+\s*\d+(?:/\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.OrgAdress = new Regex(@"\d{6}\s*г.\w+,?\s*ул.\s*\w+\s*\d+\s*\w+\s*\d+(?:/\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						//if (regexParam.OrgMobileTelefon.IsMatch(temp))
						//{
						//    infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
						//}
						if (new Regex(@"[+\d]?(?:\s*\(\d+\)\s*)?(?:\d+-\d+-\d+,?\s*)+(?:\s*\((?:\d+,?\s*)+\))?", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.OrgTel = new Regex(@"[+\d]?(?:\s*\(\d+\)\s*)?(?:\d+-\d+-\d+,?\s*)+(?:\s*\((?:\d+,?\s*)+\))?", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						if (regexParam.EMail.IsMatch(temp))
						{
							infoOrg.Email = regexParam.EMail.Match(temp).Value;
						}
						if (regexParam.Site.IsMatch(temp))
						{
							infoOrg.Site = regexParam.Site.Match(temp).Value;
						}
						else if (new Regex(@"(?<=\s|^)(?:www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.Site = new Regex(@"(?<=\s|^)(?:www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						if (new Regex(@"адрес\s*базы", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.SkladAdr.Add(new Regex(@"(?<=дрес\s*базы\s*:?\s*)[\w\s\.\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
						}
						if (new Regex(@"(?<=ИНН(?:/КПП)?\s*:?\s*)\d{9,15}\s*/\s*\d{9,15}", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.Inn_Kpp = new Regex(@"(?<=ИНН(?:/КПП)?\s*:?\s*)\d{9,15}\s*/\s*\d{9,15}", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						if (new Regex(@"(?<=Р.сч\s*)\d+", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.r_s = new Regex(@"(?<=Р.сч\s*)\d+", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						if (new Regex(@"(?<=К.сч\s*)\d+", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.k_s = new Regex(@"(?<=К.сч\s*)\d+", RegexOptions.IgnoreCase).Match(temp).Value;
						}
						if (new Regex(@"(?<=\bбик\b\s*)\d+", RegexOptions.IgnoreCase).IsMatch(temp))
						{
							infoOrg.BIK = new Regex(@"(?<=\bбик\b\s*)\d+", RegexOptions.IgnoreCase).Match(temp).Value;
						}
					}
					ProcessChanged(progress++);
				}

				SetInfoOrganization(infoOrg);
				WorkCompleted(dtProduct);
			}
			catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(filePath) + "\n\n" + ex.ToString()); }
			if (isOpenWord)
			{
				document.Close();
				application.Quit(missingObj, missingObj, missingObj);
			}
		}

		private static string StringFirstUp(string StringIn)
		{
			string StringOut = "";
			if (StringIn.Length > 2)
				StringOut = StringIn.Substring(0, 1).ToUpper() + StringIn.Substring(1, StringIn.Length - 1).ToLower();
			else StringOut = StringIn;

			return StringOut;
		}

		private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam)
		{
			if (new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).IsMatch(temp))
			{
				infoOrg.OrgAdress = new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).Match(temp).Value;
			}
			if (regexParam.OrgMobileTelefon.IsMatch(temp))
			{
				infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
			}
			if (new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).IsMatch(temp))
			{
				infoOrg.OrgTel = new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).Match(temp).Value;
			}
			if (regexParam.EMail.IsMatch(temp))
			{
				infoOrg.Email = regexParam.EMail.Match(temp).Value;
			}
			if (regexParam.Site.IsMatch(temp))
			{
				infoOrg.Site = regexParam.Site.Match(temp).Value;
			}
			if (new Regex(@"адрес\s*базы", RegexOptions.IgnoreCase).IsMatch(temp))
			{
				infoOrg.SkladAdr.Add(new Regex(@"(?<=дрес\s*базы\s*:?\s*)[\w\s\.\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
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
				Ddiam.Add(Convert.ToDouble(str));
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


		public event Action<int> ProcessChanged;

		public event Action<int> SetMaxValProgressBar;

		public event Action<InfoOrganization> SetInfoOrganization;

		public event Action<DataTable> WorkCompleted;
	}
}
