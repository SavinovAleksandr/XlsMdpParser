using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Style;
using Xls_prjt;

namespace XlsxMdpParser;

internal class Program
{
	private static void Main(string[] args)
	{
		Console.WriteLine("Перетащите один или несколько файлов в это окно и нажмите Enter:");
		List<string> inputPaths = GetInputPaths(args);
		string summaryB1Text = LoadSummaryB1Config();
		if (inputPaths.Count > 0)
		{
			foreach (string text2 in inputPaths)
			{
				try
				{
					Console.WriteLine("Получен путь: " + text2);
					ExcelOperations excelOperations = new ExcelOperations(text2, 1);
					ColumnMap columnMap = ResolveColumnMap(excelOperations);
					List<MdpBuilder> list = new List<MdpBuilder>();
					for (int i = 4; i <= excelOperations.LastColumnRow(); i++)
					{
						if (!(excelOperations.getStr(i, columnMap.SchemeNameCol) != "") || !(excelOperations.getStr(i, columnMap.SchemeNameCol) != " "))
						{
							continue;
						}
						string str = excelOperations.getStr(i, columnMap.SchemeNameCol);
						string str2 = excelOperations.getStr(i, columnMap.SchemeNumCol);
						string text3 = excelOperations.MergedCells(i, columnMap.SchemeNameCol);
						int num10 = Convert.ToInt32(text3.Split(new char[1] { ':' })[0].Substring(1));
						int num11 = Convert.ToInt32(text3.Split(new char[1] { ':' })[1].Substring(1));
						List<TNV> list2 = new List<TNV>();
						for (int j = num10; j <= num11; )
						{
							while (j <= num11 && string.IsNullOrWhiteSpace(excelOperations.getStr(j, columnMap.TnvCol)))
							{
								j++;
							}
							if (j > num11)
							{
								break;
							}
							int bRow = j;
							int eRow = j;
							while (eRow < num11 && string.IsNullOrWhiteSpace(excelOperations.getStr(eRow + 1, 4)))
							{
								eRow++;
							}
							list2.Add(new TNV
							{
								Tnv = ReadLine(excelOperations, bRow, eRow, columnMap.TnvCol),
								MdpNoPA = ReadLines(excelOperations, bRow, eRow, columnMap.MdpNoPaCol, modify: true),
								MdpPa = ((columnMap.MdpPaCol != -1) ? ReadLines(excelOperations, bRow, eRow, columnMap.MdpPaCol, modify: true) : new List<MDP>()),
								Adp = ReadLine(excelOperations, bRow, eRow, columnMap.AdpCol),
								MdpNoPaCriteria = ReadLines(excelOperations, bRow, eRow, columnMap.MdpNoPaCriteriaCol),
								MdpPaCriteria = ((columnMap.MdpPaCriteriaCol != -1) ? ReadLines(excelOperations, bRow, eRow, columnMap.MdpPaCriteriaCol) : new List<MDP>()),
								AdpCriteria = ReadLine(excelOperations, bRow, eRow, columnMap.AdpCriteriaCol),
								MdpNoPaDop = ReadDopLines(excelOperations, bRow, eRow, columnMap.MdpNoPaDopCol),
								MdpPaDop = ((columnMap.MdpPaDopCol != -1) ? ReadDopLines(excelOperations, bRow, eRow, columnMap.MdpPaDopCol) : new List<string>()),
								AdpDop = ReadDopLines(excelOperations, bRow, eRow, columnMap.AdpDopCol)
							});
							j = eRow + 1;
						}
						list.Add(new MdpBuilder
						{
							ShemeName = str,
							ShemeNum = str2,
							TnvList = list2
						});
					}
					excelOperations.AddList("new");
					int num12 = 10;
					int[] array = new int[12]
					{
						7, 40, 11, 80, 120, 30, 50, 50, 30, 25,
						25, 25
					};
					for (int k = 1; k <= array.Count(); k++)
					{
						excelOperations.Width(k, array[k - 1]);
					}
					excelOperations.setVal(1, 1, "№ п/п");
					excelOperations.Merge(1, 1, 2, 1, hor: true, vert: true);
					excelOperations.setVal(1, 2, "Схема сети");
					excelOperations.Merge(1, 2, 2, 2, hor: true, vert: true);
					excelOperations.setVal(1, 3, "ТНВ, °С");
					excelOperations.Merge(1, 3, 2, 3, hor: true, vert: true);
					excelOperations.setVal(1, 4, "МДП без ПА");
					excelOperations.Merge(1, 4, 2, 4, hor: true, vert: true);
					excelOperations.setVal(1, 5, "МДП с ПА");
					excelOperations.Merge(1, 5, 2, 5, hor: true, vert: true);
					excelOperations.setVal(1, 6, "АДП");
					excelOperations.Merge(1, 6, 2, 6, hor: true, vert: true);
					excelOperations.setVal(1, 7, "Критерий определения допустимых перетоков");
					excelOperations.Merge(1, 7, 1, 9, hor: true, vert: true);
					excelOperations.setVal(2, 7, "МДП без ПА");
					excelOperations.Format(2, 7, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(2, 8, "МДП с ПА");
					excelOperations.Format(2, 8, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(2, 9, "АДП");
					excelOperations.Format(2, 9, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(1, 10, "Контроль дополнительных параметров");
					excelOperations.Merge(1, 10, 1, 12, hor: true, vert: true);
					excelOperations.setVal(2, 10, "МДП без ПА");
					excelOperations.Format(2, 10, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(2, 11, "МДП с ПА");
					excelOperations.Format(2, 11, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(2, 12, "АДП");
					excelOperations.Format(2, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.FreezeRows(2);
					if (!columnMap.HasMdpPa)
					{
						excelOperations.HideColumn(5);
						excelOperations.HideColumn(8);
						excelOperations.HideColumn(11);
					}
					excelOperations.FormatCells(1, 1, 2, array.Count(), bold: true, italic: false, Color.PowderBlue.ToArgb());
					int num4 = 3;
					Dictionary<string, int> dictionary = new Dictionary<string, int>();
					List<int> notControlledRows = new List<int>();
					foreach (MdpBuilder item in list)
					{
						string key = item.ShemeNum.Trim(new char[1] { ' ' });
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, num4);
						}
						excelOperations.setVal(num4, 1, item.ShemeNum);
						excelOperations.Format(num4, 1, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
						string text3 = GetSchemeHeaderLine(item.ShemeName);
						excelOperations.setVal(num4, 2, text3, wrap: false);
						excelOperations.Merge(num4, 2, num4, array.Count());
						excelOperations.Format(num4, 2, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Center);
						excelOperations.Wrap(num4, 2, wrap: false);
						excelOperations.FormatCells(num4, 1, num4, array.Count(), bold: false, italic: false, Color.MistyRose.ToArgb());
						int textWidth = array.Skip(1).Sum();
						int rowHeight = EstimateMergedRowHeight(text3, textWidth, num12);
						excelOperations.Height(num4, Math.Max(20, rowHeight));
						int num5 = num4 + 1;
						int num6 = num5;
						string mergedAdpDop = GetSingleSchemeAdpDopValue(item.TnvList);
						bool mergeAdpDop = !string.IsNullOrWhiteSpace(mergedAdpDop);
						HashSet<int> hashSet = new HashSet<int>();
						excelOperations.setVal(num5, 1, item.ShemeNum);
						excelOperations.Merge(num5, 1, num5 + item.TnvList.Count - 1, 1);
						excelOperations.Format(num5, 1, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
						excelOperations.setVal(num5, 2, item.ShemeName);
						excelOperations.Merge(num5, 2, num5 + item.TnvList.Count - 1, 2);
						excelOperations.Format(num5, 2, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Center);
						foreach (TNV tnv in item.TnvList)
						{
							if (IsNotControlledPhrase(tnv.Tnv))
							{
								excelOperations.setVal(num5, 3, "Не контролируется", wrap: false);
								excelOperations.Merge(num5, 3, num5, array.Count());
								excelOperations.Format(num5, 3, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
								excelOperations.FontColor(num5, 3, Color.Red);
								excelOperations.FontStyle(num5, 3, 14f, italic: true);
								notControlledRows.Add(num5);
								hashSet.Add(num5);
								num5++;
								continue;
							}
							excelOperations.setVal(num5, 3, tnv.Tnv);
							excelOperations.Format(num5, 3, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							excelOperations.setVal(num5, 4, "");
							excelOperations.Format(num5, 4, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
							excelOperations.setVal(num5, 5, "");
							excelOperations.Format(num5, 5, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
							List<MDP> list3 = tnv.MdpNoPA.Where((MDP mDP) => mDP.Criteria != "").ToList();
							List<MDP> list4 = list3.Where((MDP mDP) => mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							List<MDP> list5 = list3.Where((MDP mDP) => !mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							if (list5.Count <= 1)
							{
								list4.Clear();
							}
							else if (list4.Count == 0)
							{
								list4.Add(new MDP
								{
									Num = -1,
									Criteria = "Минимальное из:"
								});
							}
							List<MDP> list6 = list4.Concat(list5).ToList();
							for (int l = 0; l < list6.Count; l++)
							{
								MDP mDP = list6[l];
								bool flag2 = l == list6.Count - 1;
								excelOperations.CellRichText(num5, 4, (!flag2) ? (mDP.Criteria + Environment.NewLine) : mDP.Criteria, (mDP.Num != -1) ? $"{mDP.Num}) " : "");
							}
							List<MDP> list7 = tnv.MdpPa.Where((MDP mDP) => mDP.Criteria != "").ToList();
							List<MDP> list8 = list7.Where((MDP mDP) => mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							List<MDP> list9 = list7.Where((MDP mDP) => !mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							if (list9.Count <= 1)
							{
								list8.Clear();
							}
							else if (list8.Count == 0)
							{
								list8.Add(new MDP
								{
									Num = -1,
									Criteria = "Минимальное из:"
								});
							}
							List<MDP> list10 = list8.Concat(list9).ToList();
							for (int m = 0; m < list10.Count; m++)
							{
								MDP mDP2 = list10[m];
								bool flag3 = m == list10.Count - 1;
								excelOperations.CellRichText(num5, 5, (!flag3) ? (mDP2.Criteria + Environment.NewLine) : mDP2.Criteria, (mDP2.Num != -1) ? $"{mDP2.Num}) " : "");
							}
							if (tnv.Adp != "")
							{
								excelOperations.setVal(num4 + 1, 6, tnv.Adp);
								excelOperations.Merge(num4 + 1, 6, num4 + item.TnvList.Count, 6);
								excelOperations.Format(num4 + 1, 6, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							}
							string text5 = "";
							foreach (MDP item4 in tnv.MdpNoPaCriteria.Where((MDP mDP) => mDP.Criteria != ""))
							{
								string text6 = ((item4 == tnv.MdpNoPaCriteria.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
								text5 = text5 + ((item4.Num != -1) ? $"{item4.Num}) {item4.Criteria}" : item4.Criteria) + text6;
							}
							excelOperations.setVal(num5, 7, text5);
							excelOperations.Format(num5, 7, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
							excelOperations.CellComment(num5, 4, text5);
							string text7 = "";
							foreach (MDP item5 in tnv.MdpPaCriteria.Where((MDP mDP) => mDP.Criteria != ""))
							{
								string text8 = ((item5 == tnv.MdpPaCriteria.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
								text7 = text7 + ((item5.Num != -1) ? $"{item5.Num}) {item5.Criteria}" : item5.Criteria) + text8;
							}
							excelOperations.setVal(num5, 8, text7);
							excelOperations.Format(num5, 8, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
							excelOperations.CellComment(num5, 5, text7);
							if (tnv.AdpCriteria != "")
							{
								excelOperations.setVal(num4 + 1, 9, tnv.AdpCriteria);
								excelOperations.Merge(num4 + 1, 9, num4 + item.TnvList.Count, 9);
								excelOperations.Format(num4 + 1, 9, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							}
							string text9 = "";
							foreach (string item6 in tnv.MdpNoPaDop)
							{
								string text10 = ((item6 == tnv.MdpNoPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
								text9 = text9 + item6 + text10;
							}
							excelOperations.setVal(num5, 10, text9);
							excelOperations.Format(num5, 10, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							string text11 = "";
							foreach (string item7 in tnv.MdpPaDop)
							{
								string text12 = ((item7 == tnv.MdpPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
								text11 = text11 + item7 + text12;
							}
							excelOperations.setVal(num5, 11, text11);
							excelOperations.Format(num5, 11, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							string text13 = "";
							foreach (string item8 in tnv.AdpDop)
							{
								string text14Line = ((item8 == tnv.AdpDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
								text13 = text13 + item8 + text14Line;
							}
							excelOperations.setVal(num5, 12, text13);
							excelOperations.Format(num5, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							num5++;
						}
						if (mergeAdpDop)
						{
							int num7 = num6;
							while (num7 <= num5 - 1)
							{
								while (num7 <= num5 - 1 && hashSet.Contains(num7))
								{
									num7++;
								}
								if (num7 > num5 - 1)
								{
									break;
								}
								int num8 = num7;
								while (num8 <= num5 - 1 && !hashSet.Contains(num8))
								{
									num8++;
								}
								int num9 = num8 - 1;
								excelOperations.setVal(num7, 12, mergedAdpDop);
								if (num9 > num7)
								{
									excelOperations.Merge(num7, 12, num9, 12);
								}
								excelOperations.Format(num7, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
								num7 = num8 + 1;
							}
						}
						int rowHeight2 = EstimateMergedRowHeight(item.ShemeName, array[1], num12);
						EnsureMergedSchemeBodyHeight(excelOperations, num6, num5 - 1, rowHeight2);
						excelOperations.GroupRows(num4 + 1, num5 - 1, 1, hide: false);
						num4 = num5;
					}
					excelOperations.Font("Liberation Serif", num12);
					foreach (int item9 in notControlledRows)
					{
						excelOperations.FontColor(item9, 3, Color.Red);
						excelOperations.FontStyle(item9, 3, 14f, italic: true);
					}
					for (int n = 1; n <= array.Count(); n++)
					{
						excelOperations.AutoFitWithMaxWidth(n, array[n - 1]);
					}
					if (!columnMap.HasMdpPa)
					{
						excelOperations.HideColumn(5);
						excelOperations.HideColumn(8);
						excelOperations.HideColumn(11);
					}
					excelOperations.Borders(1, 1, num4 - 1, array.Count());
					excelOperations.GroupRowsPosition();
					excelOperations.UpdateSummarySheetHyperlinks("Обшая информация о сечении", "new", dictionary);
					if (!string.IsNullOrWhiteSpace(summaryB1Text))
					{
						excelOperations.SetSheetCellValue("Обшая информация о сечении", "B1", summaryB1Text, wrap: true);
					}
					excelOperations.ConfigureSheetForPrint("Обшая информация о сечении");
					excelOperations.ConfigureSheetForPrint("new", repeatTopTwoRows: true);
					string text14 = Path.Combine(Path.GetDirectoryName(text2) ?? "", Path.GetFileNameWithoutExtension(text2) + "_корр.xlsx");
					excelOperations.Save(text14);
					Console.WriteLine("Файл успешно обработан и сохранен: " + text14);
					Console.WriteLine("Работа программы успешно завершена.");
				}
				catch (Exception ex)
				{
					Console.WriteLine("Ошибка обработки файла: " + text2);
					Console.WriteLine(ex.Message);
				}
				Console.WriteLine("");
			}
		}
		else
		{
			Console.WriteLine("Пути к файлам не получены.");
		}
		Console.WriteLine("");
		Console.ReadKey();
	}

	private static List<string> GetInputPaths(string[] args)
	{
		List<string> list = new List<string>();
		if (args != null && args.Length != 0)
		{
			foreach (string arg in args)
			{
				string text = NormalizeInputPath(arg);
				if (!string.IsNullOrWhiteSpace(text))
				{
					list.Add(text);
				}
			}
			return list;
		}
		string text2 = Console.ReadLine() ?? "";
		if (string.IsNullOrWhiteSpace(text2))
		{
			return list;
		}
		foreach (string item in SplitInputPaths(text2))
		{
			string text3 = NormalizeInputPath(item);
			if (!string.IsNullOrWhiteSpace(text3))
			{
				list.Add(text3);
			}
		}
		return list;
	}

	private static IEnumerable<string> SplitInputPaths(string raw)
	{
		MatchCollection matchCollection = Regex.Matches(raw, "\"([^\"]+)\"|([^\\s]+)");
		foreach (Match item in matchCollection)
		{
			string value = item.Groups[1].Success ? item.Groups[1].Value : item.Groups[2].Value;
			if (!string.IsNullOrWhiteSpace(value))
			{
				yield return value;
			}
		}
	}

	private static string NormalizeInputPath(string path)
	{
		StringBuilder stringBuilder = new StringBuilder(path.Length);
		foreach (char c in path)
		{
			if (c == '\0')
			{
				continue;
			}
			if (!char.IsControl(c) || c == '\t')
			{
				stringBuilder.Append(c);
			}
		}
		return stringBuilder.ToString().Trim().Trim(new char[1] { '"' });
	}

	private static string LoadSummaryB1Config()
	{
		string[] array = new string[3]
		{
			Path.Combine(Directory.GetCurrentDirectory(), "summary_b1.txt"),
			Path.Combine(AppContext.BaseDirectory, "summary_b1.txt"),
			Path.Combine(AppContext.BaseDirectory, "config", "summary_b1.txt")
		};
		foreach (string text in array)
		{
			if (File.Exists(text))
			{
				string text2 = File.ReadAllText(text, Encoding.UTF8).Replace("\r\n", "\n").Replace("\n", Environment.NewLine).Trim();
				if (!string.IsNullOrWhiteSpace(text2))
				{
					return text2;
				}
			}
		}
		return "";
	}

	private static ColumnMap ResolveColumnMap(ExcelOperations ex)
	{
		HeaderScan headerScan = HeaderScan.Create(ex, 40);
		int schemeNumCol = headerScan.FindFirst((HeaderCell h) => h.Row1.Contains("№") || h.Row2.Contains("№") || h.All.Contains("nпп"), 2);
		int schemeNameCol = headerScan.FindFirst((HeaderCell h) => h.All.Contains("схемасети"), 3);
		int tnvCol = headerScan.FindFirst((HeaderCell h) => h.All.Contains("тнв"), 4);
		int mdpNoPaCol = headerScan.FindFirst((HeaderCell h) => h.HasMdpNoPa && !h.IsCriteriaGroup && !h.IsDopGroup, 5);
		int mdpPaCol = headerScan.FindFirst((HeaderCell h) => h.HasMdpPa && !h.IsCriteriaGroup && !h.IsDopGroup, -1);
		int adpCol = headerScan.FindFirst((HeaderCell h) => h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa && !h.IsCriteriaGroup && !h.IsDopGroup, (mdpPaCol != -1) ? 7 : 6);
		int mdpNoPaCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasMdpNoPa, (mdpPaCol != -1) ? 8 : 7);
		int mdpPaCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasMdpPa, (mdpPaCol != -1) ? 9 : (-1));
		int adpCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa, (mdpPaCol != -1) ? 11 : 9);
		int mdpNoPaDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasMdpNoPa, 12);
		int mdpPaDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasMdpPa, (mdpPaCol != -1) ? 13 : (-1));
		int adpDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa, 14);
		bool hasMdpPa = mdpPaCol != -1 || mdpPaCriteriaCol != -1 || mdpPaDopCol != -1;
		if (!hasMdpPa)
		{
			mdpPaCol = -1;
			mdpPaCriteriaCol = -1;
			mdpPaDopCol = -1;
		}
		return new ColumnMap
		{
			SchemeNumCol = schemeNumCol,
			SchemeNameCol = schemeNameCol,
			TnvCol = tnvCol,
			MdpNoPaCol = mdpNoPaCol,
			MdpPaCol = mdpPaCol,
			AdpCol = adpCol,
			MdpNoPaCriteriaCol = mdpNoPaCriteriaCol,
			MdpPaCriteriaCol = mdpPaCriteriaCol,
			AdpCriteriaCol = adpCriteriaCol,
			MdpNoPaDopCol = mdpNoPaDopCol,
			MdpPaDopCol = mdpPaDopCol,
			AdpDopCol = adpDopCol,
			HasMdpPa = hasMdpPa
		};
	}

	private static string NormalizeHeader(string text)
	{
		string text2 = (text ?? "").ToLowerInvariant().Replace("_x000A_", " ");
		text2 = text2.Replace('ё', 'е').Replace('º', 'о').Replace('°', 'o');
		return Regex.Replace(text2, "[^a-zа-я0-9]+", "");
	}

	private sealed class HeaderScan
	{
		private readonly List<HeaderCell> _cells;

		private HeaderScan(List<HeaderCell> cells)
		{
			_cells = cells;
		}

		public static HeaderScan Create(ExcelOperations ex, int maxCol)
		{
			List<HeaderCell> list = new List<HeaderCell>(maxCol);
			for (int i = 1; i <= maxCol; i++)
			{
				string row = NormalizeHeader(GetHeaderCellText(ex, 1, i));
				string row2 = NormalizeHeader(GetHeaderCellText(ex, 2, i));
				string row3 = NormalizeHeader(GetHeaderCellText(ex, 3, i));
				list.Add(new HeaderCell(i, row, row2, row3));
			}
			return new HeaderScan(list);
		}

		public int FindFirst(Func<HeaderCell, bool> predicate, int fallback)
		{
			foreach (HeaderCell cell in _cells)
			{
				if (predicate(cell))
				{
					return cell.Col;
				}
			}
			return fallback;
		}

		private static string GetHeaderCellText(ExcelOperations ex, int row, int col)
		{
			string str = ex.getStr(row, col);
			if (!string.IsNullOrWhiteSpace(str))
			{
				return str;
			}
			string str2 = ex.MergedCells(row, col);
			if (string.IsNullOrWhiteSpace(str2) || !str2.Contains(":"))
			{
				return str;
			}
			string text = str2.Split(new char[1] { ':' })[0];
			if (!TryParseCellAddress(text, out var row2, out var col2))
			{
				return str;
			}
			return ex.getStr(row2, col2);
		}

		private static bool TryParseCellAddress(string address, out int row, out int col)
		{
			row = 0;
			col = 0;
			Match match = Regex.Match(address, "^([A-Za-z]+)(\\d+)$");
			if (!match.Success)
			{
				return false;
			}
			if (!int.TryParse(match.Groups[2].Value, out row))
			{
				return false;
			}
			string value = match.Groups[1].Value.ToUpperInvariant();
			int num = 0;
			foreach (char c in value)
			{
				num = num * 26 + (c - 64);
			}
			col = num;
			return col > 0 && row > 0;
		}
	}

	private sealed class HeaderCell
	{
		public int Col { get; }

		public string Row1 { get; }

		public string Row2 { get; }

		public string Row3 { get; }

		public string All { get; }

		public bool HasMdpNoPa => All.Contains("мдпбезпа");

		public bool HasMdpPa => All.Contains("мдпспа");

		public bool HasAdp => All.Contains("адп");

		public bool IsCriteriaGroup => Row1.Contains("критер") || Row2.Contains("критер") || Row3.Contains("критер") || All.Contains("критер");

		public bool IsDopGroup => Row1.Contains("контрольдоп") || Row2.Contains("контрольдоп") || Row3.Contains("контрольдоп") || All.Contains("дополнит");

		public HeaderCell(int col, string row1, string row2, string row3)
		{
			Col = col;
			Row1 = row1;
			Row2 = row2;
			Row3 = row3;
			All = row1 + row2 + row3;
		}
	}

	private sealed class ColumnMap
	{
		public int SchemeNumCol { get; set; }

		public int SchemeNameCol { get; set; }

		public int TnvCol { get; set; }

		public int MdpNoPaCol { get; set; }

		public int MdpPaCol { get; set; }

		public int AdpCol { get; set; }

		public int MdpNoPaCriteriaCol { get; set; }

		public int MdpPaCriteriaCol { get; set; }

		public int AdpCriteriaCol { get; set; }

		public int MdpNoPaDopCol { get; set; }

		public int MdpPaDopCol { get; set; }

		public int AdpDopCol { get; set; }

		public bool HasMdpPa { get; set; }
	}

	public static string ReadLine(ExcelOperations ex, int bRow, int eRow, int col)
	{
		string result = "";
		for (int i = bRow; i <= eRow; i++)
		{
			if (ex.getStr(i, col) != "" && ex.getStr(i, col) != " ")
			{
				result = ex.getStr(i, col).Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			}
		}
		return result;
	}

	public static List<string> ReadDopLines(ExcelOperations ex, int bRow, int eRow, int col)
	{
		List<string> list = new List<string>();
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStr(i, col).Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			if (text != "" && text != " ")
			{
				list.Add(text);
			}
		}
		return list;
	}

	public static List<MDP> ReadLines(ExcelOperations ex, int bRow, int eRow, int col, bool modify = false)
	{
		List<MDP> list = new List<MDP>();
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStr(i, col).Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			if (text != "" && text != " ")
			{
				if (text.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
				{
					list.Add(new MDP
					{
						Num = -1,
						Criteria = text
					});
				}
				else
				{
					Match match = Regex.Match(text, "^(-?\\d+)\\)\\s*(.*)$");
					if (match.Success)
					{
						int num = Convert.ToInt32(match.Groups[1].Value);
						string text2 = match.Groups[2].Value;
						list.Add(new MDP
						{
							Num = num,
							Criteria = (modify ? CellModifyString(text2) : text2)
						});
					}
					else
					{
						list.Add(new MDP
						{
							Num = -1,
							Criteria = (modify ? CellModifyString(text) : text)
						});
					}
				}
			}
			else
			{
				list.Add(new MDP
				{
					Num = -1,
					Criteria = text
				});
			}
		}
		return list;
	}

	public static string CellModifyString(string text)
	{
		text = Regex.Replace(text, "\\bMIN\\b", "min");
		text = text.Replace("-", " - ").Replace("+", " + ").Replace(",", ", ")
			.Replace("  ", " ")
			.Replace("==", "=")
			.Replace("*", "·");
		BracketNode node = Parse(text);
		return Reconstruct(node) ?? "";
	}

	private static int EstimateMergedRowHeight(string text, int mergedWidth, int fontSize)
	{
		string[] array = (text ?? "").Replace("_x000A_", "\n").Split('\n');
		int num = Math.Max(20, (int)Math.Round((double)mergedWidth * 1.6));
		int num2 = 0;
		foreach (string text2 in array)
		{
			int num3 = Math.Max(1, text2.TrimEnd().Length);
			num2 += Math.Max(1, (int)Math.Ceiling((double)num3 / (double)num));
		}
		int num4 = Math.Max(15, (int)Math.Round(fontSize * 1.5));
		return num2 * num4 + 2;
	}

	private static void EnsureMergedSchemeBodyHeight(ExcelOperations ex, int startRow, int endRow, int minTotalHeight)
	{
		if (endRow < startRow)
		{
			return;
		}
		double num = 0.0;
		for (int i = startRow; i <= endRow; i++)
		{
			num += ex.GetRowHeightOrDefault(i, 15.0);
		}
		double num2 = minTotalHeight - num;
		if (num2 <= 0.0)
		{
			return;
		}
		int num3 = endRow - startRow + 1;
		int num4 = (int)Math.Ceiling(num2 / (double)num3);
		for (int j = startRow; j <= endRow; j++)
		{
			int height = (int)Math.Ceiling(ex.GetRowHeightOrDefault(j, 15.0)) + num4;
			ex.Height(j, Math.Max(15, height));
		}
	}

	private static string GetSingleSchemeAdpDopValue(List<TNV> tnvList)
	{
		List<string> list = new List<string>();
		foreach (TNV tnv in tnvList)
		{
			string text = string.Join(Environment.NewLine, tnv.AdpDop.Where((string x) => !string.IsNullOrWhiteSpace(x)).Select((string x) => x.Trim()));
			if (!string.IsNullOrWhiteSpace(text))
			{
				list.Add(text);
			}
		}
		List<string> list2 = list.Distinct(StringComparer.Ordinal).ToList();
		if (list2.Count == 1)
		{
			return list2[0];
		}
		return "";
	}

	private static bool IsNotControlledPhrase(string text)
	{
		return string.Equals((text ?? "").Trim(), "Не контролируется", StringComparison.OrdinalIgnoreCase);
	}

	private static string GetSchemeHeaderLine(string shemeName)
	{
		string text = (shemeName ?? "").Replace("_x000A_", " ").Replace('\n', ' ').Replace('\r', ' ');
		text = Regex.Replace(text, "\\s+", " ").Trim();
		return text;
	}

	public static bool AreBracketsBalanced(string input)
	{
		Stack<char> stack = new Stack<char>();
		Dictionary<char, char> dictionary = new Dictionary<char, char>
		{
			{ ')', '(' },
			{ ']', '[' },
			{ '}', '{' }
		};
		foreach (char c in input)
		{
			if (Enumerable.Contains("([{", c))
			{
				stack.Push(c);
			}
			else if (Enumerable.Contains(")]}", c) && (stack.Count == 0 || stack.Pop() != dictionary[c]))
			{
				return false;
			}
		}
		return stack.Count == 0;
	}

	public static BracketNode Parse(string input, char open = '(', char close = ')')
	{
		if (string.IsNullOrEmpty(input))
		{
			return new BracketNode();
		}
		if (!AreBracketsBalanced(input))
		{
			throw new ArgumentException("Несбалансированные скобки");
		}
		int index = 0;
		return ParseRecursive(input, ref index, open, close);
	}

	private static BracketNode ParseRecursive(string input, ref int index, char open, char close)
	{
		BracketNode bracketNode = new BracketNode();
		StringBuilder stringBuilder = new StringBuilder();
		while (index < input.Length)
		{
			if (input[index] == open)
			{
				if (stringBuilder.Length > 0)
				{
					bracketNode.ContentParts.Add(stringBuilder.ToString());
					stringBuilder.Clear();
				}
				index++;
				bracketNode.ContentParts.Add(ParseRecursive(input, ref index, open, close));
				continue;
			}
			if (input[index] == close)
			{
				if (stringBuilder.Length > 0)
				{
					bracketNode.ContentParts.Add(stringBuilder.ToString());
					stringBuilder.Clear();
				}
				index++;
				return bracketNode;
			}
			stringBuilder.Append(input[index]);
			index++;
		}
		if (stringBuilder.Length > 0)
		{
			bracketNode.ContentParts.Add(stringBuilder.ToString());
		}
		return bracketNode;
	}

	public static void PrintTree(BracketNode node, string prefix = "", bool isLast = true)
	{
		Console.WriteLine(prefix + (isLast ? "└─ " : "├─ ") + "Node");
		string text = prefix + (isLast ? "    " : "│   ");
		for (int i = 0; i < node.ContentParts.Count; i++)
		{
			object obj = node.ContentParts[i];
			bool flag = i == node.ContentParts.Count - 1;
			if (obj is string text2)
			{
				Console.WriteLine(text + (flag ? "└─ " : "├─ ") + "Text: \"" + text2 + "\"");
			}
			else if (obj is BracketNode node2)
			{
				PrintTree(node2, text, flag);
			}
		}
	}

	public static string Reconstruct(BracketNode node, string bracket = "|(|")
	{
		StringBuilder stringBuilder = new StringBuilder();
		foreach (object contentPart in node.ContentParts)
		{
			if (contentPart is string text)
			{
				stringBuilder.Append(text);
				bracket = ((!text.Contains("if")) ? ((!text.Contains("min") && !text.Contains("max")) ? "|(|" : "|[|") : "|{|");
			}
			else if (contentPart is BracketNode node2)
			{
				string text2 = bracket;
				string text3 = ((text2 == "|{|") ? "|}|" : ((text2 == "|[|") ? "|]|" : "|)|"));
				stringBuilder.Append(text2 + Reconstruct(node2) + text3);
			}
		}
		return stringBuilder.ToString();
	}
}
