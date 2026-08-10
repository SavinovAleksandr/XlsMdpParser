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
		Console.WriteLine("Перетащите один или несколько Excel-файлов или папку в это окно и нажмите Enter:");
		List<InputJob> inputJobs = BuildInputJobs(args);
		string summaryB1Text = LoadSummaryB1Config();
		if (inputJobs.Count > 0)
		{
			foreach (InputJob inputJob in inputJobs)
			{
				string text2 = inputJob.InputPath;
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
						if (!TryGetSchemeRowSpan(excelOperations, i, columnMap.SchemeNameCol, out int num10, out int num11))
						{
							num10 = i;
							num11 = i;
						}
						List<TNV> list2 = ReadSchemeTnvBlocks(excelOperations, columnMap, num10, num11);
						if (list2.Count == 0)
						{
							list2.Add(ReadTnvBlock(excelOperations, columnMap, num10, num11, rowLabel: ""));
						}
						list2 = ConsolidateTnvByTemperature(list2, text2);
						EnsureDdtnControlAnnotations(list2, str2, str);
						list.Add(new MdpBuilder
						{
							ShemeName = str,
							ShemeNum = str2,
							TnvList = list2
						});
					}
					// Сначала старый лист «Ремонтные схемы» → old, затем создаём новый с тем же именем
					// (чтобы гиперссылки со сводки сразу указывали на итог и не ломались при ручном переименовании).
					excelOperations.RenameSheet("Ремонтные схемы", "old");
					excelOperations.AddList("Ремонтные схемы");
					const string outputSheetName = "Ремонтные схемы";
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
					excelOperations.setVal(1, 3, columnMap.HasArpm && !columnMap.HasTnv ? "гр. уст. АРПМ" : "ТНВ, °С");
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
					OutputColumnFlags outCols = DetectOutputColumnFlags(list);
					if (!columnMap.HasTnv && !columnMap.HasArpm)
					{
						excelOperations.HideColumn(3);
					}
					HideEmptyOutputColumns(excelOperations, outCols);
					excelOperations.FormatCells(1, 1, 2, array.Count(), bold: true, italic: false, Color.PowderBlue.ToArgb());
					int num4 = 3;
					Dictionary<string, int> dictionary = new Dictionary<string, int>();
					List<int> notControlledRows = new List<int>();
					Dictionary<string, Color> criteriaColorMap = BuildCriteriaColorMap(list);
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
						excelOperations.setVal(num4, 2, text3, wrap: true);
						excelOperations.Merge(num4, 2, num4, array.Count());
						excelOperations.Format(num4, 2, ExcelHorizontalAlignment.Justify, ExcelVerticalAlignment.Center);
						excelOperations.Wrap(num4, 2, wrap: true);
						excelOperations.FormatCells(num4, 1, num4, array.Count(), bold: false, italic: false, Color.MistyRose.ToArgb());
						int textWidth = array.Skip(1).Sum();
						int rowHeight = EstimateMergedRowHeight(text3, textWidth, num12);
						excelOperations.Height(num4, Math.Max(20, rowHeight));
						int num5 = num4 + 1;
						int num6 = num5;
						string mergedAdpDop = outCols.HasCtrlAdp ? GetSingleSchemeAdpDopValue(item.TnvList) : "";
						bool mergeAdpDop = !string.IsNullOrWhiteSpace(mergedAdpDop);
						HashSet<int> hashSet = new HashSet<int>();
						int tnvCount = Math.Max(1, item.TnvList.Count);
						bool schemeHasAdp = item.TnvList.Any((TNV t) => !string.IsNullOrWhiteSpace(t.Adp) && !IsDashOrEmpty(t.Adp));
						excelOperations.setVal(num5, 1, item.ShemeNum);
						excelOperations.Merge(num5, 1, num5 + tnvCount - 1, 1);
						excelOperations.Format(num5, 1, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
						excelOperations.setVal(num5, 2, item.ShemeName);
						excelOperations.Merge(num5, 2, num5 + tnvCount - 1, 2);
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
							List<MDP> list3 = tnv.MdpNoPA.Where((MDP mDP) => HasMeaningfulMdpText(mDP.Criteria)).ToList();
							List<MDP> list4 = list3.Where((MDP mDP) => mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							List<MDP> list5 = list3.Where((MDP mDP) => !mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).OrderBy((MDP mDP) => (mDP.Num >= 0) ? mDP.Num : int.MaxValue).ToList();
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
							List<MDP> list11 = tnv.MdpNoPaCriteria.Where((MDP mDP) => HasMeaningfulMdpText(mDP.Criteria) && !IsDashOrEmpty(mDP.Criteria)).OrderBy((MDP mDP) => (mDP.Num >= 0) ? mDP.Num : int.MaxValue).ToList();
							if (outCols.HasMdpNoPa)
							{
								WriteColoredMdpBlocks(excelOperations, num5, 4, list4.Concat(list5).ToList(), list11, criteriaColorMap);
							}
							else
							{
								excelOperations.setVal(num5, 4, "");
							}
							List<MDP> list7 = tnv.MdpPa.Where((MDP mDP) => HasMeaningfulMdpText(mDP.Criteria)).ToList();
							List<MDP> list8 = list7.Where((MDP mDP) => mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							List<MDP> list9 = list7.Where((MDP mDP) => !mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase)).ToList();
							bool hasPaSections = list9.Any((MDP m) => m.Criteria.StartsWith("—", StringComparison.Ordinal) || m.Criteria.StartsWith("[", StringComparison.Ordinal));
							if (!hasPaSections)
							{
								list9 = list9.OrderBy((MDP mDP) => (mDP.Num >= 0) ? mDP.Num : int.MaxValue).ToList();
							}
							if (list9.Count(m => !m.Criteria.StartsWith("—", StringComparison.Ordinal) && !m.Criteria.StartsWith("[", StringComparison.Ordinal)) <= 1)
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
							List<MDP> list12 = tnv.MdpPaCriteria.Where((MDP mDP) => HasMeaningfulMdpText(mDP.Criteria) && !IsDashOrEmpty(mDP.Criteria)).ToList();
							bool hasPaCritSections = list12.Any((MDP m) => m.Criteria.StartsWith("—", StringComparison.Ordinal) || m.Criteria.StartsWith("[", StringComparison.Ordinal));
							if (!hasPaCritSections)
							{
								list12 = list12.OrderBy((MDP mDP) => (mDP.Num >= 0) ? mDP.Num : int.MaxValue).ToList();
							}
							if (outCols.HasMdpPa)
							{
								WriteColoredMdpBlocks(excelOperations, num5, 5, list8.Concat(list9).ToList(), list12, criteriaColorMap);
							}
							else
							{
								excelOperations.setVal(num5, 5, "");
								excelOperations.Format(num5, 5, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
							}
							if (outCols.HasAdp && !IsDashOrEmpty(tnv.Adp))
							{
								excelOperations.setVal(num4 + 1, 6, tnv.Adp, wrap: true);
								excelOperations.Merge(num4 + 1, 6, num4 + item.TnvList.Count, 6);
								bool flagAdpMultiline = tnv.Adp.Contains('\n') || tnv.Adp.Contains('\r');
								// Top: иначе короткий текст в высоком merge визуально «прилипает» к средней ТНВ (+15).
								excelOperations.Format(num4 + 1, 6, flagAdpMultiline ? ExcelHorizontalAlignment.Left : ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Top);
							}
							if (outCols.HasMdpNoPa)
							{
								string text5 = WriteColoredCriteriaBlocks(excelOperations, num5, 7, list11, criteriaColorMap);
								excelOperations.CellComment(num5, 4, text5);
							}
							else
							{
								excelOperations.setVal(num5, 7, "");
							}
							if (outCols.HasMdpPa)
							{
								string text7 = WriteColoredCriteriaBlocks(excelOperations, num5, 8, list12, criteriaColorMap);
								excelOperations.CellComment(num5, 5, text7);
							}
							else
							{
								excelOperations.setVal(num5, 8, "");
							}
							if (schemeHasAdp && outCols.HasAdp && !IsDashOrEmpty(tnv.AdpCriteria))
							{
								excelOperations.setVal(num4 + 1, 9, tnv.AdpCriteria, wrap: true);
								excelOperations.Merge(num4 + 1, 9, num4 + item.TnvList.Count, 9);
								bool flagAdpCritMultiline = tnv.AdpCriteria.Contains('\n') || tnv.AdpCriteria.Contains('\r');
								excelOperations.Format(num4 + 1, 9, flagAdpCritMultiline ? ExcelHorizontalAlignment.Left : ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Top);
							}
							string text9 = "";
							if (outCols.HasCtrlMdpNoPa)
							{
								foreach (string item6 in tnv.MdpNoPaDop)
								{
									string text10 = ((item6 == tnv.MdpNoPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
									text9 = text9 + item6 + text10;
								}
							}
							excelOperations.setVal(num5, 10, text9);
							excelOperations.Format(num5, 10, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							string text11 = "";
							if (outCols.HasCtrlMdpPa)
							{
								foreach (string item7 in tnv.MdpPaDop)
								{
									string text12 = ((item7 == tnv.MdpPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
									text11 = text11 + item7 + text12;
								}
							}
							excelOperations.setVal(num5, 11, text11);
							excelOperations.Format(num5, 11, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
							string text13 = "";
							if (!mergeAdpDop)
							{
								foreach (string item8 in tnv.AdpDop)
								{
									string text14Line = ((item8 == tnv.AdpDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
									text13 = text13 + item8 + text14Line;
								}
							}
							excelOperations.setVal(num5, 12, text13);
							excelOperations.Format(num5, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Top);
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
								excelOperations.setVal(num7, 12, mergedAdpDop, wrap: true);
								if (num9 > num7)
								{
									excelOperations.Merge(num7, 12, num9, 12);
								}
								excelOperations.Format(num7, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Top);
								num7 = num8 + 1;
							}
						}
						int rowHeight2 = EstimateMergedRowHeight(item.ShemeName, array[1], num12);
						if (num5 - 1 >= num6)
						{
							EnsureMergedSchemeBodyHeight(excelOperations, num6, num5 - 1, rowHeight2);
							excelOperations.GroupRows(num4 + 1, num5 - 1, 1, hide: false);
						}
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
					// AutoFit снимает Hidden — прячем колонки снова.
					if (!columnMap.HasTnv && !columnMap.HasArpm)
					{
						excelOperations.HideColumn(3);
					}
					HideEmptyOutputColumns(excelOperations, outCols);
					// Высота строк строго по тексту во всех колонках; ширину не трогаем.
					excelOperations.AutoFitSheetRowsByContent(outputSheetName, 3, 15, 1.0, new int[11] { 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 });
					excelOperations.Borders(1, 1, num4 - 1, array.Count());
					excelOperations.GroupRowsPosition();
					excelOperations.UpdateSummarySheetHyperlinks("Обшая информация о сечении", outputSheetName, dictionary);
					if (!string.IsNullOrWhiteSpace(summaryB1Text))
					{
						// Справа (C1), в том же стиле, что исходный блок «УТВЕРЖДАЮ» (Говорун).
						excelOperations.SetSummaryApprovalBlock("Обшая информация о сечении", "C1", summaryB1Text);
					}
					excelOperations.SetSheetCellAlignment("Обшая информация о сечении", "A3", ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					// С 1-й строки: блок УТВЕРЖДАЮ (C1), заголовок (B3) и весь текст сводки.
					excelOperations.AutoFitSheetRowsByContent("Обшая информация о сечении", 1, 15, 1.0);
					excelOperations.ConfigureSheetForPrint("Обшая информация о сечении");
					excelOperations.ConfigureSheetForPrint(outputSheetName, repeatTopTwoRows: true);
					// После AutoFit/печати Hidden мог сброситься — прячем снова.
					if (!columnMap.HasTnv && !columnMap.HasArpm)
					{
						excelOperations.HideColumn(3);
					}
					HideEmptyOutputColumns(excelOperations, outCols);
					if (!Directory.Exists(inputJob.OutputDirectory))
					{
						Directory.CreateDirectory(inputJob.OutputDirectory);
					}
					string text14 = Path.Combine(inputJob.OutputDirectory, Path.GetFileNameWithoutExtension(text2) + "_корр.xlsx");
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
			Console.WriteLine("Пути к файлам/папкам не получены.");
		}
		Console.WriteLine("");
		WaitForExitKeyIfInteractive();
	}

	private static void WaitForExitKeyIfInteractive()
	{
		try
		{
			if (!Console.IsInputRedirected)
			{
				Console.ReadKey();
			}
		}
		catch (InvalidOperationException)
		{
		}
	}

	private static List<InputJob> BuildInputJobs(string[] args)
	{
		List<string> inputPaths = GetInputPaths(args);
		List<InputJob> list = new List<InputJob>();
		foreach (string inputPath in inputPaths)
		{
			if (File.Exists(inputPath))
			{
				list.Add(new InputJob
				{
					InputPath = inputPath,
					OutputDirectory = (Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory())
				});
				continue;
			}
			if (Directory.Exists(inputPath))
			{
				string text = Path.Combine(inputPath, "_корр");
				foreach (string item in Directory.GetFiles(inputPath, "*.xlsx", SearchOption.TopDirectoryOnly).OrderBy((string p) => p, StringComparer.OrdinalIgnoreCase))
				{
					if (!Path.GetFileName(item).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
					{
						list.Add(new InputJob
						{
							InputPath = item,
							OutputDirectory = text
						});
					}
				}
			}
		}
		return list;
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
		List<string> list = new List<string>();
		StringBuilder stringBuilder = new StringBuilder();
		bool flag = false;
		for (int i = 0; i < raw.Length; i++)
		{
			char c = raw[i];
			if (c == '"')
			{
				flag = !flag;
				continue;
			}
			if (!flag && char.IsWhiteSpace(c))
			{
				if (stringBuilder.Length > 0)
				{
					list.Add(stringBuilder.ToString());
					stringBuilder.Clear();
				}
				continue;
			}
			if (c == '\\' && i + 1 < raw.Length)
			{
				char c2 = raw[i + 1];
				if (char.IsWhiteSpace(c2) || c2 == '"' || c2 == '\\')
				{
					stringBuilder.Append(c2);
					i++;
					continue;
				}
			}
			stringBuilder.Append(c);
		}
		if (stringBuilder.Length > 0)
		{
			list.Add(stringBuilder.ToString());
		}
		return list;
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
		HeaderScan headerScan = HeaderScan.Create(ex, 50);
		int schemeNumCol = headerScan.FindFirst((HeaderCell h) => h.Row1.Contains("№") || h.Row2.Contains("№") || h.All.Contains("nпп"), 2);
		int schemeNameCol = headerScan.FindFirst((HeaderCell h) => h.All.Contains("схемасети"), 3);
		int tnvCol = headerScan.FindFirst((HeaderCell h) => h.All.Contains("тнв"), -1);
		int arpmCol = headerScan.FindFirst((HeaderCell h) => h.All.Contains("арпм") && !h.HasMdpNoPa && !h.HasMdpPa && !h.HasAdp && !h.IsCriteriaGroup && !h.IsDopGroup, -1);
		int mdpNoPaCol = headerScan.FindFirst((HeaderCell h) => h.HasMdpNoPa && !h.IsCriteriaGroup && !h.IsDopGroup, 5);
		int mdpPaCol = headerScan.FindFirst((HeaderCell h) => h.HasMdpPa && !h.IsCriteriaGroup && !h.IsDopGroup, -1);
		// АОПО в зоне МДП (после ТНВ / рядом с «МДП без ПА»). Короткие «АОПО В-Х» до ТНВ — не МДП с ПА.
		int aopoZoneStart = Math.Max(tnvCol > 0 ? tnvCol + 1 : 0, mdpNoPaCol > 0 ? mdpNoPaCol : 0);
		if (aopoZoneStart <= 0)
		{
			aopoZoneStart = Math.Max(schemeNameCol + 1, 4);
		}
		List<PaColumn> aopoValueCols = FilterAopoInZone(headerScan.FindAll((HeaderCell h) => h.IsAopo && !h.IsCriteriaGroup && !h.IsDopGroup && !h.HasAdp), aopoZoneStart, mdpNoPaCol);
		List<PaColumn> aopoCriteriaCols = FilterAopoInZone(headerScan.FindAll((HeaderCell h) => h.IsAopo && h.IsCriteriaGroup), aopoZoneStart, -1);
		List<PaColumn> aopoDopCols = FilterAopoInZone(headerScan.FindAll((HeaderCell h) => h.IsAopo && h.IsDopGroup), aopoZoneStart, -1);
		if (mdpPaCol == -1 && aopoValueCols.Count > 0)
		{
			mdpPaCol = aopoValueCols[0].Col;
			aopoValueCols = aopoValueCols.Skip(1).ToList();
		}
		else if (mdpPaCol != -1)
		{
			aopoValueCols = aopoValueCols.Where((PaColumn p) => p.Col != mdpPaCol).ToList();
		}
		int adpCol = headerScan.FindFirst((HeaderCell h) => h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa && !h.IsAopo && !h.IsCriteriaGroup && !h.IsDopGroup, -1);
		if (adpCol == -1)
		{
			adpCol = headerScan.FindFirst((HeaderCell h) => h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa && !h.IsCriteriaGroup && !h.IsDopGroup, (mdpPaCol != -1 || aopoValueCols.Count > 0) ? 7 : 6);
		}
		int mdpNoPaCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasMdpNoPa, -1);
		int mdpPaCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasMdpPa, -1);
		int adpCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa && !h.IsAopo, -1);
		if (adpCriteriaCol == -1)
		{
			adpCriteriaCol = headerScan.FindFirst((HeaderCell h) => h.IsCriteriaGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa, -1);
		}
		// Как в HTML: критерии «с ПА» — заполненные колонки между «без ПА» и «АДП»
		// (значение часто в левой ячейке merge, а подпись «МДП с ПА» — справа).
		// Важно: колонки, входящие в merge с заголовком/ячейкой критерия АДП (типично H:I
		// при подписи «АДП» на I), — это АДП, а не «МДП с ПА» (Сыктывкар и аналоги).
		List<int> paCriteriaByContent = FindPopulatedColumnsBetween(ex, mdpNoPaCriteriaCol, adpCriteriaCol, bodyStartRow: 4)
			.Where((int c) => !ColumnSharesMergeWith(ex, c, adpCriteriaCol, bodyStartRow: 4))
			.ToList();
		if (paCriteriaByContent.Count > 0)
		{
			if (mdpPaCriteriaCol == -1 || !ColumnHasMeaningfulBody(ex, mdpPaCriteriaCol, 4))
			{
				mdpPaCriteriaCol = paCriteriaByContent[0];
			}
			aopoCriteriaCols = paCriteriaByContent
				.Where((int c) => c != mdpPaCriteriaCol)
				.Select((int c) => new PaColumn { Col = c, Title = "" })
				.ToList();
		}
		else if (mdpPaCriteriaCol == -1 && aopoCriteriaCols.Count > 0)
		{
			mdpPaCriteriaCol = aopoCriteriaCols[0].Col;
			aopoCriteriaCols = aopoCriteriaCols.Skip(1).ToList();
		}
		else if (mdpPaCriteriaCol != -1)
		{
			aopoCriteriaCols = aopoCriteriaCols.Where((PaColumn p) => p.Col != mdpPaCriteriaCol).ToList();
		}
		else
		{
			// Между «без ПА» и «АДП» пусто (или только merge АДП) — сбрасываем ложный PA.
			mdpPaCriteriaCol = -1;
			aopoCriteriaCols = new List<PaColumn>();
		}
		int mdpNoPaDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasMdpNoPa, -1);
		int mdpPaDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasMdpPa, -1);
		if (mdpPaDopCol == -1 && aopoDopCols.Count > 0)
		{
			mdpPaDopCol = aopoDopCols[0].Col;
			aopoDopCols = aopoDopCols.Skip(1).ToList();
		}
		int adpDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa && !h.IsAopo, -1);
		if (adpDopCol == -1)
		{
			adpDopCol = headerScan.FindFirst((HeaderCell h) => h.IsDopGroup && h.HasAdp && !h.HasMdpNoPa && !h.HasMdpPa, -1);
		}
		bool hasMdpPa = mdpPaCol != -1 || mdpPaCriteriaCol != -1 || mdpPaDopCol != -1 || aopoValueCols.Count > 0 || aopoCriteriaCols.Count > 0;
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
			ArpmCol = arpmCol,
			HasTnv = tnvCol != -1,
			HasArpm = arpmCol != -1,
			MdpNoPaCol = mdpNoPaCol,
			MdpPaCol = mdpPaCol,
			AdpCol = adpCol,
			MdpNoPaCriteriaCol = mdpNoPaCriteriaCol,
			MdpPaCriteriaCol = mdpPaCriteriaCol,
			AdpCriteriaCol = adpCriteriaCol,
			MdpNoPaDopCol = mdpNoPaDopCol,
			MdpPaDopCol = mdpPaDopCol,
			AdpDopCol = adpDopCol,
			HasMdpPa = hasMdpPa,
			ExtraPaValueCols = aopoValueCols,
			ExtraPaCriteriaCols = aopoCriteriaCols,
			ExtraPaDopCols = aopoDopCols
		};
	}

	private static List<PaColumn> FilterAopoInZone(List<HeaderCell> cells, int zoneStart, int excludeCol)
	{
		return cells
			.Where((HeaderCell h) => h.Col >= zoneStart && h.Col != excludeCol)
			.Select((HeaderCell h) => new PaColumn { Col = h.Col, Title = GetShortPaTitle(h) })
			.ToList();
	}

	private static List<int> FindPopulatedColumnsBetween(ExcelOperations ex, int startExclusive, int endExclusive, int bodyStartRow)
	{
		List<int> list = new List<int>();
		if (startExclusive < 0 || endExclusive < 0 || endExclusive <= startExclusive + 1)
		{
			return list;
		}
		for (int col = startExclusive + 1; col < endExclusive; col++)
		{
			if (ColumnHasMeaningfulBody(ex, col, bodyStartRow))
			{
				list.Add(col);
			}
		}
		return list;
	}

	private static bool ColumnSharesMergeWith(ExcelOperations ex, int col, int otherCol, int bodyStartRow)
	{
		if (col <= 0 || otherCol <= 0 || col == otherCol)
		{
			return col > 0 && otherCol > 0 && col == otherCol;
		}
		int last = ex.LastColumnRow();
		int limit = Math.Min(last, bodyStartRow + 400);
		for (int row = bodyStartRow; row <= limit; row++)
		{
			if (ex.SharesMerge(row, col, otherCol))
			{
				return true;
			}
		}
		return false;
	}

	private static bool ColumnHasMeaningfulBody(ExcelOperations ex, int col, int bodyStartRow)
	{
		if (col <= 0)
		{
			return false;
		}
		int last = ex.LastColumnRow();
		for (int row = bodyStartRow; row <= last; row++)
		{
			// Только «свои» значения — иначе merge-slave (пустая правая ячейка объединения) даст ложный столбец.
			if (!ex.CellHasOwnValue(row, col))
			{
				continue;
			}
			string text = StripLeadingItemNumber(ex.getStr(row, col));
			if (HasMeaningfulMdpText(text))
			{
				return true;
			}
		}
		return false;
	}

	private static bool HasMeaningfulMdpItems(IEnumerable<MDP> items)
	{
		return items != null && items.Any((MDP m) => HasMeaningfulMdpText(m.Criteria));
	}

	private static bool HasMeaningfulMdpText(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string trimmed = text.Trim();
		if (IsDashOrEmpty(trimmed))
		{
			return false;
		}
		if (IsNotControlledPhrase(trimmed))
		{
			return false;
		}
		return true;
	}

	private static bool IsDashOrEmpty(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return true;
		}
		string full = text.Replace("_x000A_", " ").Trim();
		if (full == "-" || full == "—" || full == "–" || full == "−")
		{
			return true;
		}
		// «1) –» / «1) —» без содержательного критерия
		return Regex.IsMatch(full, "^\\d+\\)\\s*[-—–−]\\s*$");
	}

	private static string GetShortPaTitle(HeaderCell cell)
	{
		string text = (cell.Raw1 ?? "").Trim();
		if (string.IsNullOrWhiteSpace(text))
		{
			text = (cell.Raw2 ?? "").Trim();
		}
		if (string.IsNullOrWhiteSpace(text))
		{
			text = (cell.Raw3 ?? "").Trim();
		}
		if (text.Length > 48)
		{
			text = text.Substring(0, 48).Trim() + "…";
		}
		return string.IsNullOrWhiteSpace(text) ? ("АОПО " + cell.Col) : text;
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
				string raw1 = GetHeaderCellText(ex, 1, i);
				string raw2 = GetHeaderCellText(ex, 2, i);
				string raw3 = GetHeaderCellText(ex, 3, i);
				list.Add(new HeaderCell(i, NormalizeHeader(raw1), NormalizeHeader(raw2), NormalizeHeader(raw3), raw1, raw2, raw3));
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

		public List<HeaderCell> FindAll(Func<HeaderCell, bool> predicate)
		{
			List<HeaderCell> list = new List<HeaderCell>();
			foreach (HeaderCell cell in _cells)
			{
				if (predicate(cell))
				{
					list.Add(cell);
				}
			}
			return list;
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

		public string Raw1 { get; }

		public string Raw2 { get; }

		public string Raw3 { get; }

		public string All { get; }

		public bool HasMdpNoPa => All.Contains("мдпбезпа");

		public bool HasMdpPa => All.Contains("мдпспа");

		public bool HasAdp => All.Contains("адп");

		public bool IsAopo => All.Contains("аопо");

		public bool IsCriteriaGroup => Row1.Contains("критер") || Row2.Contains("критер") || Row3.Contains("критер") || All.Contains("критер");

		public bool IsDopGroup => Row1.Contains("контрольдоп") || Row2.Contains("контрольдоп") || Row3.Contains("контрольдоп") || All.Contains("дополнит");

		public HeaderCell(int col, string row1, string row2, string row3, string raw1 = "", string raw2 = "", string raw3 = "")
		{
			Col = col;
			Row1 = row1;
			Row2 = row2;
			Row3 = row3;
			Raw1 = raw1 ?? "";
			Raw2 = raw2 ?? "";
			Raw3 = raw3 ?? "";
			All = row1 + row2 + row3;
		}
	}

	private sealed class PaColumn
	{
		public int Col { get; set; }

		public string Title { get; set; } = "";
	}

	private sealed class OutputColumnFlags
	{
		public bool HasMdpNoPa { get; set; }

		public bool HasMdpPa { get; set; }

		public bool HasAdp { get; set; }

		public bool HasCtrlMdpNoPa { get; set; }

		public bool HasCtrlMdpPa { get; set; }

		public bool HasCtrlAdp { get; set; }
	}

	private static OutputColumnFlags DetectOutputColumnFlags(List<MdpBuilder> schemes)
	{
		OutputColumnFlags flags = new OutputColumnFlags();
		if (schemes == null)
		{
			return flags;
		}
		foreach (MdpBuilder scheme in schemes)
		{
			foreach (TNV tnv in scheme.TnvList ?? new List<TNV>())
			{
				if (IsNotControlledPhrase(tnv.Tnv))
				{
					continue;
				}
				if (HasMeaningfulMdpItems(tnv.MdpNoPA))
				{
					// Только по значениям «МДП без ПА». Критерии без значений (типично АРПМ:
					// дубль «Исключение…» в колонке «без ПА») не удерживают столбец видимым.
					flags.HasMdpNoPa = true;
				}
				if (HasMeaningfulMdpItems(tnv.MdpPa) || HasMeaningfulMdpItems(tnv.MdpPaCriteria))
				{
					flags.HasMdpPa = true;
				}
				if ((tnv.MdpPaDop ?? new List<string>()).Any(HasMeaningfulDopText))
				{
					flags.HasMdpPa = true;
					flags.HasCtrlMdpPa = true;
				}
				if (!IsDashOrEmpty(tnv.Adp) && !IsNotControlledPhrase(tnv.Adp))
				{
					flags.HasAdp = true;
				}
				if ((tnv.AdpDop ?? new List<string>()).Any(HasMeaningfulDopText))
				{
					flags.HasAdp = true;
					flags.HasCtrlAdp = true;
				}
				if ((tnv.MdpNoPaDop ?? new List<string>()).Any(HasMeaningfulDopText))
				{
					flags.HasCtrlMdpNoPa = true;
				}
			}
		}
		return flags;
	}

	private static void HideEmptyOutputColumns(ExcelOperations excelOperations, OutputColumnFlags flags)
	{
		if (!flags.HasMdpNoPa)
		{
			excelOperations.HideColumn(4);
			excelOperations.HideColumn(7);
		}
		if (!flags.HasMdpPa)
		{
			excelOperations.HideColumn(5);
			excelOperations.HideColumn(8);
		}
		if (!flags.HasAdp)
		{
			excelOperations.HideColumn(6);
			excelOperations.HideColumn(9);
		}
		if (!flags.HasCtrlMdpNoPa)
		{
			excelOperations.HideColumn(10);
		}
		if (!flags.HasCtrlMdpPa)
		{
			excelOperations.HideColumn(11);
		}
		if (!flags.HasCtrlAdp)
		{
			excelOperations.HideColumn(12);
		}
		// Если все три «контроля» пусты — группу уже скрыли по колонкам 10–12.
	}

	private sealed class ColumnMap
	{
		public int SchemeNumCol { get; set; }

		public int SchemeNameCol { get; set; }

		public int TnvCol { get; set; }

		public int ArpmCol { get; set; }

		public bool HasTnv { get; set; }

		public bool HasArpm { get; set; }

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

		public List<PaColumn> ExtraPaValueCols { get; set; } = new List<PaColumn>();

		public List<PaColumn> ExtraPaCriteriaCols { get; set; } = new List<PaColumn>();

		public List<PaColumn> ExtraPaDopCols { get; set; } = new List<PaColumn>();

		public int BlockMarkerCol
		{
			get
			{
				if (HasTnv)
				{
					return TnvCol;
				}
				if (HasArpm)
				{
					return ArpmCol;
				}
				return -1;
			}
		}
	}

	private sealed class InputJob
	{
		public string InputPath { get; set; } = "";

		public string OutputDirectory { get; set; } = "";
	}

	private static bool TryGetSchemeRowSpan(ExcelOperations ex, int row, int schemeNameCol, out int startRow, out int endRow)
	{
		startRow = row;
		endRow = row;
		string text = ex.MergedCells(row, schemeNameCol);
		if (string.IsNullOrWhiteSpace(text) || !text.Contains(":"))
		{
			return false;
		}
		string[] parts = text.Split(':');
		if (parts.Length < 2)
		{
			return false;
		}
		Match m1 = Regex.Match(parts[0], "(\\d+)$");
		Match m2 = Regex.Match(parts[1], "(\\d+)$");
		if (!m1.Success || !m2.Success)
		{
			return false;
		}
		startRow = Convert.ToInt32(m1.Groups[1].Value);
		endRow = Convert.ToInt32(m2.Groups[1].Value);
		if (endRow < startRow)
		{
			int tmp = startRow;
			startRow = endRow;
			endRow = tmp;
		}
		return true;
	}

	private static List<TNV> ReadSchemeTnvBlocks(ExcelOperations ex, ColumnMap columnMap, int schemeStart, int schemeEnd)
	{
		List<TNV> list = new List<TNV>();
		int markerCol = columnMap.BlockMarkerCol;
		if (markerCol == -1)
		{
			list.Add(ReadTnvBlock(ex, columnMap, schemeStart, schemeEnd, rowLabel: ""));
			return list;
		}
		for (int j = schemeStart; j <= schemeEnd; )
		{
			while (j <= schemeEnd && string.IsNullOrWhiteSpace(ex.getStr(j, markerCol)))
			{
				j++;
			}
			if (j > schemeEnd)
			{
				break;
			}
			int bRow = j;
			int eRow = j;
			while (eRow < schemeEnd && string.IsNullOrWhiteSpace(ex.getStr(eRow + 1, markerCol)))
			{
				eRow++;
			}
			string label = columnMap.HasTnv
				? ReadLine(ex, bRow, eRow, columnMap.TnvCol)
				: ReadLine(ex, bRow, eRow, columnMap.ArpmCol);
			list.Add(ReadTnvBlock(ex, columnMap, bRow, eRow, label));
			j = eRow + 1;
		}
		return list;
	}

	/// <summary>
	/// Одна строка на уникальную ТНВ: повторные блоки (группы АОПО/стабилизации)
	/// склеиваются в «МДП с ПА» с разделителями групп и сквозной нумерацией.
	/// Для Вологда–Архангельск (и ВР): Группа 1/2/3 → летняя / зимняя / весенне-осенняя уставка.
	/// </summary>
	private static List<TNV> ConsolidateTnvByTemperature(List<TNV> blocks, string inputPath = null)
	{
		Func<int, string> groupLabel = ResolvePaGroupLabelFormatter(inputPath);
		if (blocks == null || blocks.Count <= 1)
		{
			if (blocks != null && blocks.Count == 1)
			{
				FinalizeSingleTnvPaNumbering(blocks[0]);
			}
			return blocks ?? new List<TNV>();
		}
		Dictionary<string, List<TNV>> grouped = new Dictionary<string, List<TNV>>(StringComparer.OrdinalIgnoreCase);
		List<string> order = new List<string>();
		for (int i = 0; i < blocks.Count; i++)
		{
			TNV block = blocks[i];
			string key = NormalizeTnvKey(block.Tnv);
			if (string.IsNullOrEmpty(key))
			{
				key = "__row_" + i;
			}
			if (!grouped.ContainsKey(key))
			{
				order.Add(key);
				grouped[key] = new List<TNV>();
			}
			grouped[key].Add(block);
		}
		List<TNV> result = new List<TNV>();
		foreach (string key in order)
		{
			List<TNV> same = grouped[key];
			result.Add(same.Count == 1
				? FinalizeSingleTnvPaNumbering(same[0])
				: MergeTnvTemperatureGroup(same, groupLabel, factorSharedPa: IsVologdaArkhangelskSection(inputPath)));
		}
		return result;
	}

	private static Func<int, string> ResolvePaGroupLabelFormatter(string inputPath)
	{
		if (IsVologdaArkhangelskSection(inputPath))
		{
			return (int groupIndex) =>
			{
				string name = groupIndex switch
				{
					1 => "летняя уставка",
					2 => "зимняя уставка",
					3 => "весенне-осенняя уставка",
					_ => "Группа " + groupIndex
				};
				return "— " + name + " —";
			};
		}
		return (int groupIndex) => "— Группа " + groupIndex + " —";
	}

	private static bool IsVologdaArkhangelskSection(string inputPath)
	{
		string name = Path.GetFileNameWithoutExtension(inputPath ?? "");
		string normalized = Regex.Replace((name ?? "").ToLowerInvariant().Replace('ё', 'е'), "[^a-zа-я0-9]+", "");
		// «Вологда - Архангельск», «Вологда- Архангельск ВР»
		return normalized.Contains("вологда") && normalized.Contains("архангельск");
	}

	private static string NormalizeTnvKey(string tnv)
	{
		return (tnv ?? "").Replace("_x000A_", " ").Trim();
	}

	private static TNV FinalizeSingleTnvPaNumbering(TNV block)
	{
		RenumberPaListsInPlace(block.MdpPa, block.MdpPaCriteria, startNumber: 1);
		return block;
	}

	private static TNV MergeTnvTemperatureGroup(List<TNV> rows, Func<int, string> groupLabel, bool factorSharedPa)
	{
		TNV baseRow = rows[0];
		TNV merged = new TNV
		{
			Tnv = baseRow.Tnv,
			MdpNoPA = new List<MDP>(baseRow.MdpNoPA ?? new List<MDP>()),
			MdpNoPaCriteria = new List<MDP>(baseRow.MdpNoPaCriteria ?? new List<MDP>()),
			Adp = baseRow.Adp ?? "",
			AdpCriteria = baseRow.AdpCriteria ?? "",
			MdpNoPaDop = new List<string>(baseRow.MdpNoPaDop ?? new List<string>()),
			AdpDop = new List<string>(baseRow.AdpDop ?? new List<string>()),
			MdpPa = new List<MDP>(),
			MdpPaCriteria = new List<MDP>(),
			MdpPaDop = new List<string>()
		};
		foreach (TNV row in rows)
		{
			if (CountMeaningfulMdp(row.MdpNoPA) > CountMeaningfulMdp(merged.MdpNoPA))
			{
				merged.MdpNoPA = new List<MDP>(row.MdpNoPA);
				merged.MdpNoPaCriteria = new List<MDP>(row.MdpNoPaCriteria ?? new List<MDP>());
			}
			if (string.IsNullOrWhiteSpace(merged.Adp) && !IsDashOrEmpty(row.Adp))
			{
				merged.Adp = row.Adp;
				merged.AdpCriteria = row.AdpCriteria;
			}
			foreach (string dop in row.MdpPaDop ?? new List<string>())
			{
				if (!string.IsNullOrWhiteSpace(dop) && !merged.MdpPaDop.Contains(dop))
				{
					merged.MdpPaDop.Add(dop);
				}
			}
			foreach (string dop in row.MdpNoPaDop ?? new List<string>())
			{
				if (!string.IsNullOrWhiteSpace(dop) && !merged.MdpNoPaDop.Contains(dop))
				{
					merged.MdpNoPaDop.Add(dop);
				}
			}
			foreach (string dop in row.AdpDop ?? new List<string>())
			{
				if (!string.IsNullOrWhiteSpace(dop) && !merged.AdpDop.Contains(dop))
				{
					merged.AdpDop.Add(dop);
				}
			}
		}

		List<List<PaItemPair>> groups = new List<List<PaItemPair>>();
		foreach (TNV row in rows)
		{
			List<MDP> paFormulas = StripMinimumLabels(row.MdpPa);
			List<MDP> paCriteria = StripMinimumLabels(row.MdpPaCriteria);
			List<PaItemPair> pairs = BuildPaItemPairs(paFormulas, paCriteria);
			if (pairs.Count > 0)
			{
				groups.Add(pairs);
			}
		}
		if (groups.Count == 0)
		{
			return merged;
		}
		if (factorSharedPa && groups.Count > 1)
		{
			FillFactoredPaGroups(merged, groups, groupLabel);
		}
		else
		{
			int nextNumber = 1;
			for (int g = 0; g < groups.Count; g++)
			{
				if (groups.Count > 1)
				{
					string label = groupLabel(g + 1);
					merged.MdpPa.Add(new MDP { Num = -1, Criteria = label });
					merged.MdpPaCriteria.Add(new MDP { Num = -1, Criteria = label });
				}
				nextNumber = AppendPaPairsRenumbered(merged.MdpPa, merged.MdpPaCriteria, groups[g], nextNumber);
			}
		}
		return merged;
	}

	private sealed class PaItemPair
	{
		public string Formula { get; set; } = "";

		public string Criteria { get; set; } = "";
	}

	private static List<PaItemPair> BuildPaItemPairs(List<MDP> formulas, List<MDP> criteria)
	{
		Dictionary<int, string> criteriaByNum = new Dictionary<int, string>();
		List<string> unnumberedCriteria = new List<string>();
		foreach (MDP item in criteria ?? new List<MDP>())
		{
			if (!HasMeaningfulMdpText(item.Criteria) || IsPaGroupSeparator(item.Criteria))
			{
				continue;
			}
			if (item.Num >= 0)
			{
				if (!criteriaByNum.ContainsKey(item.Num))
				{
					criteriaByNum[item.Num] = item.Criteria;
				}
			}
			else if (!item.Criteria.TrimStart().StartsWith("—", StringComparison.Ordinal))
			{
				unnumberedCriteria.Add(item.Criteria);
			}
		}
		List<PaItemPair> pairs = new List<PaItemPair>();
		foreach (MDP item in formulas ?? new List<MDP>())
		{
			if (!HasMeaningfulMdpText(item.Criteria) || IsPaGroupSeparator(item.Criteria))
			{
				continue;
			}
			if (item.Num < 0 && item.Criteria.TrimStart().StartsWith("—", StringComparison.Ordinal))
			{
				continue;
			}
			string crit = "";
			if (item.Num >= 0)
			{
				criteriaByNum.TryGetValue(item.Num, out crit);
			}
			pairs.Add(new PaItemPair
			{
				Formula = item.Criteria,
				Criteria = crit ?? ""
			});
		}
		// Критерии без формулы (редко) — не теряем
		HashSet<string> usedCrit = new HashSet<string>(pairs.Select((PaItemPair p) => NormalizePaText(p.Criteria)), StringComparer.OrdinalIgnoreCase);
		foreach (KeyValuePair<int, string> kv in criteriaByNum.OrderBy((KeyValuePair<int, string> x) => x.Key))
		{
			if (pairs.Any((PaItemPair p) => p.Criteria == kv.Value))
			{
				continue;
			}
			if (usedCrit.Contains(NormalizePaText(kv.Value)))
			{
				continue;
			}
			pairs.Add(new PaItemPair { Formula = "", Criteria = kv.Value });
		}
		foreach (string crit in unnumberedCriteria)
		{
			if (!usedCrit.Contains(NormalizePaText(crit)))
			{
				pairs.Add(new PaItemPair { Formula = "", Criteria = crit });
			}
		}
		return pairs;
	}

	private static void FillFactoredPaGroups(TNV merged, List<List<PaItemPair>> groups, Func<int, string> groupLabel)
	{
		// Независимые от уставки: одинаковая формула + одинаковый критерий во всех группах.
		List<PaItemPair> shared = new List<PaItemPair>();
		foreach (PaItemPair candidate in groups[0])
		{
			string fKey = NormalizePaText(candidate.Formula);
			string cKey = NormalizePaText(candidate.Criteria);
			if (string.IsNullOrEmpty(fKey) && string.IsNullOrEmpty(cKey))
			{
				continue;
			}
			bool inAll = groups.All((List<PaItemPair> g) => g.Any((PaItemPair p) =>
				NormalizePaText(p.Formula) == fKey && NormalizePaText(p.Criteria) == cKey));
			if (inAll)
			{
				shared.Add(candidate);
			}
		}
		HashSet<string> sharedKeys = new HashSet<string>(
			shared.Select((PaItemPair p) => PaPairKey(p)),
			StringComparer.OrdinalIgnoreCase);

		int next = 1;
		next = AppendPaPairsRenumbered(merged.MdpPa, merged.MdpPaCriteria, shared, next);
		for (int g = 0; g < groups.Count; g++)
		{
			List<PaItemPair> specific = groups[g]
				.Where((PaItemPair p) => !sharedKeys.Contains(PaPairKey(p)))
				.ToList();
			if (specific.Count == 0)
			{
				continue;
			}
			string label = groupLabel(g + 1);
			merged.MdpPa.Add(new MDP { Num = -1, Criteria = label });
			merged.MdpPaCriteria.Add(new MDP { Num = -1, Criteria = label });
			next = AppendPaPairsRenumbered(merged.MdpPa, merged.MdpPaCriteria, specific, next);
		}
	}

	private static string PaPairKey(PaItemPair p)
	{
		return NormalizePaText(p.Formula) + "\u0001" + NormalizePaText(p.Criteria);
	}

	private static string NormalizePaText(string text)
	{
		string t = (text ?? "").Replace("_x000A_", " ").Trim();
		t = Regex.Replace(t, "\\s+", " ");
		return t;
	}

	private static int AppendPaPairsRenumbered(List<MDP> targetFormulas, List<MDP> targetCriteria, List<PaItemPair> pairs, int startNumber)
	{
		int next = startNumber;
		foreach (PaItemPair pair in pairs ?? new List<PaItemPair>())
		{
			bool hasFormula = HasMeaningfulMdpText(pair.Formula);
			bool hasCriteria = HasMeaningfulMdpText(pair.Criteria);
			if (!hasFormula && !hasCriteria)
			{
				continue;
			}
			int num = next++;
			if (hasFormula)
			{
				targetFormulas.Add(new MDP { Num = num, Criteria = pair.Formula });
			}
			if (hasCriteria)
			{
				targetCriteria.Add(new MDP { Num = num, Criteria = pair.Criteria });
			}
		}
		return next;
	}

	private static List<MDP> StripMinimumLabels(List<MDP> items)
	{
		if (items == null)
		{
			return new List<MDP>();
		}
		return items
			.Where((MDP m) => m != null && !string.IsNullOrWhiteSpace(m.Criteria) && !m.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
			.ToList();
	}

	private static List<MDP> StripGroupSeparators(List<MDP> items)
	{
		return items.Where((MDP m) => !IsPaGroupSeparator(m.Criteria)).ToList();
	}

	private static bool IsPaGroupSeparator(string text)
	{
		string t = (text ?? "").Trim();
		if (!(t.StartsWith("—", StringComparison.Ordinal) || t.StartsWith("-", StringComparison.Ordinal)))
		{
			return false;
		}
		return t.IndexOf("Группа ", StringComparison.OrdinalIgnoreCase) >= 0
			|| t.IndexOf("летняя уставка", StringComparison.OrdinalIgnoreCase) >= 0
			|| t.IndexOf("зимняя уставка", StringComparison.OrdinalIgnoreCase) >= 0
			|| t.IndexOf("весенне-осенняя уставка", StringComparison.OrdinalIgnoreCase) >= 0
			|| t.IndexOf("весенне–осенняя уставка", StringComparison.OrdinalIgnoreCase) >= 0;
	}

	private static int CountMeaningfulMdp(List<MDP> items)
	{
		return items == null ? 0 : items.Count((MDP m) => HasMeaningfulMdpText(m.Criteria) && !IsPaGroupSeparator(m.Criteria));
	}

	private static int AppendRenumberedPaGroup(List<MDP> targetFormulas, List<MDP> targetCriteria, List<MDP> formulas, List<MDP> criteria, int startNumber)
	{
		Dictionary<int, int> mapping = new Dictionary<int, int>();
		int next = startNumber;
		foreach (MDP item in formulas ?? new List<MDP>())
		{
			if (!HasMeaningfulMdpText(item.Criteria))
			{
				continue;
			}
			if (item.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}
			if (IsPaGroupSeparator(item.Criteria) || (item.Num < 0 && item.Criteria.TrimStart().StartsWith("—", StringComparison.Ordinal)))
			{
				targetFormulas.Add(new MDP { Num = -1, Criteria = item.Criteria });
				continue;
			}
			int newNum = next++;
			if (item.Num >= 0)
			{
				mapping[item.Num] = newNum;
			}
			targetFormulas.Add(new MDP
			{
				Num = newNum,
				Criteria = item.Criteria
			});
		}
		foreach (MDP item in criteria ?? new List<MDP>())
		{
			if (!HasMeaningfulMdpText(item.Criteria))
			{
				continue;
			}
			if (item.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}
			if (IsPaGroupSeparator(item.Criteria) || (item.Num < 0 && item.Criteria.TrimStart().StartsWith("—", StringComparison.Ordinal)))
			{
				targetCriteria.Add(new MDP { Num = -1, Criteria = item.Criteria });
				continue;
			}
			int newNum;
			if (item.Num >= 0 && mapping.TryGetValue(item.Num, out newNum))
			{
				targetCriteria.Add(new MDP { Num = newNum, Criteria = item.Criteria });
			}
			else if (item.Num >= 0)
			{
				newNum = next++;
				mapping[item.Num] = newNum;
				targetCriteria.Add(new MDP { Num = newNum, Criteria = item.Criteria });
			}
			else
			{
				targetCriteria.Add(new MDP { Num = -1, Criteria = item.Criteria });
			}
		}
		return next;
	}

	private static void RenumberPaListsInPlace(List<MDP> formulas, List<MDP> criteria, int startNumber)
	{
		if (formulas == null)
		{
			return;
		}
		List<MDP> tmpFormulas = new List<MDP>();
		List<MDP> tmpCriteria = new List<MDP>();
		AppendRenumberedPaGroup(tmpFormulas, tmpCriteria, formulas, criteria ?? new List<MDP>(), startNumber);
		formulas.Clear();
		formulas.AddRange(tmpFormulas);
		if (criteria != null)
		{
			criteria.Clear();
			criteria.AddRange(tmpCriteria);
		}
	}

	private static TNV ReadTnvBlock(ExcelOperations ex, ColumnMap columnMap, int bRow, int eRow, string rowLabel)
	{
		List<PaColumn> paValueCols = BuildPaColumnList(columnMap.MdpPaCol, columnMap.ExtraPaValueCols);
		List<PaColumn> paCriteriaCols = BuildPaColumnList(columnMap.MdpPaCriteriaCol, columnMap.ExtraPaCriteriaCols);
		List<PaColumn> paDopCols = BuildPaColumnList(columnMap.MdpPaDopCol, columnMap.ExtraPaDopCols);
		TNV tnv = new TNV
		{
			Tnv = rowLabel ?? "",
			MdpNoPA = (columnMap.MdpNoPaCol != -1) ? ReadLines(ex, bRow, eRow, columnMap.MdpNoPaCol, modify: true) : new List<MDP>(),
			MdpPa = MergePaValueColumns(ex, bRow, eRow, paValueCols, modify: true),
			Adp = CellModifyString(ReadLine(ex, bRow, eRow, columnMap.AdpCol)),
			MdpNoPaCriteria = (columnMap.MdpNoPaCriteriaCol != -1) ? ReadLines(ex, bRow, eRow, columnMap.MdpNoPaCriteriaCol) : new List<MDP>(),
			MdpPaCriteria = MergePaValueColumns(ex, bRow, eRow, paCriteriaCols, modify: false),
			AdpCriteria = CellModifyString(ReadLine(ex, bRow, eRow, columnMap.AdpCriteriaCol)),
			MdpNoPaDop = (columnMap.MdpNoPaDopCol != -1) ? ReadDopLines(ex, bRow, eRow, columnMap.MdpNoPaDopCol) : new List<string>(),
			MdpPaDop = MergePaDopColumns(ex, bRow, eRow, paDopCols),
			AdpDop = (columnMap.AdpDopCol != -1) ? ReadDopLines(ex, bRow, eRow, columnMap.AdpDopCol) : new List<string>()
		};
		return NormalizeNotControlledBlock(tnv);
	}

	/// <summary>
	/// Широкий merge «Не контролируется» (как в Вологде) через getStrMerged
	/// попадает во все колонки, включая контроль. Нормализуем в одну метку ТНВ.
	/// </summary>
	private static TNV NormalizeNotControlledBlock(TNV tnv)
	{
		if (tnv == null)
		{
			return tnv;
		}
		tnv.MdpNoPaDop = FilterDopLines(tnv.MdpNoPaDop);
		tnv.MdpPaDop = FilterDopLines(tnv.MdpPaDop);
		tnv.AdpDop = FilterDopLines(tnv.AdpDop);
		if (IsNotControlledPhrase(tnv.Tnv) || BlockContentIsOnlyNotControlled(tnv))
		{
			tnv.Tnv = "Не контролируется";
			tnv.MdpNoPA = new List<MDP>();
			tnv.MdpPa = new List<MDP>();
			tnv.MdpNoPaCriteria = new List<MDP>();
			tnv.MdpPaCriteria = new List<MDP>();
			tnv.Adp = "";
			tnv.AdpCriteria = "";
			tnv.MdpNoPaDop = new List<string>();
			tnv.MdpPaDop = new List<string>();
			tnv.AdpDop = new List<string>();
		}
		return tnv;
	}

	private static bool BlockContentIsOnlyNotControlled(TNV tnv)
	{
		List<string> texts = new List<string>();
		void Add(string s)
		{
			if (!string.IsNullOrWhiteSpace(s))
			{
				texts.Add(s.Trim());
			}
		}
		foreach (MDP m in (tnv.MdpNoPA ?? new List<MDP>())
			.Concat(tnv.MdpPa ?? new List<MDP>())
			.Concat(tnv.MdpNoPaCriteria ?? new List<MDP>())
			.Concat(tnv.MdpPaCriteria ?? new List<MDP>()))
		{
			Add(m.Criteria);
		}
		Add(tnv.Adp);
		Add(tnv.AdpCriteria);
		foreach (string d in (tnv.MdpNoPaDop ?? new List<string>())
			.Concat(tnv.MdpPaDop ?? new List<string>())
			.Concat(tnv.AdpDop ?? new List<string>()))
		{
			Add(d);
		}
		if (texts.Count == 0)
		{
			return false;
		}
		bool sawNotControlled = texts.Any((string t) =>
			IsNotControlledPhrase(t) || t.IndexOf("Не контролируется", StringComparison.OrdinalIgnoreCase) >= 0);
		if (!sawNotControlled)
		{
			return false;
		}
		return !texts.Any(HasRealMdpOrControlContent);
	}

	private static bool HasRealMdpOrControlContent(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		if (IsNotControlledPhrase(text.Trim()))
		{
			return false;
		}
		if (IsSectionLabelOnly(text))
		{
			return false;
		}
		return HasMeaningfulDopText(text) || (HasMeaningfulMdpText(text) && !ContainsOnlyNotControlledNoise(text));
	}

	private static bool IsSectionLabelOnly(string text)
	{
		string t = (text ?? "").Trim();
		return t.Length >= 2 && t.StartsWith("—", StringComparison.Ordinal) && t.EndsWith("—", StringComparison.Ordinal);
	}

	private static bool ContainsOnlyNotControlledNoise(string text)
	{
		return !HasMeaningfulDopText(text);
	}

	private static List<string> FilterDopLines(List<string> lines)
	{
		if (lines == null || lines.Count == 0)
		{
			return new List<string>();
		}
		return lines.Where(HasMeaningfulDopText).ToList();
	}

	private static bool HasMeaningfulDopText(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string cleaned = (text ?? "").Replace("_x000A_", "\n");
		cleaned = Regex.Replace(cleaned, @"—\s*[^—\r\n]+?\s*—", " ");
		cleaned = Regex.Replace(cleaned, @"Не контролируется", " ", RegexOptions.IgnoreCase);
		cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();
		if (string.IsNullOrWhiteSpace(cleaned) || IsDashOrEmpty(cleaned))
		{
			return false;
		}
		return true;
	}

	/// <summary>
	/// Если в критериях МДП/МДП с ПА/АДП есть «ДДТН …», в соответствующем
	/// контроле доп. параметров должна быть приписка (ДДТН/ДТН того же объекта).
	/// При отсутствии — дописываем «ДДТН …» и пишем предупреждение в консоль.
	/// </summary>
	private static void EnsureDdtnControlAnnotations(List<TNV> tnvList, string schemeNum, string schemeName)
	{
		if (tnvList == null)
		{
			return;
		}
		string schemeLabel = ((schemeNum ?? "").Trim() + " " + (schemeName ?? "").Trim()).Trim();
		foreach (TNV tnv in tnvList)
		{
			if (tnv == null || IsNotControlledPhrase(tnv.Tnv))
			{
				continue;
			}
			tnv.MdpNoPaDop = tnv.MdpNoPaDop ?? new List<string>();
			tnv.MdpPaDop = tnv.MdpPaDop ?? new List<string>();
			tnv.AdpDop = tnv.AdpDop ?? new List<string>();
			EnsureDdtnControlForCriteria(
				(tnv.MdpNoPaCriteria ?? new List<MDP>()).Select((MDP m) => m.Criteria),
				tnv.MdpNoPaDop,
				schemeLabel,
				tnv.Tnv,
				"МДП без ПА");
			EnsureDdtnControlForCriteria(
				(tnv.MdpPaCriteria ?? new List<MDP>()).Select((MDP m) => m.Criteria),
				tnv.MdpPaDop,
				schemeLabel,
				tnv.Tnv,
				"МДП с ПА");
			EnsureDdtnControlForCriteria(
				new string[1] { tnv.AdpCriteria },
				tnv.AdpDop,
				schemeLabel,
				tnv.Tnv,
				"АДП");
		}
	}

	private static void EnsureDdtnControlForCriteria(
		IEnumerable<string> criteriaTexts,
		List<string> controlLines,
		string schemeLabel,
		string tnvLabel,
		string channelName)
	{
		List<string> objects = new List<string>();
		foreach (string text in criteriaTexts ?? Array.Empty<string>())
		{
			foreach (string obj in ExtractDdtnObjects(text))
			{
				if (!objects.Any((string x) => string.Equals(NormalizeThermalObjectKey(x), NormalizeThermalObjectKey(obj), StringComparison.Ordinal)))
				{
					objects.Add(obj);
				}
			}
		}
		foreach (string obj in objects)
		{
			if (ControlCoversThermalObject(controlLines, obj))
			{
				continue;
			}
			string annotation = "ДДТН " + obj;
			controlLines.Add(annotation);
			string tnvPart = string.IsNullOrWhiteSpace(tnvLabel) ? "" : (", ТНВ/АРПМ «" + tnvLabel.Trim() + "»");
			Console.WriteLine(
				"Предупреждение: схема «" + schemeLabel + "»" + tnvPart
				+ ": в критерии «" + channelName + "» есть ДДТН «" + obj
				+ "», в контроле доп. параметров не было приписки — добавлено «" + annotation + "».");
		}
	}

	private static List<string> ExtractDdtnObjects(string text)
	{
		List<string> list = new List<string>();
		if (string.IsNullOrWhiteSpace(text))
		{
			return list;
		}
		string normalized = text.Replace("_x000A_", "\n");
		foreach (Match match in Regex.Matches(normalized, @"(?i)ДДТН\s+([^\r\n]+)"))
		{
			string obj = match.Groups[1].Value.Trim().TrimEnd('.', ';', ',');
			obj = Regex.Replace(obj, @"\s+", " ").Trim();
			if (!string.IsNullOrWhiteSpace(obj) && !IsSectionLabelOnly(obj) && !IsNotControlledPhrase(obj))
			{
				list.Add(obj);
			}
		}
		return list;
	}

	private static bool ControlCoversThermalObject(IEnumerable<string> controlLines, string obj)
	{
		string want = NormalizeThermalObjectKey(obj);
		if (string.IsNullOrEmpty(want))
		{
			return false;
		}
		foreach (string line in controlLines ?? Array.Empty<string>())
		{
			if (string.IsNullOrWhiteSpace(line))
			{
				continue;
			}
			string have = NormalizeThermalObjectKey(StripThermalControlPrefix(line));
			if (string.IsNullOrEmpty(have))
			{
				have = NormalizeThermalObjectKey(line);
			}
			if (have.Contains(want) || want.Contains(have))
			{
				return true;
			}
		}
		return false;
	}

	private static string StripThermalControlPrefix(string text)
	{
		string t = (text ?? "").Replace("_x000A_", "\n").Trim();
		t = Regex.Replace(t, @"^\s*\d+\)\s*", "");
		t = Regex.Replace(t, @"^(?i)(?:ДДТН|ДТН)\s+", "");
		return t.Trim();
	}

	private static string NormalizeThermalObjectKey(string text)
	{
		string t = NormalizePaText(text ?? "").ToLowerInvariant();
		t = Regex.Replace(t, @"^(?:\d+\)\s*)?(?:ддтн|дтн)\s+", "", RegexOptions.IgnoreCase);
		t = t.Replace("–", "-").Replace("—", "-").Replace("−", "-");
		t = Regex.Replace(t, @"\s+", "");
		return t;
	}

	private static List<PaColumn> BuildPaColumnList(int primaryCol, List<PaColumn> extras)
	{
		List<PaColumn> list = new List<PaColumn>();
		if (primaryCol != -1)
		{
			list.Add(new PaColumn { Col = primaryCol, Title = "" });
		}
		if (extras != null)
		{
			list.AddRange(extras);
		}
		return list;
	}

	private static List<MDP> MergePaValueColumns(ExcelOperations ex, int bRow, int eRow, List<PaColumn> cols, bool modify)
	{
		List<MDP> target = new List<MDP>();
		if (cols == null || cols.Count == 0)
		{
			return target;
		}
		foreach (PaColumn pa in cols)
		{
			List<MDP> meaningful = ReadLines(ex, bRow, eRow, pa.Col, modify)
				.Where((MDP m) => !string.IsNullOrWhiteSpace(m.Criteria))
				.ToList();
			if (meaningful.Count == 0)
			{
				continue;
			}
			if (target.Count > 0 && !string.IsNullOrWhiteSpace(pa.Title))
			{
				target.Add(new MDP
				{
					Num = -1,
					Criteria = "— " + pa.Title + " —"
				});
			}
			target.AddRange(meaningful);
		}
		return target;
	}

	private static List<string> MergePaDopColumns(ExcelOperations ex, int bRow, int eRow, List<PaColumn> cols)
	{
		List<string> target = new List<string>();
		if (cols == null || cols.Count == 0)
		{
			return target;
		}
		foreach (PaColumn pa in cols)
		{
			List<string> lines = ReadDopLines(ex, bRow, eRow, pa.Col).Where(HasMeaningfulDopText).ToList();
			if (lines.Count == 0)
			{
				continue;
			}
			if (target.Count > 0 && !string.IsNullOrWhiteSpace(pa.Title))
			{
				target.Add("— " + pa.Title + " —");
			}
			target.AddRange(lines);
		}
		return target;
	}

	public static string ReadLine(ExcelOperations ex, int bRow, int eRow, int col)
	{
		if (col <= 0)
		{
			return "";
		}
		string result = "";
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStrMerged(i, col);
			if (text != "" && text != " ")
			{
				result = text.Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			}
		}
		return result;
	}

	public static List<MDP> ReadLines(ExcelOperations ex, int bRow, int eRow, int col, bool modify = false)
	{
		List<MDP> list = new List<MDP>();
		if (col <= 0)
		{
			return list;
		}
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStrMerged(i, col).Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			if (text == "" || text == " ")
			{
				continue;
			}
			if (text.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
			{
				list.Add(new MDP
				{
					Num = -1,
					Criteria = text
				});
				continue;
			}
			Match match = Regex.Match(text, "^(-?\\d+)\\)\\s*(.*)$", RegexOptions.Singleline);
			if (match.Success)
			{
				int num = Convert.ToInt32(match.Groups[1].Value);
				string text2 = match.Groups[2].Value;
				string criteria = ReorderNumberedBlocks(modify ? CellModifyString(text2) : text2);
				if (!HasMeaningfulMdpText(criteria))
				{
					continue;
				}
				list.Add(new MDP
				{
					Num = num,
					Criteria = criteria
				});
			}
			else
			{
				string criteria = ReorderNumberedBlocks(modify ? CellModifyString(text) : text);
				if (!HasMeaningfulMdpText(criteria))
				{
					continue;
				}
				list.Add(new MDP
				{
					Num = -1,
					Criteria = criteria
				});
			}
		}
		return list;
	}

	public static List<string> ReadDopLines(ExcelOperations ex, int bRow, int eRow, int col)
	{
		List<string> list = new List<string>();
		if (col <= 0)
		{
			return list;
		}
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStrMerged(i, col).Trim(new char[1] { ' ' }).Replace("_x000A_", Environment.NewLine);
			if (text != "" && text != " ")
			{
				list.Add(text);
			}
		}
		return list;
	}

	public static string CellModifyString(string text)
	{
		text = StripLeadingItemNumber((text ?? "").Replace("_x000A_", Environment.NewLine).Trim());
		if (string.IsNullOrWhiteSpace(text))
		{
			return "";
		}
		if (TryFormatNestedIfCriteria(text, out var formattedIf))
		{
			return formattedIf;
		}
		if (TryFormatCaseCriteria(text, out var formattedCase))
		{
			return formattedCase;
		}
		text = Regex.Replace(text, "\\bMIN\\b", "min");
		text = text.Replace("==", "=");
		if (!AreBracketsBalanced(text))
		{
			return text;
		}
		BracketNode node = Parse(text);
		return Reconstruct(node) ?? "";
	}

	private static string StripLeadingItemNumber(string text)
	{
		Match match = Regex.Match(text ?? "", "^(-?\\d+)\\)\\s*(.*)$", RegexOptions.Singleline);
		return match.Success ? match.Groups[2].Value.Trim() : (text ?? "");
	}

	private static bool TryFormatCaseCriteria(string text, out string formatted)
	{
		formatted = "";
		string raw = TrimOuterWrapping(text).Trim();
		if (!raw.StartsWith("case(", StringComparison.OrdinalIgnoreCase) || !raw.EndsWith(")"))
		{
			return false;
		}
		string inside = raw.Substring(5, raw.Length - 6);
		List<string> args = SplitTopLevelArgs(inside);
		if (args.Count < 3)
		{
			return false;
		}
		string variable = args[0].Trim();
		StringBuilder sb = new StringBuilder();
		int i = 1;
		int written = 0;
		while (i + 1 < args.Count)
		{
			string condRaw = args[i].Trim();
			string valueRaw = args[i + 1].Trim();
			i += 2;
			if (!TryParseIsCall(condRaw, out string isValue))
			{
				continue;
			}
			string value = UnwrapReturn(valueRaw);
			if (written > 0)
			{
				sb.Append(',');
				sb.Append(Environment.NewLine);
			}
			sb.Append(variable);
			sb.Append('=');
			sb.Append(isValue);
			sb.Append(':');
			sb.Append(Environment.NewLine);
			sb.Append(NormalizeMathExpression(value));
			written++;
		}
		if (written == 0)
		{
			return false;
		}
		formatted = sb.ToString();
		return true;
	}

	private static List<string> SplitTopLevelArgs(string text)
	{
		List<string> list = new List<string>();
		int depth = 0;
		int start = 0;
		for (int i = 0; i < text.Length; i++)
		{
			char c = text[i];
			if (c == '(')
			{
				depth++;
			}
			else if (c == ')')
			{
				depth--;
			}
			else if (c == ',' && depth == 0)
			{
				list.Add(text.Substring(start, i - start).Trim());
				start = i + 1;
			}
		}
		if (start <= text.Length)
		{
			list.Add(text.Substring(start).Trim());
		}
		return list;
	}

	private static bool TryParseIsCall(string text, out string value)
	{
		value = "";
		string raw = TrimOuterWrapping(text).Trim();
		Match match = Regex.Match(raw, "^is\\((.*)\\)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
		if (!match.Success)
		{
			return false;
		}
		value = match.Groups[1].Value.Trim();
		return !string.IsNullOrWhiteSpace(value);
	}

	private static string UnwrapReturn(string text)
	{
		string raw = TrimOuterWrapping(text).Trim();
		Match match = Regex.Match(raw, "^return\\s*\\((.*)\\)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
		if (match.Success)
		{
			return match.Groups[1].Value.Trim();
		}
		match = Regex.Match(raw, "^return\\s+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
		return match.Success ? match.Groups[1].Value.Trim() : raw;
	}

	private sealed class IfExprNode
	{
		public string Condition { get; set; } = "";

		public IfExprNode TrueBranch { get; set; }

		public IfExprNode FalseBranch { get; set; }

		public string LeafExpression { get; set; }

		public bool IsLeaf => LeafExpression != null;
	}

	private sealed class IfLeafCase
	{
		public List<string> Conditions { get; } = new List<string>();

		public string Expression { get; set; } = "";
	}

	private static bool TryFormatNestedIfCriteria(string text, out string formatted)
	{
		formatted = "";
		if (!TryParseIfNode(text, out var root))
		{
			return false;
		}
		List<IfLeafCase> list = new List<IfLeafCase>();
		CollectIfLeafCases(root, new List<string>(), list);
		if (list.Count <= 1)
		{
			return false;
		}
		StringBuilder stringBuilder = new StringBuilder();
		for (int i = 0; i < list.Count; i++)
		{
			IfLeafCase ifLeafCase = list[i];
			string text2 = string.Join(" И ", ifLeafCase.Conditions.Select((string c) => NormalizeCondition(c)).Where((string c) => !string.IsNullOrWhiteSpace(c)));
			if (string.IsNullOrWhiteSpace(text2))
			{
				text2 = "Условие";
			}
			stringBuilder.Append(text2);
			stringBuilder.Append(':');
			stringBuilder.Append(Environment.NewLine);
			stringBuilder.Append(NormalizeMathExpression(ifLeafCase.Expression));
			if (i != list.Count - 1)
			{
				stringBuilder.Append(',');
			}
			if (i != list.Count - 1)
			{
				stringBuilder.Append(Environment.NewLine);
			}
		}
		formatted = stringBuilder.ToString();
		return !string.IsNullOrWhiteSpace(formatted);
	}

	private static void CollectIfLeafCases(IfExprNode node, List<string> currentConditions, List<IfLeafCase> result)
	{
		if (node.IsLeaf)
		{
			result.Add(new IfLeafCase
			{
				Expression = node.LeafExpression ?? "",
			});
			result[result.Count - 1].Conditions.AddRange(currentConditions);
			return;
		}
		List<string> list = new List<string>(currentConditions);
		list.Add(node.Condition);
		if (node.TrueBranch != null)
		{
			CollectIfLeafCases(node.TrueBranch, list, result);
		}
		List<string> list2 = new List<string>(currentConditions);
		list2.Add(NegateCondition(node.Condition));
		if (node.FalseBranch != null)
		{
			CollectIfLeafCases(node.FalseBranch, list2, result);
		}
	}

	private static bool TryParseIfNode(string text, out IfExprNode node)
	{
		node = new IfExprNode();
		string text2 = TrimOuterWrapping(text);
		if (!IsIfCall(text2))
		{
			node.LeafExpression = text2;
			return false;
		}
		if (!TrySplitIfArguments(text2, out var condition, out var trueExpr, out var falseExpr))
		{
			return false;
		}
		node.Condition = condition.Trim();
		node.TrueBranch = IsIfCall(TrimOuterWrapping(trueExpr))
			? ParseIfBranch(trueExpr)
			: new IfExprNode { LeafExpression = trueExpr.Trim() };
		node.FalseBranch = IsIfCall(TrimOuterWrapping(falseExpr))
			? ParseIfBranch(falseExpr)
			: new IfExprNode { LeafExpression = falseExpr.Trim() };
		return true;
	}

	private static IfExprNode ParseIfBranch(string text)
	{
		if (TryParseIfNode(text, out var node))
		{
			return node;
		}
		return new IfExprNode
		{
			LeafExpression = TrimOuterWrapping(text).Trim()
		};
	}

	private static bool IsIfCall(string text)
	{
		string text2 = TrimOuterWrapping(text).Trim();
		return text2.StartsWith("if(", StringComparison.OrdinalIgnoreCase);
	}

	private static bool TrySplitIfArguments(string ifCall, out string condition, out string trueExpr, out string falseExpr)
	{
		condition = "";
		trueExpr = "";
		falseExpr = "";
		string text = TrimOuterWrapping(ifCall).Trim();
		if (!text.StartsWith("if(", StringComparison.OrdinalIgnoreCase) || !text.EndsWith(")"))
		{
			return false;
		}
		string text2 = text.Substring(3, text.Length - 4);
		List<int> list = new List<int>();
		int num = 0;
		for (int i = 0; i < text2.Length; i++)
		{
			char c = text2[i];
			if (c == '(')
			{
				num++;
			}
			else if (c == ')')
			{
				num--;
			}
			else if (c == ',' && num == 0)
			{
				list.Add(i);
			}
		}
		if (list.Count != 2)
		{
			return false;
		}
		condition = text2.Substring(0, list[0]).Trim();
		trueExpr = text2.Substring(list[0] + 1, list[1] - list[0] - 1).Trim();
		falseExpr = text2.Substring(list[1] + 1).Trim();
		return condition.Length > 0 && trueExpr.Length > 0 && falseExpr.Length > 0;
	}

	private static string TrimOuterWrapping(string input)
	{
		string text = (input ?? "").Trim();
		while (text.StartsWith("(") && text.EndsWith(")") && IsOuterPair(text))
		{
			text = text.Substring(1, text.Length - 2).Trim();
		}
		return text;
	}

	private static bool IsOuterPair(string text)
	{
		int num = 0;
		for (int i = 0; i < text.Length; i++)
		{
			if (text[i] == '(')
			{
				num++;
			}
			else if (text[i] == ')')
			{
				num--;
				if (num == 0 && i < text.Length - 1)
				{
					return false;
				}
			}
		}
		return num == 0;
	}

	private static string NormalizeCondition(string condition)
	{
		string text = Regex.Replace((condition ?? "").Replace("==", "=").Trim(), "\\s+", " ");
		text = Regex.Replace(text, "(>=|<=|<>|=|>|<)(-?\\d+(?:\\.\\d+)?)([A-Za-zА-Яа-я_])", "$1$2 И $3");
		text = Regex.Replace(text, "(\\d)\\s*(?=[A-Za-zА-Яа-я_][A-Za-zА-Яа-я0-9_]*\\s*(?:=|<>|>=|<=|>|<))", "$1 И ");
		text = Regex.Replace(text, "(?<=[A-Za-zА-Яа-я0-9_\\)])И(?=[A-Za-zА-Яа-я0-9_\\(])", " И ");
		return Regex.Replace(text, "\\s+", " ").Trim();
	}

	private static string NormalizeMathExpression(string expression)
	{
		return Regex.Replace((expression ?? "").Replace("==", "=").Trim(), "\\s+", " ");
	}

	private static string NegateCondition(string condition)
	{
		string text = NormalizeCondition(condition);
		Match match = Regex.Match(text, "^(?<left>.+?)\\s*(?<op>>=|<=|>|<|=|<>)\\s*(?<right>-?\\d+(?:\\.\\d+)?)$");
		if (!match.Success)
		{
			return "НЕ(" + text + ")";
		}
		string value = match.Groups["left"].Value.Trim();
		string value2 = match.Groups["op"].Value.Trim();
		string value3 = match.Groups["right"].Value.Trim();
		if (int.TryParse(value3, out var result))
		{
			switch (value2)
			{
			case ">=":
				return $"{value}<={result - 1}";
			case "<=":
				return $"{value}>={result + 1}";
			case ">":
				return $"{value}<={result}";
			case "<":
				return $"{value}>={result}";
			}
		}
		switch (value2)
		{
		case ">=":
			return $"{value}<{value3}";
		case "<=":
			return $"{value}>{value3}";
		case ">":
			return $"{value}<={value3}";
		case "<":
			return $"{value}>={value3}";
		case "=":
			if (int.TryParse(value3, out var result2) && (result2 == 0 || result2 == 1))
			{
				return $"{value}={1 - result2}";
			}
			return $"{value}<>{value3}";
		case "<>":
			if (int.TryParse(value3, out var result3) && (result3 == 0 || result3 == 1))
			{
				return $"{value}={value3}";
			}
			return $"{value}={value3}";
		default:
			return "НЕ(" + text + ")";
		}
	}

	private static string ReorderNumberedBlocks(string text)
	{
		string text2 = (text ?? "").Replace("_x000A_", Environment.NewLine);
		MatchCollection matchCollection = Regex.Matches(text2, "(?m)^\\s*(\\d+)\\)\\s");
		if (matchCollection.Count <= 1)
		{
			return text2;
		}
		int index = matchCollection[0].Index;
		string text3 = text2.Substring(0, index);
		List<(int num, string block)> list = new List<(int, string)>();
		for (int i = 0; i < matchCollection.Count; i++)
		{
			int num2 = matchCollection[i].Index;
			int num3 = ((i == matchCollection.Count - 1) ? text2.Length : matchCollection[i + 1].Index);
			if (!int.TryParse(matchCollection[i].Groups[1].Value, out var result))
			{
				result = int.MaxValue;
			}
			string item = text2.Substring(num2, num3 - num2).TrimEnd('\r', '\n');
			list.Add((result, item));
		}
		list = list.OrderBy((ValueTuple<int, string> x) => x.Item1).ToList();
		string text4 = string.Join(Environment.NewLine, list.Select((ValueTuple<int, string> x) => x.Item2));
		if (string.IsNullOrWhiteSpace(text3))
		{
			return text4;
		}
		string text5 = text3.TrimEnd('\r', '\n');
		return text5 + Environment.NewLine + text4;
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
		if (tnvList == null || tnvList.Count == 0)
		{
			return "";
		}
		List<string> values = new List<string>();
		foreach (TNV tnv in tnvList)
		{
			foreach (string part in tnv.AdpDop ?? new List<string>())
			{
				string text = (part ?? "").Replace("_x000A_", " ").Trim();
				if (!string.IsNullOrWhiteSpace(text))
				{
					values.Add(text);
				}
			}
		}
		if (values.Count == 0)
		{
			return "";
		}
		List<string> unique = values
			.Select(NormalizePaText)
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.ToList();
		if (unique.Count == 1)
		{
			// Вернуть исходный (с нормальными пробелами) первый экземпляр
			return values.First((string v) => string.Equals(NormalizePaText(v), unique[0], StringComparison.OrdinalIgnoreCase));
		}
		return "";
	}

	private static bool IsNotControlledPhrase(string text)
	{
		return string.Equals((text ?? "").Trim(), "Не контролируется", StringComparison.OrdinalIgnoreCase);
	}

	private static readonly Color[] CriteriaPalette = new Color[12]
	{
		Color.FromArgb(0, 70, 140),
		Color.FromArgb(140, 60, 0),
		Color.FromArgb(0, 110, 40),
		Color.FromArgb(120, 0, 120),
		Color.FromArgb(160, 0, 0),
		Color.FromArgb(0, 110, 110),
		Color.FromArgb(90, 60, 0),
		Color.FromArgb(0, 0, 150),
		Color.FromArgb(150, 0, 70),
		Color.FromArgb(70, 90, 0),
		Color.FromArgb(0, 80, 160),
		Color.FromArgb(110, 40, 40)
	};

	private static Dictionary<string, Color> BuildCriteriaColorMap(List<MdpBuilder> schemes)
	{
		Dictionary<string, Color> dictionary = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
		int num = 0;
		foreach (MdpBuilder scheme in schemes)
		{
			foreach (TNV tnv in scheme.TnvList)
			{
				foreach (MDP item in tnv.MdpNoPaCriteria.Concat(tnv.MdpPaCriteria))
				{
					string text = NormalizeCriteriaKey(item.Criteria);
					if (!string.IsNullOrWhiteSpace(text) && !dictionary.ContainsKey(text))
					{
						dictionary[text] = CriteriaPalette[num % CriteriaPalette.Length];
						num++;
					}
				}
			}
		}
		return dictionary;
	}

	private static string NormalizeCriteriaKey(string text)
	{
		return Regex.Replace((text ?? "").Replace("_x000A_", " ").Trim(), "\\s+", " ");
	}

	private static Color GetColorForCriterion(Dictionary<string, Color> colorMap, string criteriaText, int fallbackNum)
	{
		string text = NormalizeCriteriaKey(criteriaText);
		if (!string.IsNullOrWhiteSpace(text) && colorMap.TryGetValue(text, out var value))
		{
			return value;
		}
		if (fallbackNum > 0)
		{
			return CriteriaPalette[(fallbackNum - 1) % CriteriaPalette.Length];
		}
		return Color.Black;
	}

	private static Color GetColorForMdpNum(Dictionary<string, Color> colorMap, List<MDP> criteriaList, int mdpNum)
	{
		MDP mDP = criteriaList.FirstOrDefault((MDP c) => c.Num == mdpNum);
		if (mDP != null)
		{
			return GetColorForCriterion(colorMap, mDP.Criteria, mdpNum);
		}
		return GetColorForCriterion(colorMap, "", mdpNum);
	}

	private static void WriteColoredMdpBlocks(ExcelOperations excelOperations, int row, int col, List<MDP> mdpBlocks, List<MDP> criteriaList, Dictionary<string, Color> colorMap)
	{
		excelOperations.ClearCell(row, col);
		excelOperations.Format(row, col, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
		for (int i = 0; i < mdpBlocks.Count; i++)
		{
			MDP mDP = mdpBlocks[i];
			bool flag = i == mdpBlocks.Count - 1;
			string text = (!flag) ? (mDP.Criteria + Environment.NewLine) : mDP.Criteria;
			string prefix = (mDP.Num != -1) ? $"{mDP.Num}) " : "";
			Color color = Color.Black;
			if (mDP.Criteria.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
			{
				color = Color.Black;
			}
			else if (mDP.Num >= 0)
			{
				color = GetColorForMdpNum(colorMap, criteriaList, mDP.Num);
			}
			excelOperations.CellRichText(row, col, text, prefix, color);
		}
	}

	private static string WriteColoredCriteriaBlocks(ExcelOperations excelOperations, int row, int col, List<MDP> criteriaList, Dictionary<string, Color> colorMap)
	{
		excelOperations.ClearCell(row, col);
		excelOperations.Format(row, col, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
		StringBuilder stringBuilder = new StringBuilder();
		for (int i = 0; i < criteriaList.Count; i++)
		{
			MDP mDP = criteriaList[i];
			bool flag = i == criteriaList.Count - 1;
			string text = (mDP.Num != -1) ? $"{mDP.Num}) {mDP.Criteria}" : mDP.Criteria;
			if (!flag)
			{
				text += Environment.NewLine;
			}
			Color color = GetColorForCriterion(colorMap, mDP.Criteria, mDP.Num);
			excelOperations.AppendColoredText(row, col, text, color);
			stringBuilder.Append(text);
		}
		return stringBuilder.ToString().TrimEnd('\r', '\n');
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
