using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using Xls_prjt;

namespace XlsxMdpParser;

internal class Program
{
	private static void Main(string[] args)
	{
		Console.InputEncoding = Encoding.Default;
		Console.OutputEncoding = Encoding.Default;
		Console.WriteLine("Перетащите файл в это окно и нажмите Enter:");
		string text = Console.ReadLine();
		if (!string.IsNullOrWhiteSpace(text))
		{
			string text2 = text.Trim(new char[1] { '"' });
			Console.WriteLine("Получен путь: " + text2);
			ExcelOperations excelOperations = new ExcelOperations(text2, 1);
			List<MdpBuilder> list = new List<MdpBuilder>();
			for (int i = 4; i <= excelOperations.LastColumnRow(); i++)
			{
				if (!(excelOperations.getStr(i, 3) != "") || !(excelOperations.getStr(i, 3) != " "))
				{
					continue;
				}
				string str = excelOperations.getStr(i, 3);
				string str2 = excelOperations.getStr(i, 2);
				string text3 = excelOperations.MergedCells(i, 3);
				int num = Convert.ToInt32(text3.Split(new char[1] { ':' })[0].Substring(1));
				int num2 = Convert.ToInt32(text3.Split(new char[1] { ':' })[1].Substring(1));
				List<TNV> list2 = new List<TNV>();
				for (int j = num; j <= num2; j++)
				{
					bool flag = excelOperations.getStr(j, 4) != "";
					int bRow = j;
					for (; j < num2 && (!flag || excelOperations.getStr(j + 1, 4) == ""); j++)
					{
						if (excelOperations.getStr(j + 1, 6).Trim(new char[1] { ' ' }).Equals("Минимальное из:"))
						{
							break;
						}
					}
					list2.Add(new TNV
					{
						Tnv = ReadLine(excelOperations, bRow, j, 4),
						MdpNoPA = ReadLines(excelOperations, bRow, j, 5, modify: true),
						MdpPa = ReadLines(excelOperations, bRow, j, 6, modify: true),
						Adp = ReadLine(excelOperations, bRow, j, 7),
						MdpNoPaCriteria = ReadLines(excelOperations, bRow, j, 8),
						MdpPaCriteria = ReadLines(excelOperations, bRow, j, 9),
						AdpCriteria = ReadLine(excelOperations, bRow, j, 11),
						MdpNoPaDop = ReadDopLines(excelOperations, bRow, j, 12),
						MdpPaDop = ReadDopLines(excelOperations, bRow, j, 13),
						AdpDop = ReadDopLines(excelOperations, bRow, j, 14)
					});
				}
				list.Add(new MdpBuilder
				{
					ShemeName = str,
					ShemeNum = str2,
					TnvList = list2
				});
			}
			excelOperations.AddList("new");
			int num3 = 10;
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
			excelOperations.FormatCells(1, 1, 2, array.Count(), bold: true, italic: false, Color.PowderBlue.ToArgb());
			int num4 = 3;
			foreach (MdpBuilder item in list)
			{
				excelOperations.setVal(num4, 1, item.ShemeNum);
				excelOperations.Format(num4, 1, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
				excelOperations.setVal(num4, 2, item.ShemeName);
				excelOperations.Merge(num4, 2, num4, array.Count());
				excelOperations.Format(num4, 2, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Center);
				excelOperations.FormatCells(num4, 1, num4, array.Count(), bold: false, italic: false, Color.MistyRose.ToArgb());
				double value = Math.Ceiling(2.5 * (double)num3 * (double)(item.ShemeName.Length / array.Sum((int x) => x)));
				excelOperations.Height(num4, Math.Max(15, Convert.ToInt32(value)));
				int num5 = num4 + 1;
				excelOperations.setVal(num5, 1, item.ShemeNum);
				excelOperations.Merge(num5, 1, num5 + item.TnvList.Count - 1, 1);
				excelOperations.Format(num5, 1, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
				excelOperations.setVal(num5, 2, item.ShemeName);
				excelOperations.Merge(num5, 2, num5 + item.TnvList.Count - 1, 2);
				excelOperations.Format(num5, 2, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Center);
				foreach (TNV tnv in item.TnvList)
				{
					excelOperations.setVal(num5, 3, tnv.Tnv);
					excelOperations.Format(num5, 3, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					excelOperations.setVal(num5, 4, "");
					excelOperations.Format(num5, 4, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
					excelOperations.setVal(num5, 5, "");
					excelOperations.Format(num5, 5, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
					foreach (MDP item2 in tnv.MdpNoPA.Where((MDP mDP) => mDP.Criteria != ""))
					{
						bool flag2 = item2 == tnv.MdpNoPA.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault();
						excelOperations.CellRichText(num5, 4, (!flag2) ? (item2.Criteria + Environment.NewLine) : item2.Criteria, (item2.Num != -1) ? $"{item2.Num}) " : "");
					}
					foreach (MDP item3 in tnv.MdpPa.Where((MDP mDP) => mDP.Criteria != ""))
					{
						bool flag3 = item3 == tnv.MdpPa.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault();
						excelOperations.CellRichText(num5, 5, (!flag3) ? (item3.Criteria + Environment.NewLine) : item3.Criteria, (item3.Num != -1) ? $"{item3.Num}) " : "");
					}
					if (tnv.Adp != "")
					{
						excelOperations.setVal(num4 + 1, 6, tnv.Adp);
						excelOperations.Merge(num4 + 1, 6, num4 + item.TnvList.Count, 6);
						excelOperations.Format(num4 + 1, 6, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					}
					string text4 = "";
					foreach (MDP item4 in tnv.MdpNoPaCriteria.Where((MDP mDP) => mDP.Criteria != ""))
					{
						string text5 = ((item4 == tnv.MdpNoPaCriteria.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
						text4 = text4 + $"{item4.Num}) {item4.Criteria}" + text5;
					}
					excelOperations.setVal(num5, 7, text4);
					excelOperations.Format(num5, 7, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
					excelOperations.CellComment(num5, 4, text4);
					string text6 = "";
					foreach (MDP item5 in tnv.MdpPaCriteria.Where((MDP mDP) => mDP.Criteria != ""))
					{
						string text7 = ((item5 == tnv.MdpPaCriteria.Where((MDP mDP) => mDP.Criteria != "").LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
						text6 = text6 + $"{item5.Num}) {item5.Criteria}" + text7;
					}
					excelOperations.setVal(num5, 8, text6);
					excelOperations.Format(num5, 8, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top);
					excelOperations.CellComment(num5, 5, text6);
					if (tnv.AdpCriteria != "")
					{
						excelOperations.setVal(num4 + 1, 9, tnv.AdpCriteria);
						excelOperations.Merge(num4 + 1, 9, num4 + item.TnvList.Count, 9);
						excelOperations.Format(num4 + 1, 9, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					}
					string text8 = "";
					foreach (string item6 in tnv.MdpNoPaDop)
					{
						string text9 = ((item6 == tnv.MdpNoPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
						text8 = text8 + item6 + text9;
					}
					excelOperations.setVal(num5, 10, text8);
					excelOperations.Format(num5, 10, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					string text10 = "";
					foreach (string item7 in tnv.MdpPaDop)
					{
						string text11 = ((item7 == tnv.MdpPaDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
						text10 = text10 + item7 + text11;
					}
					excelOperations.setVal(num5, 11, text10);
					excelOperations.Format(num5, 11, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					string text12 = "";
					foreach (string item8 in tnv.AdpDop)
					{
						string text13 = ((item8 == tnv.AdpDop.LastOrDefault()) ? "" : (Environment.NewLine ?? ""));
						text12 = text12 + item8 + text13;
					}
					excelOperations.setVal(num5, 12, text12);
					excelOperations.Format(num5, 12, ExcelHorizontalAlignment.Center, ExcelVerticalAlignment.Center);
					num5++;
				}
				excelOperations.GroupRows(num4 + 1, num5 - 1, 1, hide: false);
				num4 = num5;
			}
			excelOperations.Font("Liberation Serif", num3);
			excelOperations.Borders(1, 1, num4 - 1, array.Count());
			excelOperations.GroupRowsPosition();
			string text14 = Path.Combine(Path.GetDirectoryName(text2) ?? "", Path.GetFileNameWithoutExtension(text2) + "_modify.xlsx");
			excelOperations.Save(text14);
			Console.WriteLine("Файл успешно обработан и сохранен: " + text14);
			Console.WriteLine("Работа программы успешно завершена.");
		}
		else
		{
			Console.WriteLine("Путь не получен.");
		}
		Console.WriteLine("");
		Console.ReadKey();
	}

	public static string ReadLine(ExcelOperations ex, int bRow, int eRow, int col)
	{
		string result = "";
		for (int i = bRow; i <= eRow; i++)
		{
			if (ex.getStr(i, col) != "" && ex.getStr(i, col) != " ")
			{
				result = ex.getStr(i, col).Trim(new char[1] { ' ' });
			}
		}
		return result;
	}

	public static List<string> ReadDopLines(ExcelOperations ex, int bRow, int eRow, int col)
	{
		List<string> list = new List<string>();
		for (int i = bRow; i <= eRow; i++)
		{
			string text = ex.getStr(i, col).Trim(new char[1] { ' ' });
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
			string text = ex.getStr(i, col).Trim(new char[1] { ' ' });
			if (!text.Equals("Минимальное из") && text != "" && text != " ")
			{
				int num = text.IndexOf(' ');
				if (num > 0 && text[num - 1] == ')')
				{
					string value = text.Substring(0, num - 1);
					string criteria = (modify ? CellModifyString(text.Substring(num + 1)) : text.Substring(num + 1));
					list.Add(new MDP
					{
						Num = Convert.ToInt32(value),
						Criteria = criteria
					});
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
		text = text.Replace("-", " - ").Replace("+", " + ").Replace(",", ", ")
			.Replace("  ", " ")
			.Replace("==", "=")
			.Replace("*", "·");
		BracketNode node = Parse(text);
		return Reconstruct(node) ?? "";
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
