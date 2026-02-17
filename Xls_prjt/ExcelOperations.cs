using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Xls_prjt;

public class ExcelOperations
{
	private ExcelPackage _excel;

	private ExcelWorksheet _ws;

	public ExcelOperations(string file, dynamic list)
	{
		ExcelPackage.License.SetNonCommercialPersonal("igv");
		FileInfo newFile = new FileInfo(file);
		_excel = new ExcelPackage(newFile);
		_ws = _excel.Workbook.Worksheets[list];
	}

	public ExcelOperations(string list = "Результат")
	{
		ExcelPackage.License.SetNonCommercialPersonal("igv");
		_excel = new ExcelPackage();
		_ws = _excel.Workbook.Worksheets.Add(list);
		_ws.Cells["A1:XFD1048576"].Style.WrapText = true;
	}

	public int SheetCount(string file)
	{
		ExcelPackage.License.SetNonCommercialPersonal("igv");
		FileInfo newFile = new FileInfo(file);
		_excel = new ExcelPackage(newFile);
		return _excel.Workbook.Worksheets.Count;
	}

	public void AddList(string list)
	{
		if (_excel.Workbook.Worksheets[list] != null)
		{
			_excel.Workbook.Worksheets.Delete(list);
		}
		_ws = _excel.Workbook.Worksheets.Add(list);
		_ws.Cells["A1:XFD1048576"].Style.WrapText = true;
	}

	public int LastColumnRow()
	{
		return _ws.Dimension.End.Row;
	}

	public string MergedCells(int i, int j)
	{
		if (_ws.Cells[i, j].Merge)
		{
			int mergeCellId = _ws.GetMergeCellId(i, j);
			return _ws.MergedCells[mergeCellId - 1];
		}
		return _ws.Cells[i, j].Address + ":" + _ws.Cells[i, j].Address;
	}

	public void GroupRowsPosition(bool param = false)
	{
		_ws.OutLineSummaryBelow = param;
	}

	public void FreezeRows(int rowCount)
	{
		_ws.View.FreezePanes(rowCount + 1, 1);
	}

	public void GroupRows(int i, int j, int level = 1, bool hide = true)
	{
		_ws.Rows[i, j].OutlineLevel = level;
		_ws.Rows[i, j].Collapsed = hide;
	}

	public void setVal(int i, int j, dynamic val, bool wrap = true)
	{
		_ws.Cells[i, j].Value = (object)val;
		_ws.Cells[i, j].Style.WrapText = wrap;
	}

	public void CellRichText(int i, int j, string val, string prefix)
	{
		ExcelRange excelRange = _ws.Cells[i, j];
		ExcelRichText excelRichText2 = excelRange.RichText.Add(prefix);
		excelRichText2.Color = Color.Black;
		excelRichText2.Bold = false;
		if (val.StartsWith("Минимальное из", StringComparison.OrdinalIgnoreCase))
		{
			string text = "Минимальное из:";
			ExcelRichText excelRichText = excelRange.RichText.Add(text);
			excelRichText.Color = Color.Green;
			excelRichText.Bold = true;
			string text2 = val.Substring(Math.Min(text.Length, val.Length));
			if (text2.Length > 0)
			{
				ExcelRichText excelRichText3 = excelRange.RichText.Add(text2);
				excelRichText3.Color = Color.Black;
				excelRichText3.Bold = false;
			}
			return;
		}
		string[] array = val.Split('|', ' ');
		foreach (string text in array)
		{
			string text2 = text.Replace(",", ", ");
			switch (text)
			{
			default:
				if (!(text == "and"))
				{
					break;
				}
				goto case "+";
			case "+":
			case "-":
			case "or":
				text2 = " " + text + " ";
				break;
			}
			ExcelRichText excelRichText = excelRange.RichText.Add(text2);
			if (text == "if" || text == "{" || text == "}")
			{
				excelRichText.Color = Color.Red;
				excelRichText.Bold = true;
				continue;
			}
			switch (text)
			{
			default:
				if (!(text == "]"))
				{
					if (text == "and" || text == "or")
					{
						excelRichText.Color = Color.Blue;
						excelRichText.Bold = true;
					}
					else
					{
						excelRichText.Color = Color.Black;
						excelRichText.Bold = false;
					}
					break;
				}
				goto case "min";
			case "min":
			case "max":
			case "[":
				excelRichText.Color = Color.Green;
				excelRichText.Bold = true;
				break;
			}
		}
	}

	public void CellComment(int i, int j, string str)
	{
		if (string.IsNullOrWhiteSpace(str))
		{
			return;
		}
		ExcelRange excelRange = _ws.Cells[i, j];
		ExcelComment excelComment = excelRange.AddComment(str);
		excelComment.AutoFit = true;
	}

	public void Wrap(int i, int j, bool wrap = true)
	{
		_ws.Cells[i, j].Style.WrapText = wrap;
	}

	public void setVal(string param, dynamic val)
	{
		_ws.Cells[param].Value = (object)val;
	}

	public void SetSheetCellValue(string sheetName, string address, string value, bool wrap = true)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null)
		{
			return;
		}
		excelWorksheet.Cells[address].Value = value;
		excelWorksheet.Cells[address].Style.WrapText = wrap;
	}

	public void SetSheetCellAlignment(string sheetName, string address, ExcelHorizontalAlignment horizontal, ExcelVerticalAlignment vertical)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null)
		{
			return;
		}
		excelWorksheet.Cells[address].Style.HorizontalAlignment = horizontal;
		excelWorksheet.Cells[address].Style.VerticalAlignment = vertical;
	}

	public void AutoFitSheetRowsByContent(string sheetName, int startRow, int minHeight = 15, double extraHeightFactor = 1.0, int[] includeColumns = null)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null || excelWorksheet.Dimension == null)
		{
			return;
		}
		HashSet<int> hashSet = null;
		if (includeColumns != null && includeColumns.Length != 0)
		{
			hashSet = new HashSet<int>(includeColumns);
		}
		int num = Math.Max(startRow, 1);
		int row = excelWorksheet.Dimension.End.Row;
		int column = excelWorksheet.Dimension.End.Column;
		List<int> list = new List<int>();
		if (hashSet == null)
		{
			for (int i = 1; i <= column; i++)
			{
				list.Add(i);
			}
		}
		else
		{
			foreach (int item in hashSet)
			{
				if (item >= 1 && item <= column)
				{
					list.Add(item);
				}
			}
			list.Sort();
		}
		if (list.Count == 0)
		{
			return;
		}
		Dictionary<string, ExcelAddress> dictionary = new Dictionary<string, ExcelAddress>(StringComparer.Ordinal);
		List<double[]> list2 = new List<double[]>();
		for (int i = num; i <= row; i++)
		{
			int num2 = 1;
			foreach (int item2 in list)
			{
				ExcelRange excelRange = excelWorksheet.Cells[i, item2];
				int num3 = item2;
				int num4 = item2;
				string text;
				if (excelRange.Merge)
				{
					string text2 = excelWorksheet.MergedCells[i, item2];
					if (string.IsNullOrWhiteSpace(text2))
					{
						continue;
					}
					if (!dictionary.TryGetValue(text2, out var value))
					{
						value = new ExcelAddress(text2);
						dictionary[text2] = value;
					}
					if (value.Start.Row != i || value.Start.Column != item2)
					{
						continue;
					}
					num3 = value.Start.Column;
					num4 = value.End.Column;
					text = excelWorksheet.Cells[value.Start.Row, value.Start.Column].Value?.ToString();
				}
				else
				{
					text = excelRange.Value?.ToString();
				}
				if (string.IsNullOrWhiteSpace(text))
				{
					continue;
				}
				text = text.Replace("_x000A_", "\n");
				double num5 = 0.0;
				for (int j = num3; j <= num4; j++)
				{
					double width = excelWorksheet.Column(j).Width;
					num5 += ((width > 0.0) ? width : 8.43);
				}
				// Column G usually contains the longest narrative criteria text.
				// Use a tighter char-per-width estimate there to avoid clipped wrapped lines.
				double num6Factor = 1.9;
				if (item2 == 7)
				{
					// For long criteria text in G use a stricter estimate.
					num6Factor = ((text.Length > 160) ? 1.2 : 1.35);
				}
				int num6 = Math.Max(8, (int)Math.Round(num5 * num6Factor));
				int num7 = 0;
				string[] array = text.Split('\n');
				foreach (string text3 in array)
				{
					int num8 = Math.Max(1, text3.TrimEnd().Length);
					num7 += Math.Max(1, (int)Math.Ceiling((double)num8 / (double)num6));
				}
				if (item2 == 7 && text.Length > 120)
				{
					num7++;
				}
				num2 = Math.Max(num2, num7);
				excelWorksheet.Cells[i, item2].Style.WrapText = true;
				if (excelRange.Merge)
				{
					string text4 = excelWorksheet.MergedCells[i, item2];
					if (!string.IsNullOrWhiteSpace(text4) && dictionary.TryGetValue(text4, out var value2) && value2.End.Row > value2.Start.Row)
					{
						double num9 = Math.Max(1.0, extraHeightFactor);
						double num10 = Math.Max((double)minHeight, (double)num7 * 13.8 * num9 + 1.5);
						list2.Add(new double[3] { value2.Start.Row, value2.End.Row, num10 });
					}
				}
			}
			double num11 = Math.Max(1.0, extraHeightFactor);
			int height = Math.Max(minHeight, (int)Math.Ceiling((double)num2 * 13.8 * num11 + 1.5));
			excelWorksheet.Row(i).Height = height;
		}
		foreach (double[] item3 in list2)
		{
			int num12 = (int)item3[0];
			int num13 = (int)item3[1];
			double num14 = item3[2];
			double num15 = 0.0;
			for (int k = num12; k <= num13; k++)
			{
				num15 += ((excelWorksheet.Row(k).Height > 0.0) ? excelWorksheet.Row(k).Height : 15.0);
			}
			if (!(num15 + 0.1 < num14))
			{
				continue;
			}
			double num16 = num14 - num15;
			int num17 = num13 - num12 + 1;
			double num18 = num16 / (double)num17;
			for (int l = num12; l <= num13; l++)
			{
				double num19 = (excelWorksheet.Row(l).Height > 0.0) ? excelWorksheet.Row(l).Height : 15.0;
				excelWorksheet.Row(l).Height = num19 + num18;
			}
		}
	}

	public string getStr(int i, int j)
	{
		return (_ws.Cells[i, j].Value != null) ? _ws.Cells[i, j].Value.ToString() : "";
	}

	public string getStr(string param)
	{
		return (_ws.Cells[param].Value != null) ? _ws.Cells[param].Value.ToString() : "";
	}

	public int getInt(int i, int j)
	{
		return (_ws.Cells[i, j].Value != null) ? Convert.ToInt32(_ws.Cells[i, j].Value) : 0;
	}

	public int getInt(string param)
	{
		return (_ws.Cells[param].Value != null) ? Convert.ToInt32(_ws.Cells[param].Value) : 0;
	}

	public double getDbl(int i, int j)
	{
		return (_ws.Cells[i, j].Value != null) ? Convert.ToDouble(_ws.Cells[i, j].Value) : 0.0;
	}

	public double getDbl(string param)
	{
		return (_ws.Cells[param].Value != null) ? Convert.ToDouble(_ws.Cells[param].Value) : 0.0;
	}

	public void Save(string file = "")
	{
		if (file != "")
		{
			_excel.SaveAs(new FileInfo(file));
		}
		else
		{
			_excel.SaveAs(new FileInfo(Path.Combine(AppContext.BaseDirectory, "tmp.xlsx")));
		}
	}

	public void Borders(string param)
	{
		_ws.Cells[param].Style.Border.Top.Style = ExcelBorderStyle.Thin;
		_ws.Cells[param].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		_ws.Cells[param].Style.Border.Left.Style = ExcelBorderStyle.Thin;
		_ws.Cells[param].Style.Border.Right.Style = ExcelBorderStyle.Thin;
	}

	public void Borders(int bRow, int bCol, int eRow, int eCol)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
	}

	public void FormatCells(string param, bool bold, bool italic = false)
	{
		_ws.Cells[param].Style.Font.Bold = bold;
		_ws.Cells[param].Style.Font.Italic = italic;
	}

	public void FormatCells(int i, int j, bool bold, bool italic = false)
	{
		_ws.Cells[i, j].Style.Font.Bold = bold;
		_ws.Cells[i, j].Style.Font.Italic = italic;
	}

	public void FormatCells(int bRow, int bCol, int eRow, int eCol, bool bold, bool italic = false)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Font.Bold = bold;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Font.Italic = italic;
	}

	public void FormatCells(string param, bool bold, bool italic = false, int _color = -329006)
	{
		_ws.Cells[param].Style.Font.Bold = bold;
		_ws.Cells[param].Style.Font.Italic = italic;
		_ws.Cells[param].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[param].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void FormatCells(int i, int j, bool bold, bool italic = false, int _color = -329006)
	{
		_ws.Cells[i, j].Style.Font.Bold = bold;
		_ws.Cells[i, j].Style.Font.Italic = italic;
		_ws.Cells[i, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[i, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void FormatCells(int bRow, int bCol, int eRow, int eCol, bool bold, bool italic = false, int _color = -329006)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Font.Bold = bold;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Font.Italic = italic;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void FormatCells(string param, int _color = -329006)
	{
		_ws.Cells[param].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[param].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void FormatCells(int i, int j, int _color = -329006)
	{
		_ws.Cells[i, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[i, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void FormatCells(int bRow, int bCol, int eRow, int eCol, int _color = -329006)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(_color));
	}

	public void Merge(string param)
	{
		_ws.Cells[param].Merge = true;
	}

	public void Merge(int bRow, int bCol, int eRow, int eCol, bool hor = false, bool vert = false)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Merge = true;
		if (hor)
		{
			_ws.Cells[bRow, bCol, eRow, eCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
		}
		if (vert)
		{
			_ws.Cells[bRow, bCol, eRow, eCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
		}
	}

	public void Format(int i, int j, ExcelHorizontalAlignment excelHorizontalAlignment, ExcelVerticalAlignment excelVerticalAlignment, int rotation = 0)
	{
		_ws.Cells[i, j].Style.HorizontalAlignment = excelHorizontalAlignment;
		_ws.Cells[i, j].Style.VerticalAlignment = excelVerticalAlignment;
		_ws.Cells[i, j].Style.TextRotation = rotation;
	}

	public void Format(int bRow, int bCol, int eRow, int eCol, ExcelHorizontalAlignment excelHorizontalAlignment, ExcelVerticalAlignment excelVerticalAlignment, int rotation = 0)
	{
		_ws.Cells[bRow, bCol, eRow, eCol].Style.HorizontalAlignment = excelHorizontalAlignment;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.VerticalAlignment = excelVerticalAlignment;
		_ws.Cells[bRow, bCol, eRow, eCol].Style.TextRotation = rotation;
	}

	public void Font(string name = "Times New Roman", int size = 10)
	{
		_ws.Cells["A1:XFD1048576"].Style.Font.Name = name;
		_ws.Cells["A1:XFD1048576"].Style.Font.Size = size;
	}

	public void FontColor(int i, int j, Color color)
	{
		_ws.Cells[i, j].Style.Font.Color.SetColor(color);
	}

	public void FontStyle(int i, int j, float size, bool italic = false, bool bold = false)
	{
		_ws.Cells[i, j].Style.Font.Size = size;
		_ws.Cells[i, j].Style.Font.Italic = italic;
		_ws.Cells[i, j].Style.Font.Bold = bold;
	}

	public bool IsValue(string param)
	{
		return _ws.Cells[param].Value != null;
	}

	public bool IsValue(int i, int j)
	{
		return _ws.Cells[i, j].Value != null;
	}

	public void Width(int col, int width, bool auto = false)
	{
		_ws.Column(col).Width = width;
		if (auto)
		{
			_ws.Column(col).AutoFit();
		}
	}

	public void AutoFitWithMaxWidth(int col, int maxWidth)
	{
		ExcelColumn excelColumn = _ws.Column(col);
		excelColumn.AutoFit();
		double num = excelColumn.Width;
		double num2 = EstimateWidthByText(col);
		double width = Math.Max(num, num2);
		if (width > maxWidth)
		{
			width = maxWidth;
		}
		excelColumn.Width = Math.Max(4.0, width);
	}

	private double EstimateWidthByText(int col)
	{
		if (_ws.Dimension == null)
		{
			return 8.0;
		}
		int num = 0;
		for (int i = 1; i <= _ws.Dimension.End.Row; i++)
		{
			string text = _ws.Cells[i, col].Text ?? "";
			if (text.Length == 0)
			{
				continue;
			}
			string[] array = text.Replace("_x000A_", "\n").Split('\n');
			foreach (string text2 in array)
			{
				int length = text2.TrimEnd().Length;
				if (length > num)
				{
					num = length;
				}
			}
		}
		return (double)num * 1.1 + 2.0;
	}

	public void HideColumn(int col)
	{
		_ws.Column(col).Hidden = true;
	}

	public void Height(int row, int height)
	{
		_ws.Row(row).Height = height;
	}

	public double GetRowHeightOrDefault(int row, double defaultHeight = 15.0)
	{
		double height = _ws.Row(row).Height;
		if (height <= 0.0)
		{
			return defaultHeight;
		}
		return height;
	}

	public void UpdateSummarySheetHyperlinks(string summarySheetName, string targetSheetName, Dictionary<string, int> schemeRows)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[summarySheetName];
		if (excelWorksheet == null || excelWorksheet.Dimension == null)
		{
			return;
		}
		int num = 0;
		for (int i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
		{
			string text = excelWorksheet.Cells[i, 1].Value?.ToString()?.Trim() ?? "";
			if (text.Equals("Ремонтные схемы:", StringComparison.OrdinalIgnoreCase))
			{
				num = i + 1;
				break;
			}
		}
		if (num == 0)
		{
			return;
		}
		for (int j = num; j <= excelWorksheet.Dimension.End.Row; j++)
		{
			string text2 = excelWorksheet.Cells[j, 1].Value?.ToString()?.Trim() ?? "";
			if (text2.Length == 0)
			{
				continue;
			}
			Match match = Regex.Match(text2, "^(\\d+)\\.");
			if (!match.Success)
			{
				continue;
			}
			string key = match.Groups[1].Value;
			if (!schemeRows.TryGetValue(key, out var value))
			{
				continue;
			}
			ExcelRange excelRange = excelWorksheet.Cells[j, 1];
			excelRange.Hyperlink = new ExcelHyperLink($"'{targetSheetName}'!B{value}", text2);
		}
	}

	public void ConfigureSheetForPrint(string sheetName, bool repeatTopTwoRows = false)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null)
		{
			return;
		}
		excelWorksheet.PrinterSettings.Orientation = eOrientation.Landscape;
		excelWorksheet.PrinterSettings.PaperSize = ePaperSize.A4;
		excelWorksheet.PrinterSettings.FitToPage = true;
		excelWorksheet.PrinterSettings.FitToWidth = 1;
		excelWorksheet.PrinterSettings.FitToHeight = 0;
		excelWorksheet.PrinterSettings.HorizontalCentered = true;
		excelWorksheet.PrinterSettings.VerticalCentered = false;
		if (excelWorksheet.Dimension != null)
		{
			excelWorksheet.PrinterSettings.PrintArea = excelWorksheet.Cells[excelWorksheet.Dimension.Address];
		}
		if (repeatTopTwoRows)
		{
			excelWorksheet.PrinterSettings.RepeatRows = excelWorksheet.Cells["1:2"];
		}
	}

	public int ValToColor(dynamic value)
	{
		int result = Color.YellowGreen.ToArgb();
		if (value >= 30 && value < 40)
		{
			result = Color.YellowGreen.ToArgb();
		}
		else if (value >= 40 && value < 50)
		{
			result = Color.LightGreen.ToArgb();
		}
		else if (value >= 50 && value < 60)
		{
			result = Color.GreenYellow.ToArgb();
		}
		else if (value >= 60 && value < 70)
		{
			result = Color.Yellow.ToArgb();
		}
		else if (value >= 70 && value < 80)
		{
			result = Color.Orange.ToArgb();
		}
		else if (value >= 80 && value < 90)
		{
			result = Color.SandyBrown.ToArgb();
		}
		else if (value >= 90 && value < 100)
		{
			result = Color.Tomato.ToArgb();
		}
		else if (value >= 100)
		{
			result = Color.OrangeRed.ToArgb();
		}
		else if (value < 30)
		{
			result = Color.White.ToArgb();
		}
		return result;
	}

	public int ValToColorVoltage(dynamic value)
	{
		int result = Color.YellowGreen.ToArgb();
		if (value >= 10 && value <= 15)
		{
			result = Color.GreenYellow.ToArgb();
		}
		else if (value >= 8 && value < 10)
		{
			result = Color.Yellow.ToArgb();
		}
		else if (value >= 6 && value < 8)
		{
			result = Color.Orange.ToArgb();
		}
		else if (value >= 4 && value < 6)
		{
			result = Color.SandyBrown.ToArgb();
		}
		else if (value >= 2.5 && value < 4)
		{
			result = Color.Tomato.ToArgb();
		}
		else if (value <= 2.5)
		{
			result = Color.OrangeRed.ToArgb();
		}
		else if (value > 15)
		{
			result = Color.White.ToArgb();
		}
		return result;
	}

	public int VoltageToColor(dynamic value)
	{
		int result = Color.YellowGreen.ToArgb();
		if (value >= 16)
		{
			result = Color.OrangeRed.ToArgb();
		}
		else if (value >= 14 && value < 16)
		{
			result = Color.Tomato.ToArgb();
		}
		else if (value >= 12 && value < 14)
		{
			result = Color.SandyBrown.ToArgb();
		}
		else if (value >= 10 && value < 12)
		{
			result = Color.Orange.ToArgb();
		}
		else if (value >= 7.5 && value < 10)
		{
			result = Color.Yellow.ToArgb();
		}
		else if (value >= 5 && value < 7.5)
		{
			result = Color.GreenYellow.ToArgb();
		}
		else if (value <= 5)
		{
			result = Color.White.ToArgb();
		}
		return result;
	}
}
