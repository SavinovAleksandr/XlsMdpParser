using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
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
			_excel.SaveAs(new FileInfo(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "", "tmp.xlsx")));
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
