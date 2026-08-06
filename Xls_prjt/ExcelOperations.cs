using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

	public void RenameSheet(string fromName, string toName)
	{
		ExcelWorksheet sheet = _excel.Workbook.Worksheets[fromName];
		if (sheet == null || string.Equals(fromName, toName, StringComparison.OrdinalIgnoreCase))
		{
			return;
		}
		if (_excel.Workbook.Worksheets[toName] != null)
		{
			_excel.Workbook.Worksheets.Delete(toName);
		}
		sheet.Name = toName;
	}

	public void ActivateSheet(string sheetName)
	{
		ExcelWorksheet sheet = _excel.Workbook.Worksheets[sheetName];
		if (sheet != null)
		{
			_ws = sheet;
		}
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
		CellRichText(i, j, val, prefix, Color.Black);
	}

	public void CellRichText(int i, int j, string val, string prefix, Color textColor)
	{
		ExcelRange excelRange = _ws.Cells[i, j];
		ExcelRichText excelRichText2 = excelRange.RichText.Add(prefix);
		excelRichText2.Color = textColor;
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
				excelRichText3.Color = textColor;
				excelRichText3.Bold = false;
			}
			return;
		}
		string[] array = Regex.Split(val, "(\\s+|\\|)");
		foreach (string text in array)
		{
			if (string.IsNullOrEmpty(text) || text == "|")
			{
				continue;
			}
			string text2 = text;
			string text3 = text.Trim();
			switch (text3)
			{
			default:
				if (!(text3 == "and"))
				{
					break;
				}
				goto case "+";
			case "+":
			case "-":
			case "or":
				text2 = " " + text3 + " ";
				break;
			}
			ExcelRichText excelRichText = excelRange.RichText.Add(text2);
			if (text3 == "if" || text3 == "{" || text3 == "}")
			{
				excelRichText.Color = Color.Red;
				excelRichText.Bold = true;
				continue;
			}
			switch (text3)
			{
			default:
				if (!(text3 == "]"))
				{
					if (text3 == "and" || text3 == "or")
					{
						excelRichText.Color = Color.Blue;
						excelRichText.Bold = true;
					}
					else
					{
						excelRichText.Color = textColor;
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

	public void ClearCell(int i, int j)
	{
		_ws.Cells[i, j].Value = null;
		if (_ws.Cells[i, j].IsRichText)
		{
			_ws.Cells[i, j].RichText.Clear();
		}
	}

	public void AppendColoredText(int i, int j, string text, Color color, bool bold = false)
	{
		if (string.IsNullOrEmpty(text))
		{
			return;
		}
		ExcelRichText excelRichText = _ws.Cells[i, j].RichText.Add(text);
		excelRichText.Color = color;
		excelRichText.Bold = bold;
	}

	public void CellComment(int i, int j, string str)
	{
		if (string.IsNullOrWhiteSpace(str))
		{
			return;
		}
		str = NormalizeExcelNewlines(str);
		ExcelRange excelRange = _ws.Cells[i, j];
		if (excelRange.Comment != null)
		{
			_ws.Comments.Remove(excelRange.Comment);
		}
		ExcelComment excelComment = excelRange.AddComment(str, "XlsxMdpParser");
		// AutoFit даёт узкое и очень высокое окно — задаём широкий прямоугольник вправо.
		excelComment.AutoFit = false;
		SizeCommentBox(excelComment, i, j, str);
	}

	private static string NormalizeExcelNewlines(string text)
	{
		return (text ?? "")
			.Replace("_x000A_", "\n")
			.Replace("\r\n", "\n")
			.Replace('\r', '\n');
	}

	private static void SizeCommentBox(ExcelComment comment, int cellRow, int cellCol, string text)
	{
		string[] lines = NormalizeExcelNewlines(text)
			.Split('\n')
			.Where((string l) => !string.IsNullOrWhiteSpace(l))
			.ToArray();
		if (lines.Length == 0)
		{
			lines = new string[1] { text };
		}

		int maxLine = lines.Max((string l) => l.Length);
		// Широкий прямоугольник вправо.
		int widthChars = Math.Min(Math.Max(maxLine, 40), 100);
		int colSpan = Math.Min(11, Math.Max(6, (widthChars + 12) / 12));

		// Сколько символов реально помещается по ширине окна (~9 символов на колонку якоря).
		int charsPerVisualLine = Math.Max(40, colSpan * 9);
		int visualLines = 0;
		foreach (string line in lines)
		{
			visualLines += Math.Max(1, (line.Length + charsPerVisualLine - 1) / charsPerVisualLine);
		}

		// Высота в пикселях якоря, НЕ в строках листа (строки листа у нас очень высокие).
		const int padPx = 10;
		const int linePx = 15;
		int heightPx = padPx + visualLines * linePx;
		heightPx = Math.Min(160, Math.Max(28, heightPx));

		comment.From.Row = cellRow - 1;
		comment.From.Column = cellCol;
		comment.From.RowOffset = 2;
		comment.From.ColumnOffset = 8;
		// To в той же строке листа — высота только через RowOffset.
		comment.To.Row = comment.From.Row;
		comment.To.Column = comment.From.Column + colSpan;
		comment.To.RowOffset = comment.From.RowOffset + heightPx;
		comment.To.ColumnOffset = 0;
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

	/// <summary>
	/// Блок «УТВЕРЖДАЮ» справа на листе сводки (как исходный блок Говоруна в C1).
	/// </summary>
	public void SetSummaryApprovalBlock(string sheetName, string address, string value)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null || string.IsNullOrWhiteSpace(value))
		{
			return;
		}
		ExcelRange cell = excelWorksheet.Cells[address];
		cell.Value = value.Trim();
		cell.Style.WrapText = true;
		cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
		cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
		cell.Style.Font.Name = "Liberation Serif";
		cell.Style.Font.Size = 14f;
		cell.Style.Font.Bold = false;
		cell.Style.Font.Italic = false;
		// Не оставляем старый левый блок (B1), если ранее писали туда.
		if (!string.Equals(address, "B1", StringComparison.OrdinalIgnoreCase))
		{
			ExcelRange leftCell = excelWorksheet.Cells["B1"];
			if (leftCell.Value != null && leftCell.Value.ToString().IndexOf("УТВЕРЖДАЮ", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				leftCell.Value = null;
			}
		}
	}

	public void AutoFitSheetRowsByContent(string sheetName, int startRow, int minHeight = 15, double extraHeightFactor = 1.0, int[] includeColumns = null)
	{
		ExcelWorksheet excelWorksheet = _excel.Workbook.Worksheets[sheetName];
		if (excelWorksheet == null || excelWorksheet.Dimension == null)
		{
			return;
		}
		int firstRow = Math.Max(startRow, 1);
		int lastRow = excelWorksheet.Dimension.End.Row;
		int lastCol = excelWorksheet.Dimension.End.Column;
		List<int> columns = new List<int>();
		if (includeColumns == null || includeColumns.Length == 0)
		{
			for (int c = 1; c <= lastCol; c++)
			{
				columns.Add(c);
			}
		}
		else
		{
			foreach (int col in includeColumns.Distinct().OrderBy((int c) => c))
			{
				if (col >= 1 && col <= lastCol)
				{
					columns.Add(col);
				}
			}
		}
		if (columns.Count == 0)
		{
			return;
		}

		Dictionary<string, ExcelAddress> mergeCache = new Dictionary<string, ExcelAddress>(StringComparer.Ordinal);
		// Требуемая суммарная высота для вертикальных объединений: [startRow, endRow, neededPt]
		List<double[]> verticalMergeNeeds = new List<double[]>();

		double padFactor = Math.Max(1.0, extraHeightFactor);
		for (int r = firstRow; r <= lastRow; r++)
		{
			int maxVisualLines = 1;
			float maxFontSize = 11f;
			foreach (int col in columns)
			{
				// Не пропускаем колонки с малой шириной: на сводке объединения A:C
				// начинаются в A (ширина ~0), иначе текст не измеряется.
				if (excelWorksheet.Column(col).Hidden)
				{
					continue;
				}
				ExcelRange cell = excelWorksheet.Cells[r, col];
				int widthStart = col;
				int widthEnd = col;
				string text;
				float fontSize = cell.Style.Font.Size > 0 ? cell.Style.Font.Size : 11f;
				if (cell.Merge)
				{
					string mergeAddr = excelWorksheet.MergedCells[r, col];
					if (string.IsNullOrWhiteSpace(mergeAddr))
					{
						continue;
					}
					if (!mergeCache.TryGetValue(mergeAddr, out ExcelAddress merge))
					{
						merge = new ExcelAddress(mergeAddr);
						mergeCache[mergeAddr] = merge;
					}
					// Текст и ширину считаем только у верхней-левой ячейки объединения.
					if (merge.Start.Row != r || merge.Start.Column != col)
					{
						continue;
					}
					widthStart = merge.Start.Column;
					widthEnd = merge.End.Column;
					ExcelRange topLeft = excelWorksheet.Cells[merge.Start.Row, merge.Start.Column];
					text = GetRangeDisplayText(topLeft);
					fontSize = topLeft.Style.Font.Size > 0 ? topLeft.Style.Font.Size : fontSize;
					topLeft.Style.WrapText = true;
				}
				else
				{
					text = GetRangeDisplayText(cell);
					cell.Style.WrapText = true;
				}
				if (string.IsNullOrWhiteSpace(text))
				{
					continue;
				}
				// Скрытые колонки (напр. МДП с ПА) не дают видимой ширины в Excel —
				// иначе длинные розовые заголовки B:L считаются в 1 строку и обрезаются.
				double widthUnits = GetVisibleColumnsWidth(excelWorksheet, widthStart, widthEnd);
				if (widthUnits < 1.0)
				{
					continue;
				}
				int visualLines = CountWrappedLines(text, widthUnits);
				maxVisualLines = Math.Max(maxVisualLines, visualLines);
				maxFontSize = Math.Max(maxFontSize, fontSize);

				if (cell.Merge)
				{
					string mergeAddr = excelWorksheet.MergedCells[r, col];
					if (!string.IsNullOrWhiteSpace(mergeAddr) && mergeCache.TryGetValue(mergeAddr, out ExcelAddress merge) && merge.End.Row > merge.Start.Row)
					{
						double needed = EstimateRowHeightPoints(visualLines, fontSize, padFactor, minHeight);
						verticalMergeNeeds.Add(new double[3] { merge.Start.Row, merge.End.Row, needed });
					}
				}
			}
			excelWorksheet.Row(r).Height = EstimateRowHeightPoints(maxVisualLines, maxFontSize, padFactor, minHeight);
			excelWorksheet.Row(r).CustomHeight = true;
		}

		foreach (double[] need in verticalMergeNeeds)
		{
			int mergeStart = (int)need[0];
			int mergeEnd = (int)need[1];
			double needed = need[2];
			double actual = 0.0;
			for (int r = mergeStart; r <= mergeEnd; r++)
			{
				actual += excelWorksheet.Row(r).Height > 0.0 ? excelWorksheet.Row(r).Height : 15.0;
			}
			if (actual + 0.05 >= needed)
			{
				continue;
			}
			double add = (needed - actual) / (mergeEnd - mergeStart + 1);
			for (int r = mergeStart; r <= mergeEnd; r++)
			{
				double cur = excelWorksheet.Row(r).Height > 0.0 ? excelWorksheet.Row(r).Height : 15.0;
				excelWorksheet.Row(r).Height = cur + add;
			}
		}
	}

	private static string GetRangeDisplayText(ExcelRange cell)
	{
		if (cell == null)
		{
			return "";
		}
		try
		{
			if (cell.IsRichText)
			{
				string rich = cell.RichText?.Text;
				if (!string.IsNullOrEmpty(rich))
				{
					return rich;
				}
			}
		}
		catch
		{
		}
		return cell.Value?.ToString() ?? "";
	}

	private static double GetVisibleColumnsWidth(ExcelWorksheet sheet, int startCol, int endCol)
	{
		double widthUnits = 0.0;
		for (int c = startCol; c <= endCol; c++)
		{
			ExcelColumn column = sheet.Column(c);
			if (column.Hidden)
			{
				continue;
			}
			double w = column.Width;
			// Колонки с почти нулевой шириной фактически не видны (как на листе сводки A/D/E).
			if (w < 0.5)
			{
				continue;
			}
			widthUnits += w;
		}
		return widthUnits;
	}

	private static int CountWrappedLines(string text, double columnWidthUnits)
	{
		text = (text ?? "").Replace("_x000A_", "\n").Replace("\r\n", "\n").Replace('\r', '\n');
		// Ширина колонки Excel ≈ символы '0'. Кириллица Liberation Serif шире —
		// берём 0.88, чтобы переносы не недооценивались (иначе текст обрезается).
		double charsPerWidth = 0.88;
		int charsPerLine = Math.Max(6, (int)Math.Floor(Math.Max(1.0, columnWidthUnits) * charsPerWidth));
		int lines = 0;
		foreach (string part in text.Split('\n'))
		{
			int len = Math.Max(1, part.TrimEnd().Length);
			lines += Math.Max(1, (int)Math.Ceiling(len / (double)charsPerLine));
		}
		return Math.Max(1, lines);
	}

	private static double EstimateRowHeightPoints(int visualLines, float fontSize, double padFactor, int minHeight)
	{
		double size = fontSize > 0 ? fontSize : 11.0;
		// Межстрочный интервал Excel ≈ 1.25 размера шрифта; небольшой запас на поля.
		double height = visualLines * size * 1.25 * padFactor + 4.0;
		return Math.Max(minHeight, Math.Ceiling(height));
	}

	public string getStr(int i, int j)
	{
		return (_ws.Cells[i, j].Value != null) ? _ws.Cells[i, j].Value.ToString() : "";
	}

	/// <summary>
	/// Читает значение ячейки; если она часть объединения и пуста — берёт верхний левый угол.
	/// Не использовать для детекта границ схемы (там пустые «хвосты» merge как раз нужны).
	/// </summary>
	public string getStrMerged(int i, int j)
	{
		ExcelRange cell = _ws.Cells[i, j];
		if (cell.Value != null)
		{
			return cell.Value.ToString();
		}
		if (!cell.Merge)
		{
			return "";
		}
		string merged = MergedCells(i, j);
		if (string.IsNullOrWhiteSpace(merged) || !merged.Contains(":"))
		{
			return "";
		}
		string topLeft = merged.Split(':')[0];
		object value = _ws.Cells[topLeft].Value;
		return value != null ? value.ToString() : "";
	}

	public bool CellHasOwnValue(int i, int j)
	{
		return _ws.Cells[i, j].Value != null && !string.IsNullOrWhiteSpace(_ws.Cells[i, j].Value.ToString());
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
