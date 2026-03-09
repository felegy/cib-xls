using System.CommandLine;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UglyToad.PdfPig;

var inputOption = new Option<FileInfo>("--input", "-i")
{
	Description = "Input PDF file path.",
	Required = true,
};

var outputOption = new Option<FileInfo?>("--output", "-o")
{
	Description = "Output XLSX file path. If omitted, uses the input name with .xlsx extension.",
};

var sheetNameOption = new Option<string>("--sheet-name", "-s")
{
	Description = "Worksheet name for parsed transactions.",
	DefaultValueFactory = _ => "Transactions",
};

var rootCommand = new RootCommand("CIB PDF -> XLSX importer");
rootCommand.Add(inputOption);
rootCommand.Add(outputOption);
rootCommand.Add(sheetNameOption);

rootCommand.SetAction(parseResult =>
{
	var input = parseResult.GetValue(inputOption)!;
	var output = parseResult.GetValue(outputOption);
	var sheetName = parseResult.GetValue(sheetNameOption) ?? "Transactions";

	if (!input.Exists)
	{
		Console.Error.WriteLine($"Input file not found: {input.FullName}");
		return 1;
	}

	var outputPath = output?.FullName ?? Path.ChangeExtension(input.FullName, ".xlsx");
	var parsedRows = ExtractRows(input.FullName).ToList();
	var parsedTransactions = parsedRows
		.SelectMany(ExtractTransactionsFromPageText)
		.ToList();

	WriteWorkbook(outputPath, parsedRows, parsedTransactions, sheetName);

	Console.WriteLine($"PDF sorok: {parsedRows.Count}");
	Console.WriteLine($"Feldolgozott tranzakciok: {parsedTransactions.Count}");
	Console.WriteLine($"Kesz: {outputPath}");

	return 0;
});

var parseResult = rootCommand.Parse(args);
return await parseResult.InvokeAsync();

static IEnumerable<RawPdfLine> ExtractRows(string pdfPath)
{
	using var document = PdfDocument.Open(pdfPath);
	var lineSplitter = new Regex(@"\r\n|\n|\r", RegexOptions.Compiled);

	for (var pageNumber = 1; pageNumber <= document.NumberOfPages; pageNumber++)
	{
		var page = document.GetPage(pageNumber);
		var text = page.Text;

		foreach (var line in lineSplitter.Split(text))
		{
			var trimmed = line.Trim();
			if (string.IsNullOrWhiteSpace(trimmed))
			{
				continue;
			}

			yield return new RawPdfLine(pageNumber, trimmed);
		}
	}
}

static IEnumerable<TransactionRow> ExtractTransactionsFromPageText(RawPdfLine line)
{
	var cleaned = line.Text
		.Replace("\u00A0", " ", StringComparison.Ordinal)
		.Trim();

	var dateRegex = new Regex(@"\d{4}\.\d{2}\.\d{2}", RegexOptions.Compiled);
	var dateMatches = dateRegex.Matches(cleaned);

	for (var i = 0; i < dateMatches.Count; i++)
	{
		var start = dateMatches[i].Index;
		var end = i + 1 < dateMatches.Count ? dateMatches[i + 1].Index : cleaned.Length;
		var segment = cleaned[start..end].Trim();

		var tx = TryParseTransactionSegment(line.Page, segment);
		if (tx is not null)
		{
			yield return tx;
		}
	}
}

static TransactionRow? TryParseTransactionSegment(int page, string segment)
{
	if (segment.Length < 10)
	{
		return null;
	}

	var dateRaw = segment[..10];
	var date = ParseDate(dateRaw);
	if (date is null)
	{
		return null;
	}

	var amountMatch = Regex.Matches(segment, @"(?<sign>[+-])\s*(?<value>\d[\d\s]*,\d{2})\s*(?<currency>[A-Z]{3})")
		.Cast<Match>()
		.LastOrDefault();

	if (amountMatch is null)
	{
		return null;
	}

	var amountValue = amountMatch.Groups["value"].Value.Replace(" ", string.Empty, StringComparison.Ordinal);
	var amountSign = amountMatch.Groups["sign"].Value;
	var normalizedAmount = $"{amountSign}{amountValue}".Replace(',', '.');

	if (!decimal.TryParse(normalizedAmount, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var amount))
	{
		return null;
	}

	var descriptionPart = segment[10..].Trim();
	var description = RemoveRange(descriptionPart, amountMatch.Index - 10, amountMatch.Length).Trim();
	if (description.Length == 0)
	{
		description = descriptionPart.Trim();
	}

	return new TransactionRow(
		page,
		date.Value,
		description,
		amount,
		null,
		segment);
}

static DateTime? ParseDate(string value)
{
	var formats = new[]
	{
		"yyyy.MM.dd",
		"yyyy-MM-dd",
		"yyyy/MM/dd",
		"dd.MM.yyyy",
		"dd-MM-yyyy",
		"dd/MM/yyyy",
	};

	if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
	{
		return dt;
	}

	if (DateTime.TryParse(value, CultureInfo.GetCultureInfo("hu-HU"), DateTimeStyles.None, out dt))
	{
		return dt;
	}

	return null;
}

static string RemoveRange(string source, int start, int length)
{
	if (start < 0 || length <= 0 || start >= source.Length)
	{
		return source;
	}

	var safeLength = Math.Min(length, source.Length - start);
	return source.Remove(start, safeLength);
}

static void WriteWorkbook(
	string outputPath,
	IReadOnlyList<RawPdfLine> rawLines,
	IReadOnlyList<TransactionRow> transactions,
	string sheetName)
{
	var outputDir = Path.GetDirectoryName(outputPath);
	if (!string.IsNullOrWhiteSpace(outputDir))
	{
		Directory.CreateDirectory(outputDir);
	}

	using var workbook = new XLWorkbook();

	var txSheet = workbook.Worksheets.Add(SanitizeSheetName(sheetName));
	txSheet.Cell(1, 1).Value = "Page";
	txSheet.Cell(1, 2).Value = "Date";
	txSheet.Cell(1, 3).Value = "Description";
	txSheet.Cell(1, 4).Value = "Amount";
	txSheet.Cell(1, 5).Value = "Balance";
	txSheet.Cell(1, 6).Value = "RawLine";

	for (var i = 0; i < transactions.Count; i++)
	{
		var row = i + 2;
		var tx = transactions[i];
		txSheet.Cell(row, 1).Value = tx.Page;
		txSheet.Cell(row, 2).Value = tx.Date;
		txSheet.Cell(row, 3).Value = tx.Description;
		txSheet.Cell(row, 4).Value = tx.Amount;
		txSheet.Cell(row, 5).Value = tx.Balance;
		txSheet.Cell(row, 6).Value = tx.RawLine;
	}

	txSheet.Column(2).Style.DateFormat.Format = "yyyy-MM-dd";
	txSheet.Column(4).Style.NumberFormat.Format = "#,##0.00";
	txSheet.Column(5).Style.NumberFormat.Format = "#,##0.00";
	txSheet.SheetView.FreezeRows(1);
	txSheet.Columns().AdjustToContents();

	var rawSheet = workbook.Worksheets.Add("RawLines");
	rawSheet.Cell(1, 1).Value = "Page";
	rawSheet.Cell(1, 2).Value = "Line";

	for (var i = 0; i < rawLines.Count; i++)
	{
		var row = i + 2;
		rawSheet.Cell(row, 1).Value = rawLines[i].Page;
		rawSheet.Cell(row, 2).Value = rawLines[i].Text;
	}

	rawSheet.SheetView.FreezeRows(1);
	rawSheet.Columns().AdjustToContents();

	workbook.SaveAs(outputPath);
}

static string SanitizeSheetName(string sheetName)
{
	var invalid = new[] { '[', ']', '*', '?', '/', '\\', ':' };
	var sb = new StringBuilder(sheetName.Length);
	foreach (var c in sheetName)
	{
		sb.Append(invalid.Contains(c) ? '_' : c);
	}

	var candidate = sb.ToString().Trim();
	if (candidate.Length == 0)
	{
		candidate = "Transactions";
	}

	return candidate.Length <= 31 ? candidate : candidate[..31];
}

internal sealed record RawPdfLine(int Page, string Text);

internal sealed record TransactionRow(
	int Page,
	DateTime Date,
	string Description,
	decimal Amount,
	decimal? Balance,
	string RawLine);
