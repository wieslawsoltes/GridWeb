using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;
var directory = args.Length > 0 ? args[0] : "test-results/pivots";
var reports = new List<object>();
var errors = 0;
foreach (var file in Directory.GetFiles(directory, "*.xlsx"))
{
    using var document = SpreadsheetDocument.Open(file, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
    var failures = validator.Validate(document).Select(e => new { e.Id, e.Description, Path = e.Path?.XPath, Part = e.Part?.Uri.ToString() }).ToArray();
    var pivots = document.WorkbookPart!.WorksheetParts.SelectMany(p => p.PivotTableParts).ToArray();
    if (pivots.Length == 0 || pivots.Any(p => p.PivotTableCacheDefinitionPart?.PivotTableCacheRecordsPart == null)) throw new InvalidDataException("Missing native pivot/cache relationships");
    reports.Add(new { file = Path.GetFileName(file), pivots = pivots.Length, failures });
    Console.WriteLine($"{Path.GetFileName(file)}: {pivots.Length} pivot(s), {failures.Length} schema error(s)");
    foreach (var failure in failures) Console.WriteLine(JsonSerializer.Serialize(failure));
    errors += failures.Length;
}
if (reports.Count != 4) throw new InvalidDataException("Expected four independently validated pivot fixtures");
File.WriteAllText(Path.Combine(directory, "schema-report.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
return errors == 0 ? 0 : 1;
