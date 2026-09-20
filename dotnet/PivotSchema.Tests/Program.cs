using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;

var directory = args.Length > 0 ? args[0] : "test-results/pivots";
var reports = new List<object>();
var errors = 0;
var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
FixtureContracts.ValidateFileSet(Directory.GetFiles(directory, "*.xlsx").Select(f => Path.GetFileName(f)!));
foreach (var name in FixtureContracts.Files)
{
    using var document = SpreadsheetDocument.Open(Path.Combine(directory, name), false);
    // Validate every document, including the names-only fixture, against the same schema.
    var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
    var failures = validator.Validate(document).Select(e => new { e.Id, e.Description, Path = e.Path?.XPath, Part = e.Part?.Uri.ToString() }).ToArray();
    var contractFailures = new List<string>();
    var pivots = 0;
    try { pivots = FixtureContracts.Validate(document, name); }
    catch (InvalidDataException e) { contractFailures.Add(e.Message); }
    reports.Add(new { file = name, pivots, failures, contractFailures });
    Console.WriteLine($"{name}: {pivots} pivot(s), {failures.Length} schema error(s), {contractFailures.Count} contract error(s)");
    foreach (var failure in failures) Console.WriteLine(JsonSerializer.Serialize(failure));
    foreach (var failure in contractFailures) Console.WriteLine(failure);
    errors += failures.Length + contractFailures.Count;
}
File.WriteAllText(Path.Combine(directory, "schema-report.json"), JsonSerializer.Serialize(reports, jsonOptions));
// Run only after valid baselines. Negative tests mutate in-memory copies, never output files.
if (errors == 0)
{
    var passed = FixtureContracts.RunNegativeTests(directory);
    File.WriteAllText(Path.Combine(directory, "schema-contract-tests.json"), JsonSerializer.Serialize(new { passed }, jsonOptions));
    Console.WriteLine($"{passed.Length} negative fixture-contract checks passed");
}
return errors == 0 ? 0 : 1;
