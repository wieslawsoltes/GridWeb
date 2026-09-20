using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary>Independent semantic checks for every generated OOXML fixture.</summary>
internal static class FixtureContracts
{
    internal static readonly string[] PivotFiles = ["filtered.xlsx", "multi.xlsx", "simple.xlsx", "totals.xlsx"];
    internal const string NamesFile = "scoped-names.xlsx";
    internal static readonly string[] Files = [.. PivotFiles, NamesFile];

    internal static void ValidateFileSet(IEnumerable<string> names)
    {
        if (!names.Order(StringComparer.Ordinal).SequenceEqual(Files.Order(StringComparer.Ordinal)))
            throw new InvalidDataException("Expected exactly four pivot fixtures and scoped-names.xlsx; missing or unexpected fixture files");
    }

    internal static int Validate(SpreadsheetDocument document, string file)
    {
        var part = document.WorkbookPart ?? throw new InvalidDataException("Missing workbook part");
        var pivots = part.WorksheetParts.SelectMany(p => p.PivotTableParts).ToArray();
        if (PivotFiles.Contains(file, StringComparer.Ordinal))
        {
            if (pivots.Length != 1 || pivots[0].PivotTableCacheDefinitionPart?.PivotTableCacheRecordsPart == null)
                throw new InvalidDataException("Expected one native pivot with complete cache relationships");
            var expectedName = "Pivot_" + Path.GetFileNameWithoutExtension(file);
            if (pivots[0].PivotTableDefinition.Name?.Value != expectedName)
                throw new InvalidDataException("Wrong pivot definition for " + file);
            return pivots.Length;
        }
        if (file != NamesFile) throw new InvalidDataException("Unknown fixture " + file);
        if (pivots.Length != 0) throw new InvalidDataException("The scoped-name fixture must not acquire pivot parts");
        var sheets = part.Workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToArray() ?? [];
        if (!sheets.Select(s => s.Name?.Value).SequenceEqual(new[] { "North", "South" }))
            throw new InvalidDataException("Scoped-name fixture worksheet order changed");
        var names = part.Workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>().ToArray() ?? [];
        if (names.Length != 6) throw new InvalidDataException("Expected all six independent scoped definitions");
        CheckName(names, "Rate", null, "2", false, "Workbook rate");
        CheckName(names, "NorthSum", null, "SUM('North'!AMOUNTS)", false, "");
        CheckName(names, "Rate", 0, "3", true, "Local & private");
        CheckName(names, "Amounts", 0, "'North'!A1:A3", false, "");
        CheckName(names, "Rate", 1, "4", false, "");
        CheckName(names, "Amounts", 1, "'South'!A1:A3", false, "");
        return 0;
    }

    private static void CheckName(DefinedName[] names, string name, uint? localSheetId, string formula, bool hidden, string comment)
    {
        // Scope is part of identity: accepting only a count would miss local-to-global flattening.
        var matches = names.Where(n => n.Name?.Value == name && n.LocalSheetId?.Value == localSheetId).ToArray();
        if (matches.Length != 1 || matches[0].Text != formula ||
            (matches[0].Hidden?.Value ?? false) != hidden || (matches[0].Comment?.Value ?? "") != comment)
            throw new InvalidDataException($"Lost scope, formula or metadata for {localSheetId?.ToString() ?? "Workbook"}!{name}");
    }

    /// <summary>Prove that loosening fixture classification never masks missing pivots or names.</summary>
    internal static string[] RunNegativeTests(string directory)
    {
        var passed = new List<string>();
        void Reject(string name, Action action)
        {
            try { action(); }
            catch (InvalidDataException) { passed.Add(name); return; }
            throw new InvalidOperationException("Fixture contract accepted corrupt input: " + name);
        }
        void Mutation(string name, string file, Action<SpreadsheetDocument> mutate)
        {
            using var stream = new MemoryStream();
            stream.Write(File.ReadAllBytes(Path.Combine(directory, file)));
            stream.Position = 0;
            using var document = SpreadsheetDocument.Open(stream, true);
            Validate(document, file); // A broken baseline cannot make a negative test pass.
            mutate(document);
            Reject(name, () => Validate(document, file));
        }
        DefinedName LocalRate(SpreadsheetDocument d) => d.WorkbookPart!.Workbook
            .GetFirstChild<DefinedNames>()!.Elements<DefinedName>()
            .Single(n => n.Name?.Value == "Rate" && n.LocalSheetId?.Value == 0);

        Reject("missing names fixture", () => ValidateFileSet(PivotFiles));
        Reject("missing pivot fixture", () => ValidateFileSet(Files.Where(f => f != "totals.xlsx")));
        Reject("unexpected fixture", () => ValidateFileSet([.. Files, "unexpected.xlsx"]));
        foreach (var file in PivotFiles)
            Mutation("missing pivot in " + file, file, d =>
            {
                var sheet = d.WorkbookPart!.WorksheetParts.Single(p => p.PivotTableParts.Any());
                sheet.DeletePart(sheet.PivotTableParts.Single());
            });
        Mutation("missing pivot cache records", "simple.xlsx", d =>
        {
            var cache = d.WorkbookPart!.WorksheetParts.SelectMany(p => p.PivotTableParts)
                .Single().PivotTableCacheDefinitionPart!;
            cache.DeletePart(cache.PivotTableCacheRecordsPart!);
        });
        Mutation("local name flattened to workbook scope", NamesFile, d => LocalRate(d).LocalSheetId = null);
        Mutation("local name moved to another sheet", NamesFile, d => LocalRate(d).LocalSheetId = 1U);
        Mutation("lost hidden state", NamesFile, d => LocalRate(d).Hidden = false);
        Mutation("lost name comment", NamesFile, d => LocalRate(d).Comment = null);
        Mutation("changed name formula", NamesFile, d => LocalRate(d).Text = "99");
        Mutation("missing name definition", NamesFile, d => LocalRate(d).Remove());
        return passed.ToArray();
    }
}
