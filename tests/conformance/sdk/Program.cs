using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;

if (args.Length != 1) throw new ArgumentException("Pass the exported QA corpus directory");
var files = Directory.GetFiles(args[0], "*.output.xlsx").Order().ToArray();
if (files.Length == 0) throw new InvalidOperationException("No corpus outputs found; refusing a vacuous pass");
using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(args[0], "manifest.json")));
var expected = manifest.RootElement.EnumerateArray().Select(c => c.GetProperty("id").GetString() + ".output.xlsx").Order().ToArray();
if (!files.Select(Path.GetFileName).SequenceEqual(expected)) throw new InvalidOperationException("Output set does not match corpus manifest");
var validator = new OpenXmlValidator(FileFormatVersions.Office2016) { MaxNumberOfErrors = 0 };
var results = new List<object>();
var failed = false;
foreach (var file in files) {
    using var document = SpreadsheetDocument.Open(file, false);
    var errors = validator.Validate(document).Select(e => new {
        e.Id, e.Description, Part = e.Part?.Uri.ToString(), Path = e.Path?.XPath
    }).ToArray();
    results.Add(new { File = Path.GetFileName(file), Errors = errors });
    failed |= errors.Length != 0;
}
// Calibrate the independent validator too: a known bad element must be rejected.
using (var document = SpreadsheetDocument.Open(files[0], false)) {
    var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet ?? throw new InvalidOperationException("Missing worksheet");
    worksheet.AppendChild(new OpenXmlUnknownElement("audit", "invalid", "urn:invalid"));
    if (!validator.Validate(document).Any()) throw new InvalidOperationException("SDK calibration failed: accepted invalid worksheet");
}
File.WriteAllText(Path.Combine(args[0], "sdk-results.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Open XML SDK: {files.Length} output packages checked; calibration passed");
return failed ? 1 : 0;
