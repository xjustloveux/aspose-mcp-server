using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.MailMerging;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for performing mail merge operations on Word document templates
/// </summary>
public class WordMailMergeTool : IAsposeTool
{
    public string Description => @"Perform mail merge on a Word document template.

Usage examples:
- Single record: word_mail_merge(templatePath='template.docx', outputPath='output.docx', data={'name':'John','address':'123 Main St'})
- Multiple records: word_mail_merge(templatePath='template.docx', outputPath='output.docx', dataArray=[{'name':'John'},{'name':'Jane'}])";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            templatePath = new
            {
                type = "string",
                description = "Template file path (required)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (required). For multiple records, files will be named output_1.docx, output_2.docx, etc."
            },
            data = new
            {
                type = "object",
                description = "Key-value pairs for mail merge fields (for single record)"
            },
            dataArray = new
            {
                type = "array",
                description =
                    "Array of objects for multiple records. Each object contains key-value pairs for mail merge fields. Example: [{'name':'John','city':'NYC'},{'name':'Jane','city':'LA'}]",
                items = new { type = "object" }
            },
            cleanupOptions = new
            {
                type = "array",
                description = @"Cleanup options to apply after mail merge. Available options:
- 'removeUnusedFields': Remove merge fields that were not populated
- 'removeUnusedRegions': Remove mail merge regions that were not populated
- 'removeEmptyParagraphs': Remove paragraphs that become empty after merge
- 'removeContainingFields': Remove paragraphs containing empty merge fields
- 'removeStaticFields': Remove static fields (like PAGE, DATE)
Default: ['removeUnusedFields', 'removeEmptyParagraphs']",
                items = new { type = "string" }
            }
        },
        required = new[] { "templatePath", "outputPath" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            try
            {
                var templatePath = ArgumentHelper.GetString(arguments, "templatePath");
                var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
                var data = ArgumentHelper.GetObject(arguments, "data", false);
                var dataArray = ArgumentHelper.GetArray(arguments, "dataArray", false);
                var cleanupOptionsArray = ArgumentHelper.GetArray(arguments, "cleanupOptions", false);

                SecurityHelper.ValidateFilePath(templatePath, "templatePath", true);
                SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

                // Validate that either data or dataArray is provided
                if (data == null && dataArray == null)
                    throw new ArgumentException(
                        "Either 'data' (for single record) or 'dataArray' (for multiple records) must be provided");

                if (data != null && dataArray != null)
                    throw new ArgumentException(
                        "Cannot specify both 'data' and 'dataArray'. Use 'data' for single record or 'dataArray' for multiple records");

                // Parse cleanup options
                var cleanupOptions = ParseCleanupOptions(cleanupOptionsArray);

                if (dataArray is { Count: > 0 })
                    // Multiple records mode
                    return ExecuteMultipleRecords(templatePath, outputPath, dataArray, cleanupOptions);

                if (data != null)
                    // Single record mode
                    return ExecuteSingleRecord(templatePath, outputPath, data, cleanupOptions);

                throw new ArgumentException("No data provided for mail merge");
            }
            catch (FileNotFoundException ex)
            {
                return $"Error: Template file not found - {ex.Message}";
            }
            catch (IOException ex)
            {
                return $"Error: File access error - {ex.Message}. The file may be in use by another application.";
            }
            catch (ArgumentException ex)
            {
                return $"Error: Invalid argument - {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: Mail merge failed - {ex.Message}";
            }
        });
    }

    /// <summary>
    ///     Executes mail merge for a single record
    /// </summary>
    /// <param name="templatePath">Template file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="data">Key-value pairs for mail merge fields</param>
    /// <param name="cleanupOptions">Cleanup options to apply after merge</param>
    /// <returns>Success message</returns>
    private string ExecuteSingleRecord(string templatePath, string outputPath, JsonObject data,
        MailMergeCleanupOptions cleanupOptions)
    {
        var doc = new Document(templatePath) { MailMerge = { CleanupOptions = cleanupOptions } };

        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.Save(outputPath);

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully");
        result.AppendLine($"Template: {Path.GetFileName(templatePath)}");
        result.AppendLine($"Output: {outputPath}");
        result.AppendLine($"Fields merged: {fieldNames.Length}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");

        return result.ToString();
    }

    /// <summary>
    ///     Executes mail merge for multiple records
    /// </summary>
    /// <param name="templatePath">Template file path</param>
    /// <param name="outputPath">Base output file path (will be suffixed with _1, _2, etc.)</param>
    /// <param name="dataArray">Array of objects containing key-value pairs for each record</param>
    /// <param name="cleanupOptions">Cleanup options to apply after merge</param>
    /// <returns>Success message with list of output files</returns>
    private string ExecuteMultipleRecords(string templatePath, string outputPath, JsonArray dataArray,
        MailMergeCleanupOptions cleanupOptions)
    {
        var outputFiles = new List<string>();
        var outputDir = Path.GetDirectoryName(outputPath) ?? ".";
        var outputName = Path.GetFileNameWithoutExtension(outputPath);
        var outputExt = Path.GetExtension(outputPath);

        for (var i = 0; i < dataArray.Count; i++)
        {
            var recordData = dataArray[i] as JsonObject;
            if (recordData == null) continue;

            var doc = new Document(templatePath) { MailMerge = { CleanupOptions = cleanupOptions } };

            var fieldNames = recordData.Select(kvp => kvp.Key).ToArray();
            var fieldValues = recordData.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Generate output filename for this record
            var recordOutputPath = dataArray.Count == 1
                ? outputPath
                : Path.Combine(outputDir, $"{outputName}_{i + 1}{outputExt}");

            doc.Save(recordOutputPath);
            outputFiles.Add(recordOutputPath);
        }

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully (multiple records)");
        result.AppendLine($"Template: {Path.GetFileName(templatePath)}");
        result.AppendLine($"Records processed: {outputFiles.Count}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");
        result.AppendLine("Output files:");
        foreach (var file in outputFiles) result.AppendLine($"  - {file}");

        return result.ToString();
    }

    /// <summary>
    ///     Parses cleanup options from JSON array
    /// </summary>
    /// <param name="optionsArray">Array of cleanup option strings</param>
    /// <returns>Combined MailMergeCleanupOptions flags</returns>
    private MailMergeCleanupOptions ParseCleanupOptions(JsonArray? optionsArray)
    {
        // Default cleanup options
        if (optionsArray == null || optionsArray.Count == 0)
            return MailMergeCleanupOptions.RemoveUnusedFields | MailMergeCleanupOptions.RemoveEmptyParagraphs;

        var options = MailMergeCleanupOptions.None;

        foreach (var option in optionsArray)
        {
            var optionStr = option?.ToString().ToLower();
            options |= optionStr switch
            {
                "removeunusedfields" => MailMergeCleanupOptions.RemoveUnusedFields,
                "removeunusedregions" => MailMergeCleanupOptions.RemoveUnusedRegions,
                "removeemptyparagraphs" => MailMergeCleanupOptions.RemoveEmptyParagraphs,
                "removecontainingfields" => MailMergeCleanupOptions.RemoveContainingFields,
                "removestaticfields" => MailMergeCleanupOptions.RemoveStaticFields,
                _ => MailMergeCleanupOptions.None
            };
        }

        return options;
    }
}