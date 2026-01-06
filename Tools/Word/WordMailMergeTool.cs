using System.ComponentModel;
using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.MailMerging;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for performing mail merge operations on Word document templates.
/// </summary>
[McpServerToolType]
public class WordMailMergeTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordMailMergeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordMailMergeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Performs mail merge on a Word document template.
    /// </summary>
    /// <param name="templatePath">Template file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID to use template from session.</param>
    /// <param name="outputPath">Output file path (required).</param>
    /// <param name="data">Key-value pairs for mail merge fields (for single record), as JSON object.</param>
    /// <param name="dataArray">Array of objects for multiple records, as JSON array.</param>
    /// <param name="cleanupOptions">Cleanup options to apply after mail merge, as comma-separated string.</param>
    /// <returns>A message indicating the mail merge result with field and file information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither templatePath nor sessionId is provided,
    ///     neither data nor dataArray is provided, or both data and dataArray are provided.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled or document cloning fails.</exception>
    [McpServerTool(Name = "word_mail_merge")]
    [Description(@"Perform mail merge on a Word document template.

Usage examples:
- Single record: word_mail_merge(templatePath='template.docx', outputPath='output.docx', data={'name':'John','address':'123 Main St'})
- Multiple records: word_mail_merge(templatePath='template.docx', outputPath='output.docx', dataArray=[{'name':'John'},{'name':'Jane'}])
- From session: word_mail_merge(sessionId='sess_xxx', outputPath='output.docx', data={'name':'John'})")]
    public string Execute(
        [Description("Template file path (required if no sessionId)")]
        string? templatePath = null,
        [Description("Session ID to use template from session")]
        string? sessionId = null,
        [Description(
            "Output file path (required). For multiple records, files will be named output_1.docx, output_2.docx, etc.")]
        string? outputPath = null,
        [Description("Key-value pairs for mail merge fields (for single record), as JSON object")]
        string? data = null,
        [Description(
            "Array of objects for multiple records, as JSON array. Each object contains key-value pairs for mail merge fields. Example: [{'name':'John','city':'NYC'},{'name':'Jane','city':'LA'}]")]
        string? dataArray = null,
        [Description(@"Cleanup options to apply after mail merge, as comma-separated string. Available options:
- 'removeUnusedFields': Remove merge fields that were not populated
- 'removeUnusedRegions': Remove mail merge regions that were not populated
- 'removeEmptyParagraphs': Remove paragraphs that become empty after merge
- 'removeContainingFields': Remove paragraphs containing empty merge fields
- 'removeStaticFields': Remove static fields (like PAGE, DATE)
Default: 'removeUnusedFields,removeEmptyParagraphs'")]
        string? cleanupOptions = null)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (string.IsNullOrEmpty(templatePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either templatePath or sessionId must be provided");

        if (!string.IsNullOrEmpty(templatePath))
            SecurityHelper.ValidateFilePath(templatePath, "templatePath", true);

        JsonObject? dataObject = null;
        JsonArray? dataArrayObject = null;

        if (!string.IsNullOrEmpty(data)) dataObject = JsonNode.Parse(data) as JsonObject;

        if (!string.IsNullOrEmpty(dataArray)) dataArrayObject = JsonNode.Parse(dataArray) as JsonArray;

        if (dataObject == null && dataArrayObject == null)
            throw new ArgumentException(
                "Either 'data' (for single record) or 'dataArray' (for multiple records) must be provided");

        if (dataObject != null && dataArrayObject != null)
            throw new ArgumentException(
                "Cannot specify both 'data' and 'dataArray'. Use 'data' for single record or 'dataArray' for multiple records");

        var cleanupOptionsFlags = ParseCleanupOptions(cleanupOptions);

        if (dataArrayObject is { Count: > 0 })
            return ExecuteMultipleRecords(templatePath, sessionId, outputPath, dataArrayObject, cleanupOptionsFlags);

        if (dataObject != null)
            return ExecuteSingleRecord(templatePath, sessionId, outputPath, dataObject, cleanupOptionsFlags);

        throw new ArgumentException("No data provided for mail merge");
    }

    /// <summary>
    ///     Executes mail merge for a single record.
    /// </summary>
    /// <param name="templatePath">The template file path.</param>
    /// <param name="sessionId">The session ID for reading template from session.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="data">The JSON object containing field names and values.</param>
    /// <param name="cleanupOptions">The mail merge cleanup options to apply.</param>
    /// <returns>A message indicating the result of the mail merge operation.</returns>
    private string ExecuteSingleRecord(string? templatePath, string? sessionId, string outputPath, JsonObject data,
        MailMergeCleanupOptions cleanupOptions)
    {
        Document doc;
        string templateSource;

        if (!string.IsNullOrEmpty(sessionId))
        {
            using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, null, _identityAccessor);
            doc = ctx.Document.Clone() ?? throw new InvalidOperationException("Failed to clone document from session");
            templateSource = $"session {sessionId}";
        }
        else
        {
            doc = new Document(templatePath);
            templateSource = Path.GetFileName(templatePath!);
        }

        doc.MailMerge.CleanupOptions = cleanupOptions;

        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.Save(outputPath);

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully");
        result.AppendLine($"Template: {templateSource}");
        result.AppendLine($"Output: {outputPath}");
        result.AppendLine($"Fields merged: {fieldNames.Length}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");

        return result.ToString();
    }

    /// <summary>
    ///     Executes mail merge for multiple records.
    /// </summary>
    /// <param name="templatePath">The template file path.</param>
    /// <param name="sessionId">The session ID for reading template from session.</param>
    /// <param name="outputPath">The base output file path (files will be numbered).</param>
    /// <param name="dataArray">The JSON array containing multiple record objects.</param>
    /// <param name="cleanupOptions">The mail merge cleanup options to apply.</param>
    /// <returns>A message indicating the result of the mail merge operation.</returns>
    private string ExecuteMultipleRecords(string? templatePath, string? sessionId, string outputPath,
        JsonArray dataArray,
        MailMergeCleanupOptions cleanupOptions)
    {
        List<string> outputFiles = [];
        var outputDir = Path.GetDirectoryName(outputPath) ?? ".";
        var outputName = Path.GetFileNameWithoutExtension(outputPath);
        var outputExt = Path.GetExtension(outputPath);

        Document? templateDoc = null;
        string templateSource;

        if (!string.IsNullOrEmpty(sessionId))
        {
            using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, null, _identityAccessor);
            templateDoc = ctx.Document;
            templateSource = $"session {sessionId}";
        }
        else
        {
            templateSource = Path.GetFileName(templatePath!);
        }

        for (var i = 0; i < dataArray.Count; i++)
        {
            var recordData = dataArray[i] as JsonObject;
            if (recordData == null) continue;

            Document doc;
            if (templateDoc != null)
                doc = templateDoc.Clone() ??
                      throw new InvalidOperationException("Failed to clone document from session");
            else
                doc = new Document(templatePath);

            doc.MailMerge.CleanupOptions = cleanupOptions;

            var fieldNames = recordData.Select(kvp => kvp.Key).ToArray();
            var fieldValues = recordData.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

            doc.MailMerge.Execute(fieldNames, fieldValues);

            var recordOutputPath = dataArray.Count == 1
                ? outputPath
                : Path.Combine(outputDir, $"{outputName}_{i + 1}{outputExt}");

            doc.Save(recordOutputPath);
            outputFiles.Add(recordOutputPath);
        }

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully (multiple records)");
        result.AppendLine($"Template: {templateSource}");
        result.AppendLine($"Records processed: {outputFiles.Count}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");
        result.AppendLine("Output files:");
        foreach (var file in outputFiles) result.AppendLine($"  - {file}");

        return result.ToString();
    }

    /// <summary>
    ///     Parses cleanup options from comma-separated string.
    /// </summary>
    /// <param name="optionsString">The comma-separated cleanup options string.</param>
    /// <returns>The parsed MailMergeCleanupOptions flags.</returns>
    private static MailMergeCleanupOptions ParseCleanupOptions(string? optionsString)
    {
        if (string.IsNullOrEmpty(optionsString))
            return MailMergeCleanupOptions.RemoveUnusedFields | MailMergeCleanupOptions.RemoveEmptyParagraphs;

        var options = MailMergeCleanupOptions.None;
        var optionsList =
            optionsString.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var option in optionsList)
            options |= option.ToLower() switch
            {
                "removeunusedfields" => MailMergeCleanupOptions.RemoveUnusedFields,
                "removeunusedregions" => MailMergeCleanupOptions.RemoveUnusedRegions,
                "removeemptyparagraphs" => MailMergeCleanupOptions.RemoveEmptyParagraphs,
                "removecontainingfields" => MailMergeCleanupOptions.RemoveContainingFields,
                "removestaticfields" => MailMergeCleanupOptions.RemoveStaticFields,
                _ => MailMergeCleanupOptions.None
            };

        return options;
    }
}