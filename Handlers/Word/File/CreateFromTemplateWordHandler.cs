using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for creating Word documents from templates.
/// </summary>
public class CreateFromTemplateWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "create_from_template";

    /// <summary>
    ///     Creates a document from a template using LINQ Reporting Engine.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, dataJson, either templatePath or sessionId
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var templatePath = parameters.GetOptional<string?>("templatePath");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var dataJson = parameters.GetOptional<string?>("dataJson");

        if (string.IsNullOrEmpty(templatePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException(
                "Either templatePath or sessionId is required for create_from_template operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for create_from_template operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        if (string.IsNullOrEmpty(dataJson))
            throw new ArgumentException("dataJson parameter is required for create_from_template");

        Document doc;
        string templateSource;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            var sessionDoc = context.SessionManager.GetDocument<Document>(sessionId, identity);
            doc = sessionDoc.Clone() ?? throw new InvalidOperationException("Failed to clone document from session");
            templateSource = $"session {sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(templatePath!, "templatePath", true);
            if (!System.IO.File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            doc = new Document(templatePath);
            templateSource = Path.GetFileName(templatePath);
        }

        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs
        };

        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(dataJson));
        var loadOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = ["yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss"],
            SimpleValueParseMode = JsonSimpleValueParseMode.Strict
        };
        var dataSource = new JsonDataSource(jsonStream, loadOptions);

        engine.BuildReport(doc, dataSource, "ds");

        doc.Save(outputPath);
        return $"Document created from template ({templateSource}) using LINQ Reporting Engine: {outputPath}";
    }
}
