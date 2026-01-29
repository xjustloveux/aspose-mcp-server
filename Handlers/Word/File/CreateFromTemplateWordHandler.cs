using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for creating Word documents from templates.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled or clone fails.</exception>
    /// <exception cref="FileNotFoundException">Thrown when template file is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractCreateFromTemplateParameters(parameters);

        if (string.IsNullOrEmpty(p.TemplatePath) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException(
                "Either templatePath or sessionId is required for create_from_template operation");
        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for create_from_template operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        if (string.IsNullOrEmpty(p.DataJson))
            throw new ArgumentException("dataJson parameter is required for create_from_template");

        Document doc;
        string templateSource;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            var sessionDoc = context.SessionManager.GetDocument<Document>(p.SessionId, identity);
            doc = sessionDoc.Clone() ?? throw new InvalidOperationException("Failed to clone document from session");
            templateSource = $"session {p.SessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(p.TemplatePath!, "templatePath", true);
            if (!System.IO.File.Exists(p.TemplatePath))
                throw new FileNotFoundException($"Template file not found: {p.TemplatePath}");
            doc = new Document(p.TemplatePath);
            templateSource = Path.GetFileName(p.TemplatePath);
        }

        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs
        };

        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(p.DataJson));
        var loadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true,
            ExactDateTimeParseFormats = ["yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss"],
            SimpleValueParseMode = JsonSimpleValueParseMode.Strict
        };
        var dataSource = new JsonDataSource(jsonStream, loadOptions);

        engine.BuildReport(doc, dataSource, "ds");

        doc.Save(p.OutputPath);
        return new SuccessResult
        {
            Message = $"Document created from template ({templateSource}) using LINQ Reporting Engine: {p.OutputPath}"
        };
    }

    private static CreateFromTemplateParameters ExtractCreateFromTemplateParameters(OperationParameters parameters)
    {
        return new CreateFromTemplateParameters(
            parameters.GetOptional<string?>("templatePath"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("dataJson"));
    }

    private sealed record CreateFromTemplateParameters(
        string? TemplatePath,
        string? SessionId,
        string? OutputPath,
        string? DataJson);
}
