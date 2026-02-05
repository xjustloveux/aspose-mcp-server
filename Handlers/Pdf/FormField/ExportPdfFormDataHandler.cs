using Aspose.Pdf;
using Aspose.Pdf.Facades;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Pdf.FormField;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for exporting form data from a PDF document.
/// </summary>
[ResultType(typeof(ExportFormDataResult))]
public class ExportPdfFormDataHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "export";

    /// <summary>
    ///     Exports form data from a PDF document to FDF, XFDF, or XML format.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: dataPath (output file path)
    ///     Optional: format (fdf, xfdf, xml; default: xfdf)
    /// </param>
    /// <returns>Export result with file path and format info.</returns>
    /// <exception cref="ArgumentException">Thrown when dataPath is missing or format is unknown.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractExportParameters(parameters);

        SecurityHelper.ValidateFilePath(p.DataPath, "dataPath", true);

        using var form = new Form(context.Document);

        using var stream = new FileStream(p.DataPath, FileMode.Create);
        switch (p.Format.ToLowerInvariant())
        {
            case "fdf":
                form.ExportFdf(stream);
                break;
            case "xfdf":
                form.ExportXfdf(stream);
                break;
            case "xml":
                form.ExportXml(stream);
                break;
            default:
                throw new ArgumentException($"Unknown export format: {p.Format}. Supported: fdf, xfdf, xml");
        }

        return new ExportFormDataResult
        {
            Message = $"Form data exported to {p.Format.ToUpperInvariant()} format.",
            Format = p.Format.ToLowerInvariant(),
            ExportPath = p.DataPath
        };
    }

    /// <summary>
    ///     Extracts export parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted export parameters.</returns>
    private static ExportParameters ExtractExportParameters(OperationParameters parameters)
    {
        return new ExportParameters(
            parameters.GetRequired<string>("dataPath"),
            parameters.GetOptional("format", "xfdf")
        );
    }

    /// <summary>
    ///     Parameters for export operation.
    /// </summary>
    /// <param name="DataPath">The output file path for exported data.</param>
    /// <param name="Format">The export format (fdf, xfdf, xml).</param>
    private sealed record ExportParameters(string DataPath, string Format);
}
