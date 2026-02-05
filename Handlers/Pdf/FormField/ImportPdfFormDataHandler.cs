using Aspose.Pdf;
using Aspose.Pdf.Facades;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for importing form data into a PDF document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ImportPdfFormDataHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "import";

    /// <summary>
    ///     Imports form data from an FDF, XFDF, or XML file into the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: dataPath (input data file path)
    ///     Optional: format (fdf, xfdf, xml; default: auto-detect from extension)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when dataPath is missing or format is unknown.</exception>
    /// <exception cref="FileNotFoundException">Thrown when data file does not exist.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractImportParameters(parameters);

        SecurityHelper.ValidateFilePath(p.DataPath, "dataPath", true);
        if (!File.Exists(p.DataPath))
            throw new FileNotFoundException($"Data file not found: {p.DataPath}");

        var format = p.Format ?? DetectFormatFromExtension(p.DataPath);

        using var form = new Form(context.Document);
        using var stream = new FileStream(p.DataPath, FileMode.Open, FileAccess.Read);

        switch (format.ToLowerInvariant())
        {
            case "fdf":
                form.ImportFdf(stream);
                break;
            case "xfdf":
                form.ImportXfdf(stream);
                break;
            case "xml":
                form.ImportXml(stream);
                break;
            default:
                throw new ArgumentException($"Unknown import format: {format}. Supported: fdf, xfdf, xml");
        }

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Form data imported from {format.ToUpperInvariant()} file."
        };
    }

    /// <summary>
    ///     Detects the data format from the file extension.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <returns>The detected format string.</returns>
    private static string DetectFormatFromExtension(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".fdf" => "fdf",
            ".xfdf" => "xfdf",
            ".xml" => "xml",
            _ => "xfdf"
        };
    }

    /// <summary>
    ///     Extracts import parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted import parameters.</returns>
    private static ImportParameters ExtractImportParameters(OperationParameters parameters)
    {
        return new ImportParameters(
            parameters.GetRequired<string>("dataPath"),
            parameters.GetOptional<string?>("format")
        );
    }

    /// <summary>
    ///     Parameters for import operation.
    /// </summary>
    /// <param name="DataPath">The input data file path.</param>
    /// <param name="Format">The import format (fdf, xfdf, xml), or null for auto-detect.</param>
    private sealed record ImportParameters(string DataPath, string? Format);
}
