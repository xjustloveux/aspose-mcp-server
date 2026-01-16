using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for creating a new PowerPoint presentation.
/// </summary>
public class CreatePresentationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new PowerPoint presentation.
    /// </summary>
    /// <param name="context">The presentation context (not used for create).</param>
    /// <param name="parameters">
    ///     Required: path or outputPath
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractCreateParameters(parameters);

        var savePath = p.Path ?? p.OutputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(savePath, allowAbsolutePaths: true);

        using var presentation = new Presentation();
        presentation.Save(savePath, SaveFormat.Pptx);

        return Success($"PowerPoint presentation created successfully. Output: {savePath}");
    }

    /// <summary>
    ///     Extracts create parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted create parameters.</returns>
    private static CreateParameters ExtractCreateParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("outputPath"));
    }

    /// <summary>
    ///     Record for holding create presentation parameters.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path.</param>
    private record CreateParameters(string? Path, string? OutputPath);
}
