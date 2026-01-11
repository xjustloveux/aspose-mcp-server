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
        var path = parameters.GetOptional<string?>("path");
        var outputPath = parameters.GetOptional<string?>("outputPath");

        var savePath = path ?? outputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(savePath, allowAbsolutePaths: true);

        using var presentation = new Presentation();
        presentation.Save(savePath, SaveFormat.Pptx);

        return Success($"PowerPoint presentation created successfully. Output: {savePath}");
    }
}
