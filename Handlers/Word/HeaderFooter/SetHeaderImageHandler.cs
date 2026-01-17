using Aspose.Words.Drawing;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header images in Word documents.
/// </summary>
public class SetHeaderImageHandler : HeaderFooterImageHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "set_header_image";

    /// <inheritdoc />
    protected override bool IsHeader => true;

    /// <inheritdoc />
    protected override string TargetName => "Header";

    /// <inheritdoc />
    protected override RelativeVerticalPosition VerticalPosition => RelativeVerticalPosition.TopMargin;
}
