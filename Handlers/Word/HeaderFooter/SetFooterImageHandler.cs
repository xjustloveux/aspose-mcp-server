using Aspose.Words.Drawing;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting footer images in Word documents.
/// </summary>
public class SetFooterImageHandler : HeaderFooterImageHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "set_footer_image";

    /// <inheritdoc />
    protected override bool IsHeader => false;

    /// <inheritdoc />
    protected override string TargetName => "Footer";

    /// <inheritdoc />
    protected override RelativeVerticalPosition VerticalPosition => RelativeVerticalPosition.BottomMargin;
}
