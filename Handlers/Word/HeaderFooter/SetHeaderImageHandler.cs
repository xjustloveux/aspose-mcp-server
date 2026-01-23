using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header images in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
