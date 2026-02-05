using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.ContentControl;

/// <summary>
///     Handler for deleting content controls from a Word document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteWordContentControlHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a content control identified by index or tag.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: index (0-based) or tag (to identify the content control)
    ///     Optional: keepContent (default: true) — whether to keep the content after removing the control
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when the content control cannot be found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var index = parameters.GetOptional<int?>("index");
        var tag = parameters.GetOptional<string?>("tag");
        var keepContent = parameters.GetOptional("keepContent", true);

        var doc = context.Document;
        var sdt = EditWordContentControlHandler.FindContentControl(doc, index, tag);

        var identifier = !string.IsNullOrEmpty(sdt.Tag) ? $"tag='{sdt.Tag}'" : $"index={index}";
        var sdtType = sdt.SdtType.ToString();

        if (keepContent)
            sdt.RemoveSelfOnly();
        else
            sdt.Remove();

        MarkModified(context);

        var contentAction = keepContent ? "Content preserved." : "Content removed.";
        return new SuccessResult
        {
            Message = $"Content control ({sdtType}, {identifier}) deleted. {contentAction}"
        };
    }
}
