using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Info;

/// <summary>
///     Handler for retrieving statistics from PDF documents.
/// </summary>
public class GetPdfStatisticsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_statistics";

    /// <summary>
    ///     Retrieves statistics about the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing document statistics.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;

        var totalAnnotations = 0;
        var totalParagraphs = 0;
        for (var i = 1; i <= document.Pages.Count; i++)
        {
            var page = document.Pages[i];
            totalAnnotations += page.Annotations.Count;
            totalParagraphs += page.Paragraphs.Count;
        }

        if (context.SessionId != null)
            return JsonResult(new
            {
                totalPages = document.Pages.Count,
                isEncrypted = document.IsEncrypted,
                isLinearized = document.IsLinearized,
                bookmarks = document.Outlines.Count,
                formFields = document.Form?.Count ?? 0,
                totalAnnotations,
                totalParagraphs,
                note = "File size info not available in session mode"
            });

        if (string.IsNullOrEmpty(context.SourcePath))
            throw new ArgumentException("path is required for get_statistics operation");

        SecurityHelper.ValidateFilePath(context.SourcePath, "path", true);

        var fileInfo = new FileInfo(context.SourcePath);

        return JsonResult(new
        {
            fileSizeBytes = fileInfo.Length,
            fileSizeKb = Math.Round(fileInfo.Length / 1024.0, 2),
            totalPages = document.Pages.Count,
            isEncrypted = document.IsEncrypted,
            isLinearized = document.IsLinearized,
            bookmarks = document.Outlines.Count,
            formFields = document.Form?.Count ?? 0,
            totalAnnotations,
            totalParagraphs
        });
    }
}
