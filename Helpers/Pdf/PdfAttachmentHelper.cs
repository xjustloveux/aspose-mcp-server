using Aspose.Pdf;
using AsposeMcpServer.Results.Pdf.Attachment;

namespace AsposeMcpServer.Helpers.Pdf;

/// <summary>
///     Helper class for PDF attachment operations.
/// </summary>
public static class PdfAttachmentHelper
{
    /// <summary>
    ///     Collects all attachment names from the embedded files collection.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <returns>A list of attachment names.</returns>
    public static List<string> CollectAttachmentNames(EmbeddedFileCollection embeddedFiles)
    {
        List<string> names = [];
        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                names.Add(file.Name ?? "");
            }
            catch
            {
                // Ignore errors reading individual attachment names
            }

        return names;
    }

    /// <summary>
    ///     Finds an attachment by name in the embedded files collection.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <param name="attachmentName">The name of the attachment to find.</param>
    /// <returns>An <see cref="AttachmentSearchResult" /> with search results.</returns>
    public static AttachmentSearchResult FindAttachment(
        EmbeddedFileCollection embeddedFiles, string attachmentName)
    {
        List<string> allNames = [];

        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                var name = file.Name ?? "";
                allNames.Add(name);

                var fileName = Path.GetFileName(name);
                if (string.Equals(name, attachmentName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(fileName, attachmentName, StringComparison.OrdinalIgnoreCase))
                    return new AttachmentSearchResult(true, name, allNames);
            }
            catch
            {
                // Ignore errors reading individual attachment
            }

        return new AttachmentSearchResult(false, "", allNames);
    }

    /// <summary>
    ///     Collects detailed information about all attachments.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <returns>A list of attachment information objects.</returns>
    public static List<AttachmentInfo> CollectAttachmentInfo(EmbeddedFileCollection embeddedFiles)
    {
        List<AttachmentInfo> attachmentList = [];

        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                long? size = null;

                try
                {
                    if (file.Contents != null)
                        size = file.Contents.Length;
                }
                catch
                {
                    // Ignore errors reading file size
                }

                attachmentList.Add(new AttachmentInfo
                {
                    Index = i,
                    Name = file.Name ?? "(unnamed)",
                    Description = !string.IsNullOrEmpty(file.Description) ? file.Description : null,
                    MimeType = !string.IsNullOrEmpty(file.MIMEType) ? file.MIMEType : null,
                    Size = size
                });
            }
            catch
            {
                // Skip attachments that cannot be read
            }

        return attachmentList;
    }
}
