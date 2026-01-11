using Aspose.Pdf;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

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
    /// <returns>A tuple indicating whether the attachment was found, its actual name, and all available names.</returns>
    public static (bool found, string actualName, List<string> allNames) FindAttachment(
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
                    return (true, name, allNames);
            }
            catch
            {
                // Ignore errors reading individual attachment
            }

        return (false, "", allNames);
    }

    /// <summary>
    ///     Collects detailed information about all attachments.
    /// </summary>
    /// <param name="embeddedFiles">The collection of embedded files.</param>
    /// <returns>A list of attachment information objects.</returns>
    public static List<object> CollectAttachmentInfo(EmbeddedFileCollection embeddedFiles)
    {
        List<object> attachmentList = [];

        for (var i = 1; i <= embeddedFiles.Count; i++)
            try
            {
                var file = embeddedFiles[i];
                var attachmentInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["name"] = file.Name ?? "(unnamed)",
                    ["description"] = !string.IsNullOrEmpty(file.Description) ? file.Description : null,
                    ["mimeType"] = !string.IsNullOrEmpty(file.MIMEType) ? file.MIMEType : null
                };

                try
                {
                    if (file.Contents != null)
                        attachmentInfo["sizeBytes"] = file.Contents.Length;
                }
                catch
                {
                    attachmentInfo["sizeBytes"] = null;
                }

                attachmentList.Add(attachmentInfo);
            }
            catch (Exception ex)
            {
                attachmentList.Add(new { index = i, error = ex.Message });
            }

        return attachmentList;
    }
}
