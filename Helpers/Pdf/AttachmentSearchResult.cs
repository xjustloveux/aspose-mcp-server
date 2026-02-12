namespace AsposeMcpServer.Helpers.Pdf;

/// <summary>
///     Represents the result of searching for an attachment in a PDF document.
/// </summary>
/// <param name="Found">Whether the attachment was found.</param>
/// <param name="ActualName">The actual name of the found attachment, or empty if not found.</param>
/// <param name="AllNames">List of all attachment names in the collection.</param>
public record AttachmentSearchResult(bool Found, string ActualName, List<string> AllNames);
