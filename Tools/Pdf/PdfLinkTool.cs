using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing links in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfLinkTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfLinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfLinkTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PDF link operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="pageIndex">Page index (1-based, required for add, delete, edit).</param>
    /// <param name="linkIndex">Link index (0-based, required for delete, edit).</param>
    /// <param name="x">X position of link area in PDF coordinates (required for add).</param>
    /// <param name="y">Y position of link area in PDF coordinates (required for add).</param>
    /// <param name="width">Width of link area in PDF points (required for add).</param>
    /// <param name="height">Height of link area in PDF points (required for add).</param>
    /// <param name="url">URL to link to (for add, edit).</param>
    /// <param name="targetPage">Target page number (1-based, for add, edit).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_link")]
    [Description(@"Manage links in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add link: pdf_link(operation='add', path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=30, url='https://example.com')
- Delete link: pdf_link(operation='delete', path='doc.pdf', pageIndex=1, linkIndex=0)
- Edit link: pdf_link(operation='edit', path='doc.pdf', pageIndex=1, linkIndex=0, url='https://newurl.com')
- Get links: pdf_link(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a link (required params: path, pageIndex, x, y, width, height, url)
- 'delete': Delete a link (required params: path, pageIndex, linkIndex)
- 'edit': Edit a link (required params: path, pageIndex, linkIndex, url)
- 'get': Get all links (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit)")]
        int pageIndex = 0,
        [Description("Link index (0-based, required for delete, edit)")]
        int linkIndex = 0,
        [Description("X position of link area in PDF coordinates, origin at bottom-left corner (required for add)")]
        double x = 0,
        [Description("Y position of link area in PDF coordinates, origin at bottom-left corner (required for add)")]
        double y = 0,
        [Description("Width of link area in PDF points (required for add)")]
        double width = 0,
        [Description("Height of link area in PDF points (required for add)")]
        double height = 0,
        [Description("URL to link to (for add, edit)")]
        string? url = null,
        [Description("Target page number (1-based, for add, edit)")]
        int? targetPage = null)
    {
        return operation.ToLower() switch
        {
            "add" => AddLink(sessionId, path, outputPath, pageIndex, x, y, width, height, url, targetPage),
            "delete" => DeleteLink(sessionId, path, outputPath, pageIndex, linkIndex),
            "edit" => EditLink(sessionId, path, outputPath, pageIndex, linkIndex, url, targetPage),
            "get" => GetLinks(sessionId, path, pageIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a link annotation to the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="x">The X position of the link area.</param>
    /// <param name="y">The Y position of the link area.</param>
    /// <param name="width">The width of the link area.</param>
    /// <param name="height">The height of the link area.</param>
    /// <param name="url">Optional URL for an external link.</param>
    /// <param name="targetPage">Optional target page number for an internal link.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string AddLink(string? sessionId, string? path, string? outputPath, int pageIndex, double x, double y,
        double width, double height, string? url, int? targetPage)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);
        LinkAnnotation link;

        if (!string.IsNullOrEmpty(url))
        {
            link = new LinkAnnotation(page, rect) { Action = new GoToURIAction(url) };
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            link = new LinkAnnotation(page, rect) { Action = new GoToAction(document.Pages[targetPage.Value]) };
        }
        else
        {
            throw new ArgumentException("Either url or targetPage must be provided");
        }

        page.Annotations.Add(link);

        ctx.Save(outputPath);

        return $"Added link to page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a link annotation from the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="linkIndex">The 0-based link index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page or link index is invalid.</exception>
    private string DeleteLink(string? sessionId, string? path, string? outputPath, int pageIndex, int linkIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();
        if (linkIndex < 0 || linkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        page.Annotations.Delete(links[linkIndex]);

        ctx.Save(outputPath);

        return $"Deleted link {linkIndex} from page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing link annotation on the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="linkIndex">The 0-based link index.</param>
    /// <param name="url">Optional new URL for an external link.</param>
    /// <param name="targetPage">Optional new target page number for an internal link.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page or link index is invalid.</exception>
    private string EditLink(string? sessionId, string? path, string? outputPath, int pageIndex, int linkIndex,
        string? url, int? targetPage)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();
        if (linkIndex < 0 || linkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        var link = links[linkIndex];
        if (!string.IsNullOrEmpty(url))
        {
            link.Action = new GoToURIAction(url);
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            link.Action = new GoToAction(document.Pages[targetPage.Value]);
        }

        ctx.Save(outputPath);

        return $"Edited link {linkIndex} on page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves all link annotations from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="pageIndex">Optional 1-based page index to get links from a specific page.</param>
    /// <returns>A JSON string containing link information.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    private string GetLinks(string? sessionId, string? path, int? pageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        List<object> linkList = [];

        if (pageIndex is > 0)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex.Value];
            var links = page.Annotations.OfType<LinkAnnotation>().ToList();

            if (links.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    pageIndex = pageIndex.Value,
                    items = Array.Empty<object>(),
                    message = $"No links found on page {pageIndex.Value}"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            for (var i = 0; i < links.Count; i++)
            {
                var link = links[i];
                var linkInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["pageIndex"] = pageIndex.Value,
                    ["x"] = link.Rect.LLX,
                    ["y"] = link.Rect.LLY
                };
                if (link.Action is GoToURIAction uriAction)
                {
                    linkInfo["type"] = "url";
                    linkInfo["url"] = uriAction.URI;
                }
                else if (link.Action is GoToAction gotoAction)
                {
                    linkInfo["type"] = "page";
                    if (gotoAction.Destination is XYZExplicitDestination xyzDest)
                        linkInfo["destinationPage"] = xyzDest.PageNumber;
                    else if (gotoAction.Destination is ExplicitDestination explicitDest)
                        linkInfo["destinationPage"] = explicitDest.PageNumber;
                }

                linkList.Add(linkInfo);
            }

            var result = new
            {
                count = linkList.Count,
                pageIndex = pageIndex.Value,
                items = linkList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            for (var p = 1; p <= document.Pages.Count; p++)
            {
                var page = document.Pages[p];
                var links = page.Annotations.OfType<LinkAnnotation>().ToList();
                for (var i = 0; i < links.Count; i++)
                {
                    var link = links[i];
                    var linkInfo = new Dictionary<string, object?>
                    {
                        ["index"] = i,
                        ["pageIndex"] = p,
                        ["x"] = link.Rect.LLX,
                        ["y"] = link.Rect.LLY
                    };
                    if (link.Action is GoToURIAction uriAction)
                    {
                        linkInfo["type"] = "url";
                        linkInfo["url"] = uriAction.URI;
                    }
                    else if (link.Action is GoToAction gotoAction)
                    {
                        linkInfo["type"] = "page";
                        if (gotoAction.Destination is XYZExplicitDestination xyzDest)
                            linkInfo["destinationPage"] = xyzDest.PageNumber;
                        else if (gotoAction.Destination is ExplicitDestination explicitDest)
                            linkInfo["destinationPage"] = explicitDest.PageNumber;
                    }

                    linkList.Add(linkInfo);
                }
            }

            if (linkList.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No links found in document"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var result = new
            {
                count = linkList.Count,
                items = linkList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}