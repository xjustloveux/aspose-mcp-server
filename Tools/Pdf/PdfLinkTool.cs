using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing links in PDF documents (add, delete, edit, get)
/// </summary>
public class PdfLinkTool : IAsposeTool
{
    public string Description => @"Manage links in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add link: pdf_link(operation='add', path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=30, url='https://example.com')
- Delete link: pdf_link(operation='delete', path='doc.pdf', pageIndex=1, linkIndex=0)
- Edit link: pdf_link(operation='edit', path='doc.pdf', pageIndex=1, linkIndex=0, url='https://newurl.com')
- Get links: pdf_link(operation='get', path='doc.pdf', pageIndex=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a link (required params: path, pageIndex, x, y, width, height, url)
- 'delete': Delete a link (required params: path, pageIndex, linkIndex)
- 'edit': Edit a link (required params: path, pageIndex, linkIndex, url)
- 'get': Get all links (required params: path, pageIndex)",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, delete, edit)"
            },
            linkIndex = new
            {
                type = "number",
                description = "Link index (0-based, required for delete, edit)"
            },
            x = new
            {
                type = "number",
                description =
                    "X position of link area in PDF coordinates, origin at bottom-left corner (required for add)"
            },
            y = new
            {
                type = "number",
                description =
                    "Y position of link area in PDF coordinates, origin at bottom-left corner (required for add)"
            },
            width = new
            {
                type = "number",
                description = "Width of link area in PDF points (required for add)"
            },
            height = new
            {
                type = "number",
                description = "Height of link area in PDF points (required for add)"
            },
            url = new
            {
                type = "string",
                description = "URL to link to (for add, edit)"
            },
            targetPage = new
            {
                type = "number",
                description = "Target page number (1-based, for add, edit)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddLink(path, outputPath!, arguments),
            "delete" => await DeleteLink(path, outputPath!, arguments),
            "edit" => await EditLink(path, outputPath!, arguments),
            "get" => await GetLinks(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a link to a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, x, y, width, height, url or targetPage</param>
    /// <returns>Success message</returns>
    private Task<string> AddLink(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var x = ArgumentHelper.GetDouble(arguments, "x");
            var y = ArgumentHelper.GetDouble(arguments, "y");
            var width = ArgumentHelper.GetDouble(arguments, "width");
            var height = ArgumentHelper.GetDouble(arguments, "height");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var targetPage = ArgumentHelper.GetIntNullable(arguments, "targetPage");

            using var document = new Document(path);
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
            document.Save(outputPath);
            return $"Added link to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a link from a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, linkIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteLink(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var linkIndex = ArgumentHelper.GetInt(arguments, "linkIndex");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var links = page.Annotations.OfType<LinkAnnotation>().ToList();
            if (linkIndex < 0 || linkIndex >= links.Count)
                throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

            page.Annotations.Delete(links[linkIndex]);
            document.Save(outputPath);
            return $"Deleted link {linkIndex} from page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a link in a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, linkIndex, optional url, targetPage</param>
    /// <returns>Success message</returns>
    private Task<string> EditLink(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var linkIndex = ArgumentHelper.GetInt(arguments, "linkIndex");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var targetPage = ArgumentHelper.GetIntNullable(arguments, "targetPage");

            using var document = new Document(path);
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

            document.Save(outputPath);
            return $"Edited link {linkIndex} on page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all links from a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="arguments">JSON arguments containing optional pageIndex</param>
    /// <returns>JSON string with all links</returns>
    private Task<string> GetLinks(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

            using var document = new Document(path);
            var linkList = new List<object>();

            if (pageIndex.HasValue)
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
        });
    }
}