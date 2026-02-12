using Aspose.Slides;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Represents the result of creating a hyperlink.
/// </summary>
/// <param name="Hyperlink">The created hyperlink object.</param>
/// <param name="Description">A description of the hyperlink target (URL or slide reference).</param>
public record HyperlinkCreationResult(IHyperlink Hyperlink, string Description);
