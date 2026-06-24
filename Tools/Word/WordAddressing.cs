using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Shared definitions for the unified paragraph-addressing parameters that every Word tool with a
///     paragraph-index operation exposes (<c>storyType</c>, <c>headerFooterType</c>,
///     <c>containerIndex</c>, <c>handle</c>). Centralizing the descriptions and the forwarding logic
///     keeps the addressing model identical across tools, so an LLM agent learns it once.
/// </summary>
internal static class WordAddressing
{
    /// <summary>Description for the <c>storyType</c> tool parameter.</summary>
    public const string StoryTypeDesc =
        "Story the paragraph index is relative to: Body (default), Header, Footer, TextBox, Comment, " +
        "Footnote, Endnote. Use the storyType reported by 'get'/'search' to address non-body paragraphs.";

    /// <summary>Description for the <c>headerFooterType</c> tool parameter.</summary>
    public const string HeaderFooterTypeDesc =
        "For Header/Footer stories: Primary (default), First, or Even.";

    /// <summary>Description for the <c>containerIndex</c> tool parameter.</summary>
    public const string ContainerIndexDesc =
        "Instance selector for multi-instance stories: TextBox shape order, Comment id, or " +
        "footnote/endnote order (as reported by 'get'/'search').";

    /// <summary>Description for the <c>handle</c> tool parameter.</summary>
    public const string HandleDesc =
        "Stable paragraph handle from a prior 'get'/'search' result (session mode only). When set it " +
        "takes precedence over the index coordinates and survives index shifts from edits.";

    /// <summary>
    ///     Forwards the unified addressing parameters into an <see cref="OperationParameters" /> bag,
    ///     setting only the values the caller supplied so story/section/handle defaults still apply.
    /// </summary>
    /// <param name="parameters">The operation parameters to populate.</param>
    /// <param name="storyType">The story type, or null for Body.</param>
    /// <param name="headerFooterType">The Header/Footer discriminator, or null for Primary.</param>
    /// <param name="containerIndex">The container instance selector, or null.</param>
    /// <param name="handle">The stable paragraph handle, or null.</param>
    public static void Apply(OperationParameters parameters, string? storyType, string? headerFooterType,
        int? containerIndex, string? handle)
    {
        if (!string.IsNullOrEmpty(storyType)) parameters.Set("storyType", storyType);
        if (!string.IsNullOrEmpty(headerFooterType)) parameters.Set("headerFooterType", headerFooterType);
        if (containerIndex.HasValue) parameters.Set("containerIndex", containerIndex.Value);
        if (!string.IsNullOrEmpty(handle)) parameters.Set("handle", handle);
    }
}
