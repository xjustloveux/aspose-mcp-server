using Aspose.Words.Loading;
using PdfLoadOptions = Aspose.Pdf.LoadOptions;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Blocks a document from pulling in resources the caller never named.
///     HTML, MHT and content-sniffed Word inputs can reference images, style sheets and other
///     resources by URI; loading them turns a document conversion into a request issued by the
///     server, which reaches internal hosts and cloud metadata endpoints that the caller cannot
///     otherwise address, and can also embed the contents of an unrelated local file into the
///     output. Every loader in this project therefore refuses remote URIs outright and holds local
///     references to the same allowlist that governs the caller's own paths.
/// </summary>
public static class ExternalResourceGuard
{
    /// <summary>
    ///     Builds the callback that Aspose.Words consults for each external resource.
    /// </summary>
    /// <param name="allowedBasePaths">
    ///     Admin-configured allowlist. An empty list means no allowlist is configured, in which case
    ///     local references are permitted and only remote URIs are refused.
    /// </param>
    /// <returns>A callback that skips every resource the policy refuses.</returns>
    public static IResourceLoadingCallback CreateWordsCallback(IReadOnlyList<string> allowedBasePaths)
    {
        return new BlockingResourceLoadingCallback(allowedBasePaths);
    }

    /// <summary>
    ///     Builds the strategy that Aspose.Pdf consults for each external resource of an HTML or
    ///     MHT input.
    /// </summary>
    /// <param name="allowedBasePaths">
    ///     Admin-configured allowlist. An empty list means no allowlist is configured, in which case
    ///     local references are permitted and only remote URIs are refused.
    /// </param>
    /// <returns>
    ///     A strategy that returns empty data for a refused resource. Conversion continues without
    ///     the resource rather than failing, matching how the converter already treats a resource
    ///     that cannot be fetched.
    /// </returns>
    public static PdfLoadOptions.ResourceLoadingStrategy CreatePdfStrategy(IReadOnlyList<string> allowedBasePaths)
    {
        return uri =>
        {
            if (IsAllowed(uri, allowedBasePaths))
                return new PdfLoadOptions.ResourceLoadingResult([]) { LoadingCancelled = true };

            return new PdfLoadOptions.ResourceLoadingResult([])
            {
                ExceptionOfLoadingIfAny = new UnauthorizedAccessException(
                    "Loading external resources referenced by the document is not permitted.")
            };
        };
    }

    /// <summary>
    ///     Decides whether a referenced resource may be loaded. Exposed so the policy can be
    ///     exercised directly; the library's callback argument type cannot be constructed in a test.
    /// </summary>
    /// <param name="uri">The resource URI exactly as the document declared it.</param>
    /// <param name="allowedBasePaths">Admin-configured allowlist; empty means no allowlist.</param>
    /// <returns><c>true</c> when the reference is local and inside the allowlist; otherwise <c>false</c>.</returns>
    public static bool IsAllowed(string? uri, IReadOnlyList<string> allowedBasePaths)
    {
        // An empty URI carries no reference to fetch: the caller already supplied the bytes
        // (for example DocumentBuilder.InsertImage with a stream), so there is nothing to refuse.
        if (string.IsNullOrWhiteSpace(uri)) return true;

        // Inline data carries no request and no file read.
        if (uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return true;

        if (Uri.TryCreate(uri, UriKind.Absolute, out var absolute) && !absolute.IsFile)
            return false;

        // A reference that names another host is a network fetch even when .NET classifies it as
        // a file URI, and refusing it must not depend on an allowlist being configured (R2-S05).
        if (NamesAnotherHost(uri, absolute)) return false;

        if (allowedBasePaths.Count == 0) return true;

        try
        {
            var localPath = absolute is { IsFile: true } ? absolute.LocalPath : uri;
            SecurityHelper.ResolveAndEnsureWithinAllowlist(localPath, allowedBasePaths, "resourceUri");
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or IOException)
        {
            return false;
        }
    }

    /// <summary>
    ///     Whether a reference points at another machine rather than at this one.
    ///     <para>
    ///         <c>file://host/share/x</c> is reported by <c>System.Uri</c> as
    ///         <c>IsFile = true</c>, so the remote check above lets it through; following it opens
    ///         an SMB connection to a host the document chose, which both fetches remote content
    ///         and hands that host an NTLM authentication attempt. A bare <c>\\host\share</c> and a
    ///         protocol-relative <c>//host/share</c> reach the same place, and the latter is not an
    ///         absolute URI at all, so it never met the remote check.
    ///     </para>
    /// </summary>
    /// <param name="uri">The resource URI exactly as the document declared it.</param>
    /// <param name="absolute">The parsed URI, or <c>null</c> when the reference is not absolute.</param>
    /// <returns><c>true</c> when the reference resolves to a remote host.</returns>
    private static bool NamesAnotherHost(string uri, Uri? absolute)
    {
        var trimmed = uri.TrimStart();
        if (trimmed.Length >= 2 && IsSeparator(trimmed[0]) && IsSeparator(trimmed[1]))
            return true;

        return absolute is not null && (absolute.IsUnc || !string.IsNullOrEmpty(absolute.Host));
    }

    /// <summary>
    ///     Whether a character separates path segments on either platform.
    /// </summary>
    /// <param name="value">The character to test.</param>
    /// <returns><c>true</c> for a forward or backward slash.</returns>
    private static bool IsSeparator(char value)
    {
        return value == '/' || value == '\\';
    }

    /// <summary>
    ///     Aspose.Words callback that skips every resource the policy refuses.
    /// </summary>
    /// <param name="allowedBasePaths">Admin-configured allowlist; empty means no allowlist.</param>
    private sealed class BlockingResourceLoadingCallback(IReadOnlyList<string> allowedBasePaths)
        : IResourceLoadingCallback
    {
        /// <inheritdoc />
        public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
        {
            // Both forms are checked: Uri is the reference Aspose.Words will act on, OriginalUri is
            // the reference as written in the document. A redirect or rewrite between the two must
            // not be able to turn a refused target into an accepted one.
            return IsAllowed(args.Uri, allowedBasePaths) && IsAllowed(args.OriginalUri, allowedBasePaths)
                ? ResourceLoadingAction.Default
                : ResourceLoadingAction.Skip;
        }
    }
}
