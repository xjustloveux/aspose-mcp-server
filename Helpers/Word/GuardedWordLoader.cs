using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     The one way this server opens a Word document (A-03).
///     <para>
///         Aspose.Words decides the real format from the file's content, not its extension, so an
///         HTML or MHT payload named <c>.docx</c> reaches the same loader — and those formats make
///         the library fetch whatever they reference. <c>DocumentContext</c> supplied a resource
///         guard for that reason, but eight other places called <c>new Document(path)</c> directly
///         and got no guard at all, which meant the server's SSRF policy depended on which entry
///         point the caller happened to use.
///     </para>
///     <para>
///         Routing every load through here makes the policy a property of loading a document
///         rather than of remembering to ask for it.
///     </para>
/// </summary>
public static class GuardedWordLoader
{
    /// <summary>Opens a Word document with external resource loading restricted.</summary>
    /// <param name="path">Path to the document. Callers resolve it against the allowlist first.</param>
    /// <param name="allowedBasePaths">
    ///     Roots a referenced local resource may come from. An empty list means no allowlist is
    ///     configured, and a local reference is then unrestricted — the same meaning the list has
    ///     everywhere else in the server, rather than the opposite one this comment used to
    ///     claim (R3-DOC04). A remote reference, and one naming another host, are refused either
    ///     way.
    /// </param>
    /// <param name="password">Optional password for an encrypted document.</param>
    /// <returns>The loaded document.</returns>
    public static Document Load(string path, IReadOnlyList<string> allowedBasePaths,
        string? password = null)
    {
        return new Document(path, BuildOptions(allowedBasePaths, password));
    }

    /// <summary>Builds the load options carrying the resource guard.</summary>
    /// <param name="allowedBasePaths">Roots a referenced local resource may come from.</param>
    /// <param name="password">Optional password for an encrypted document.</param>
    /// <returns>Load options with the guard installed.</returns>
    public static LoadOptions BuildOptions(IReadOnlyList<string> allowedBasePaths,
        string? password = null)
    {
        var options = new LoadOptions
        {
            ResourceLoadingCallback = ExternalResourceGuard.CreateWordsCallback(allowedBasePaths)
        };

        if (!string.IsNullOrEmpty(password)) options.Password = password;
        return options;
    }
}
