using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Tool-layer helpers shared by the three <c>*_ole_object</c> tools. Keeps the
///     password-ignored-note attachment logic in one place so the F-5 locked-shape
///     contract cannot drift between Word, Excel, and PowerPoint.
/// </summary>
public static class OleToolHelper
{
    /// <summary>
    ///     When <paramref name="passwordIgnored" /> is <c>true</c>, returns a copy of the
    ///     supplied result record with the locked-shape
    ///     <see cref="PasswordIgnoredNote" /> attached; otherwise returns the input
    ///     unchanged. Non-OLE result types pass through (defensive — the dispatcher only
    ///     feeds OLE results here).
    /// </summary>
    /// <param name="result">The handler result to (optionally) augment.</param>
    /// <param name="passwordIgnored">
    ///     <c>true</c> when the tool layer detected that a non-null password was supplied
    ///     in session-mode and ignored.
    /// </param>
    /// <returns>The original or augmented result.</returns>
    public static object AttachPasswordIgnoredNote(object result, bool passwordIgnored)
    {
        if (!passwordIgnored) return result;
        var note = new PasswordIgnoredNote();
        return result switch
        {
            OleListResult list => list with { PasswordIgnored = note },
            OleExtractResult ext => ext with { PasswordIgnored = note },
            OleExtractAllResult all => all with { PasswordIgnored = note },
            OleRemoveResult rem => rem with { PasswordIgnored = note },
            _ => result
        };
    }
}
