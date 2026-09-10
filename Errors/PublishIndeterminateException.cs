namespace AsposeMcpServer.Errors;

/// <summary>
///     A publish whose outputs are in place but whose record could not be settled.
///     <para>
///         Every other failure a publish can raise means the destinations were put back and nothing
///         changed, so retrying is the right response. This one does not: the outputs are delivered
///         and only the journal that says so could neither be marked committed nor removed. A caller
///         that treats it like the others repeats a publish that already happened, and a later start
///         that reads the journal undoes one that already succeeded (R18-CONTRACT01).
///     </para>
///     <para>
///         A distinct type, because the difference cannot be read off the message, and because the
///         batch that raises it has already passed its commit point and will refuse to publish
///         again.
///     </para>
/// </summary>
public class PublishIndeterminateException : IOException
{
    /// <summary>Creates the exception.</summary>
    /// <param name="message">What happened, composed in this repository.</param>
    /// <param name="journalPath">The record an operator has to deal with.</param>
    /// <param name="inner">The failure that left the record unsettled.</param>
    public PublishIndeterminateException(string message, string journalPath, Exception inner)
        : base(message, inner)
    {
        JournalPath = journalPath;
    }

    /// <summary>The publish record that needs an operator before the next start.</summary>
    /// <remarks>
    ///     Named rather than only described, because the one action that resolves this — look at
    ///     that file and decide whether the publish stands — needs the path to act on.
    /// </remarks>
    public string JournalPath { get; }
}
