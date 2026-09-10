namespace AsposeMcpServer.Core.Session;

/// <summary>
///     What sealing a session for closing achieved.
///     <para>
///         A close used to take the same barrier a save takes, and its single <c>false</c> meant
///         both "another operation holds this session" and "the in-flight operations did not
///         finish". The close reported the second, skipped its save and disposed the document
///         under the save that held it, so neither wrote (R4-S06). These are separate answers
///         because a caller has to act on them differently.
///     </para>
/// </summary>
public enum SessionSealOutcome
{
    /// <summary>
    ///     The session is sealed, nothing else holds it and no operation is in flight. The caller
    ///     owns the document outright and may save and dispose it.
    /// </summary>
    Sealed,

    /// <summary>
    ///     Another caller had already sealed this session; it is closing, but not by this caller.
    ///     The document belongs to that caller, so this one must not save or dispose it.
    /// </summary>
    AlreadyClosing,

    /// <summary>
    ///     A save or auto-save held the session exclusively and did not finish within the timeout.
    ///     The session is sealed against new work, but the document is still someone else's.
    /// </summary>
    HeldByAnotherOperation,

    /// <summary>
    ///     Operations that started before the seal were still running when the timeout passed.
    ///     The document may be mid-mutation, so saving it would write a torn state.
    /// </summary>
    ActiveOperationsRemain
}
