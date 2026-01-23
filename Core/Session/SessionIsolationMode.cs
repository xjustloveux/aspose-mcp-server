namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Session isolation mode for configurable group-based isolation.
///     The group identifier is determined by the GROUP_IDENTIFIER_CLAIM configuration.
/// </summary>
public enum SessionIsolationMode
{
    /// <summary>
    ///     No isolation - all users can access all sessions (backward compatible with Stdio mode)
    /// </summary>
    None,

    /// <summary>
    ///     Group-level isolation - users within the same group can access each other's sessions.
    ///     The group is determined by the configured GROUP_IDENTIFIER_CLAIM (e.g., tenant_id, team_id, sub).
    /// </summary>
    Group
}
