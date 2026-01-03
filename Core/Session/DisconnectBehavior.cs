namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Behavior when client disconnects with open sessions
/// </summary>
public enum DisconnectBehavior
{
    /// <summary>Automatically save to original file</summary>
    AutoSave,

    /// <summary>Discard all changes</summary>
    Discard,

    /// <summary>Save to temp file for recovery</summary>
    SaveToTemp,

    /// <summary>Prompt on reconnect</summary>
    PromptOnReconnect
}