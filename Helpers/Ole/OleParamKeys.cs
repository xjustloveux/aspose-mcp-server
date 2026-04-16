namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Parameter key constants shared by the three OLE tools (Word / Excel / PowerPoint)
///     and their 12 handlers. Keeping keys in one place guarantees cross-tool parity
///     (AC-18 / AC-19) and prevents silent drift if one tool renames a key.
/// </summary>
public static class OleParamKeys
{
    /// <summary>Source file path (file-mode only).</summary>
    public const string Path = "path";

    /// <summary>Optional password for protected sources (file-mode only).</summary>
    public const string Password = "password";

    /// <summary>Optional rewrite target for file-mode <c>remove</c>.</summary>
    public const string OutputPath = "outputPath";

    /// <summary>Destination directory for <c>extract</c> / <c>extract_all</c>.</summary>
    public const string OutputDirectory = "outputDirectory";

    /// <summary>Zero-based OLE index.</summary>
    public const string OleIndex = "oleIndex";

    /// <summary>Optional sanitized filename override for <c>extract</c>.</summary>
    public const string OutputFileName = "outputFileName";

    /// <summary>
    ///     Session-mode advisory flag — <c>true</c> when the tool layer has detected
    ///     that a non-null <c>password</c> was supplied but ignored because the session was
    ///     already unlocked.
    /// </summary>
    public const string PasswordIgnored = "passwordIgnored";
}
