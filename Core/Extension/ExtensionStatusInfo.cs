namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Status information for an extension.
/// </summary>
public class ExtensionStatusInfo
{
    /// <summary>
    ///     Gets or sets the extension ID.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the extension name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the extension version.
    /// </summary>
    public string Version { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the localized title for display.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    ///     Gets or sets the description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    ///     Gets or sets the author.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    ///     Gets or sets the website URL.
    /// </summary>
    public string? WebsiteUrl { get; set; }

    /// <summary>
    ///     Gets or sets whether the extension is available.
    /// </summary>
    public bool IsAvailable { get; set; }

    /// <summary>
    ///     Gets or sets whether the extension is currently initializing.
    /// </summary>
    public bool IsInitializing { get; set; }

    /// <summary>
    ///     Gets or sets the reason the extension is unavailable.
    /// </summary>
    public string? UnavailableReason { get; set; }

    /// <summary>
    ///     Gets or sets the current state.
    /// </summary>
    public ExtensionState State { get; set; }

    /// <summary>
    ///     Gets or sets the last activity time.
    /// </summary>
    public DateTime? LastActivity { get; set; }

    /// <summary>
    ///     Gets or sets the restart count.
    /// </summary>
    public int RestartCount { get; set; }
}
