using AsposeMcpServer.Errors.Ole;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Per-tool extension whitelist enforcement. Each of the three OLE tools accepts a
///     fixed set of container formats; anything else is rejected via
///     <see cref="ArgumentException" /> before any filesystem access.
/// </summary>
public static class OleExtensionGuard
{
    /// <summary>Accepted Word container extensions (<c>.docx</c>, <c>.doc</c>).</summary>
    public static readonly IReadOnlyList<string> WordExtensions = [".docx", ".doc"];

    /// <summary>Accepted Excel container extensions (<c>.xlsx</c>, <c>.xls</c>).</summary>
    public static readonly IReadOnlyList<string> ExcelExtensions = [".xlsx", ".xls"];

    /// <summary>Accepted PowerPoint container extensions (<c>.pptx</c>, <c>.ppt</c>).</summary>
    public static readonly IReadOnlyList<string> PptExtensions = [".pptx", ".ppt"];

    /// <summary>
    ///     Asserts that <paramref name="path" /> ends with one of the Word container
    ///     extensions.
    /// </summary>
    /// <param name="path">Path previously validated by <c>ValidateUserPath</c>.</param>
    /// <exception cref="ArgumentException">Thrown when the extension is not accepted.</exception>
    public static void EnsureWordExtension(string path)
    {
        EnsureExtension(path, WordExtensions);
    }

    /// <summary>
    ///     Asserts that <paramref name="path" /> ends with one of the Excel container
    ///     extensions.
    /// </summary>
    /// <param name="path">Path previously validated by <c>ValidateUserPath</c>.</param>
    /// <exception cref="ArgumentException">Thrown when the extension is not accepted.</exception>
    public static void EnsureExcelExtension(string path)
    {
        EnsureExtension(path, ExcelExtensions);
    }

    /// <summary>
    ///     Asserts that <paramref name="path" /> ends with one of the PowerPoint container
    ///     extensions.
    /// </summary>
    /// <param name="path">Path previously validated by <c>ValidateUserPath</c>.</param>
    /// <exception cref="ArgumentException">Thrown when the extension is not accepted.</exception>
    public static void EnsurePptExtension(string path)
    {
        EnsureExtension(path, PptExtensions);
    }

    /// <summary>
    ///     Core extension-match helper.
    /// </summary>
    /// <param name="path">Path to check. May be null or empty.</param>
    /// <param name="accepted">The accepted extension set.</param>
    /// <exception cref="ArgumentException">Thrown when the extension is not accepted.</exception>
    private static void EnsureExtension(string? path, IReadOnlyList<string> accepted)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException(OleErrorMessageBuilder.InvalidPath(path), nameof(path));

        var ext = Path.GetExtension(path);
        foreach (var candidate in accepted)
            if (string.Equals(ext, candidate, StringComparison.OrdinalIgnoreCase))
                return;

        throw new ArgumentException(OleErrorMessageBuilder.InvalidPath(path), nameof(path));
    }
}
