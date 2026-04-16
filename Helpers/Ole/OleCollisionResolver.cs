namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Resolves filename collisions during <c>extract_all</c> by appending <c>" (2)"</c>,
///     <c>" (3)"</c>, ... before the extension. Stateful per <c>extract_all</c> call so
///     two OLE objects with the same sanitized name produced in-memory do not collide
///     even before the first <see cref="Reserve" /> touches the filesystem.
/// </summary>
public sealed class OleCollisionResolver
{
    /// <summary>
    ///     Names already reserved during this resolver's lifetime. Stored case-insensitively
    ///     because NTFS / HFS+ are case-insensitive by default and a round-trip across OS
    ///     boundaries should not produce two files that differ only in case.
    /// </summary>
    private readonly HashSet<string> _reserved = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    ///     Returns a path within <paramref name="outputDirectory" /> that does not already
    ///     exist on disk and has not been reserved by a previous <see cref="Reserve" /> call
    ///     on this instance.
    /// </summary>
    /// <param name="outputDirectory">
    ///     Destination directory. Must already exist and be writable; the resolver does not
    ///     create it. Callers typically invoke <see cref="Directory.CreateDirectory(string)" />
    ///     before the first <see cref="Reserve" />.
    /// </param>
    /// <param name="preferredName">
    ///     Sanitized preferred filename (already passed through
    ///     <see cref="OleSanitizerHelper.SanitizeOleFileName" />). Must not be null or empty.
    /// </param>
    /// <returns>
    ///     Absolute file path guaranteed unique at call time (both against the filesystem
    ///     and against previously reserved names within this resolver).
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when <paramref name="outputDirectory" /> or <paramref name="preferredName" />
    ///     is null, empty, or whitespace.
    /// </exception>
    public string Reserve(string outputDirectory, string preferredName)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException("Output directory is required", nameof(outputDirectory));
        if (string.IsNullOrWhiteSpace(preferredName))
            throw new ArgumentException("Preferred name is required", nameof(preferredName));

        var stem = Path.GetFileNameWithoutExtension(preferredName);
        var ext = Path.GetExtension(preferredName);
        var candidate = preferredName;
        var counter = 2;

        while (IsReservedOrExists(outputDirectory, candidate))
        {
            candidate = $"{stem} ({counter}){ext}";
            counter++;
        }

        _reserved.Add(candidate);
        return Path.Combine(outputDirectory, candidate);
    }

    /// <summary>
    ///     Determines whether the candidate name has been previously reserved by this
    ///     resolver or already exists on disk in the destination directory.
    /// </summary>
    /// <param name="outputDirectory">Destination directory being written to.</param>
    /// <param name="candidate">Candidate filename (not a full path).</param>
    /// <returns><c>true</c> if the name is already claimed; otherwise <c>false</c>.</returns>
    private bool IsReservedOrExists(string outputDirectory, string candidate)
    {
        if (_reserved.Contains(candidate)) return true;
        return File.Exists(Path.Combine(outputDirectory, candidate));
    }
}
