namespace AsposeMcpServer.Helpers;

/// <summary>
///     Security helper class for file path and name validation
/// </summary>
public static class SecurityHelper
{
    /// <summary>
    ///     Windows MAX_PATH limit for file paths
    /// </summary>
    private const int MaxPathLength = 260;

    /// <summary>
    ///     Maximum allowed file name length
    /// </summary>
    private const int MaxFileNameLength = 255;

    /// <summary>
    ///     Maximum allowed array size to prevent resource exhaustion
    /// </summary>
    private const int MaxArraySize = 1000;

    /// <summary>
    ///     Maximum allowed string length to prevent resource exhaustion
    /// </summary>
    private const int MaxStringLength = 10000;

    /// <summary>
    ///     Sanitizes a file name to prevent path traversal attacks
    /// </summary>
    /// <param name="fileName">Original file name</param>
    /// <returns>Sanitized file name safe for use in file operations</returns>
    public static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return "file";

        if (fileName.Length > MaxFileNameLength) fileName = fileName.Substring(0, MaxFileNameLength);

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace("\\", "_");
        sanitized = sanitized.Replace("/", "_");
        sanitized = sanitized.Replace(":", "_");
        sanitized = sanitized.Trim();
        sanitized = sanitized.Trim('.', ' ');

        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "file";

        if (sanitized.Length > MaxFileNameLength) sanitized = sanitized.Substring(0, MaxFileNameLength);

        return sanitized;
    }

    /// <summary>
    ///     Validates that a file path is safe and doesn't contain path traversal attempts
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths (default: false for security)</param>
    /// <returns>True if path is safe, false otherwise</returns>
    public static bool IsSafeFilePath(string filePath, bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return false;

        if (filePath.Length > MaxPathLength) return false;

        if (filePath.Contains("..", StringComparison.Ordinal) || filePath.Contains('~')) return false;

        if (filePath.Contains("//", StringComparison.Ordinal) ||
            filePath.Contains("\\\\", StringComparison.Ordinal)) return false;

        if (filePath.Length >= 3 && char.IsLetter(filePath[0]) && filePath[1] == ':' &&
            (filePath[2] == '\\' || filePath[2] == '/') && !allowAbsolutePaths)
            return false;

        if (Path.IsPathRooted(filePath))
        {
            if (!allowAbsolutePaths) return false;

            try
            {
                var fullPath = Path.GetFullPath(filePath);
                if (fullPath.Contains("..", StringComparison.Ordinal)) return false;
            }
            catch
            {
                return false;
            }
        }

        if (filePath.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;

        return true;
    }

    /// <summary>
    ///     Validates and sanitizes a file path, throwing exception if invalid
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths</param>
    /// <returns>Validated file path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid</exception>
    public static string ValidateFilePath(string filePath, string paramName = "path", bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException($"{paramName} cannot be null or empty");

        if (!IsSafeFilePath(filePath, allowAbsolutePaths))
            throw new ArgumentException($"{paramName} contains invalid characters or path traversal attempts");

        return filePath;
    }

    /// <summary>
    ///     Validates that a file path is within the allowed base paths when a whitelist is configured.
    ///     If allowedBasePaths is empty, all paths are permitted (backward compatible).
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="allowedBasePaths">List of allowed base directories (empty = allow all)</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <exception cref="ArgumentException">Thrown if path is outside all allowed directories</exception>
    public static void ValidatePathWithinAllowedBases(string filePath, IReadOnlyList<string> allowedBasePaths,
        string paramName = "path")
    {
        if (allowedBasePaths.Count == 0)
            return;

        var fullPath = Path.GetFullPath(filePath);
        foreach (var basePath in allowedBasePaths)
        {
            var normalizedBase = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                                 + Path.DirectorySeparatorChar;
            if (fullPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                return;
        }

        throw new ArgumentException(
            $"{paramName} is outside the allowed directories. Configure --allowed-path to permit access.");
    }

    /// <summary>
    ///     Validates and sanitizes a file name pattern (for use in split/export tools)
    /// </summary>
    /// <param name="pattern">File name pattern (may contain placeholders like {index}, {name})</param>
    /// <returns>Sanitized pattern safe for use</returns>
    public static string SanitizeFileNamePattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern)) return "file_{index}";

        if (pattern.Length > MaxFileNameLength) pattern = pattern.Substring(0, MaxFileNameLength);

        var sanitized = pattern.Replace("\\", "_").Replace("/", "_");
        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace(":", "_");
        sanitized = sanitized.Trim('.', ' ');

        if (string.IsNullOrWhiteSpace(sanitized)) sanitized = "file_{index}";

        return sanitized;
    }

    /// <summary>
    ///     Validates array size to prevent resource exhaustion
    /// </summary>
    /// <param name="array">Array to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxSize">Maximum allowed size (default: MaxArraySize)</param>
    /// <exception cref="ArgumentException">Thrown if array is too large</exception>
    public static void ValidateArraySize<T>(IEnumerable<T> array, string paramName = "array", int? maxSize = null)
    {
        var limit = maxSize ?? MaxArraySize;

        var count = array is ICollection<T> collection
            ? collection.Count
            : array.Count();

        if (count > limit) throw new ArgumentException($"{paramName} exceeds maximum allowed size of {limit}");
    }

    /// <summary>
    ///     Validates string length to prevent resource exhaustion
    /// </summary>
    /// <param name="value">String to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxLength">Maximum allowed length (default: MaxStringLength)</param>
    /// <exception cref="ArgumentException">Thrown if string is too long</exception>
    public static void ValidateStringLength(string value, string paramName = "value", int? maxLength = null)
    {
        var limit = maxLength ?? MaxStringLength;

        if (value.Length > limit) throw new ArgumentException($"{paramName} exceeds maximum allowed length of {limit}");
    }
}
