using System.IO;

namespace AsposeMcpServer.Core;

/// <summary>
/// Security helper class for file path and name validation
/// </summary>
public static class SecurityHelper
{
    /// <summary>
    /// Sanitizes a file name to prevent path traversal attacks
    /// </summary>
    /// <param name="fileName">Original file name</param>
    /// <returns>Sanitized file name safe for use in file operations</returns>
    public static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return "file";
        }

        // Remove path separators and other dangerous characters
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

        // Remove path traversal attempts
        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace("\\", "_");
        sanitized = sanitized.Replace("/", "_");

        // Remove leading/trailing dots and spaces
        sanitized = sanitized.Trim('.', ' ');

        // Ensure it's not empty
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "file";
        }

        // Limit length to prevent issues
        if (sanitized.Length > 255)
        {
            sanitized = sanitized.Substring(0, 255);
        }

        return sanitized;
    }

    /// <summary>
    /// Validates that a file path is safe and doesn't contain path traversal attempts
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <returns>True if path is safe, false otherwise</returns>
    public static bool IsSafeFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return false;
        }

        // Check for path traversal patterns
        if (filePath.Contains("..") || filePath.Contains("~"))
        {
            return false;
        }

        // Check for absolute paths that might be dangerous
        if (Path.IsPathRooted(filePath))
        {
            // Allow absolute paths but log them (could be restricted further if needed)
            return true;
        }

        return true;
    }

    /// <summary>
    /// Validates and sanitizes a file name pattern (for use in split/export tools)
    /// </summary>
    /// <param name="pattern">File name pattern (may contain placeholders like {index}, {name})</param>
    /// <returns>Sanitized pattern safe for use</returns>
    public static string SanitizeFileNamePattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
        {
            return "file_{index}";
        }

        // Remove path separators
        var sanitized = pattern.Replace("\\", "_").Replace("/", "_");
        
        // Remove path traversal attempts
        sanitized = sanitized.Replace("..", "");
        
        // Remove leading/trailing dots and spaces
        sanitized = sanitized.Trim('.', ' ');

        // Ensure it's not empty
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "file_{index}";
        }

        return sanitized;
    }
}

