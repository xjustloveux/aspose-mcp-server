using System.IO;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Core;

/// <summary>
/// Security helper class for file path and name validation
/// </summary>
public static class SecurityHelper
{
    // Maximum allowed path length (Windows MAX_PATH is 260, but we use a safer limit)
    private const int MaxPathLength = 260;
    
    // Maximum allowed file name length
    private const int MaxFileNameLength = 255;
    
    // Maximum allowed array size for input validation
    private const int MaxArraySize = 1000;
    
    // Maximum allowed string length for input parameters
    private const int MaxStringLength = 10000;

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

        // Limit length first
        if (fileName.Length > MaxFileNameLength)
        {
            fileName = fileName.Substring(0, MaxFileNameLength);
        }

        // Remove path separators and other dangerous characters
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

        // Remove path traversal attempts (more comprehensive)
        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace("\\", "_");
        sanitized = sanitized.Replace("/", "_");
        sanitized = sanitized.Replace(":", "_"); // Remove drive letters
        sanitized = Regex.Replace(sanitized, @"^\s+|\s+$", ""); // Remove leading/trailing whitespace

        // Remove leading/trailing dots and spaces
        sanitized = sanitized.Trim('.', ' ');

        // Ensure it's not empty
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "file";
        }

        // Final length check
        if (sanitized.Length > MaxFileNameLength)
        {
            sanitized = sanitized.Substring(0, MaxFileNameLength);
        }

        return sanitized;
    }

    /// <summary>
    /// Validates that a file path is safe and doesn't contain path traversal attempts
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths (default: false for security)</param>
    /// <returns>True if path is safe, false otherwise</returns>
    public static bool IsSafeFilePath(string filePath, bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return false;
        }

        // Check path length
        if (filePath.Length > MaxPathLength)
        {
            return false;
        }

        // Check for path traversal patterns
        if (filePath.Contains("..") || filePath.Contains("~"))
        {
            return false;
        }

        // Check for dangerous patterns
        if (filePath.Contains("//") || filePath.Contains("\\\\"))
        {
            return false;
        }

        // Check for absolute paths
        if (Path.IsPathRooted(filePath))
        {
            if (!allowAbsolutePaths)
            {
                return false;
            }
            
            // Additional validation for absolute paths
            try
            {
                var fullPath = Path.GetFullPath(filePath);
                // Reject if normalized path contains traversal
                if (fullPath.Contains(".."))
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        // Check for invalid characters
        if (filePath.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Validates and sanitizes a file path, throwing exception if invalid
    /// </summary>
    /// <param name="filePath">File path to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="allowAbsolutePaths">Whether to allow absolute paths</param>
    /// <returns>Validated file path</returns>
    /// <exception cref="ArgumentException">Thrown if path is invalid</exception>
    public static string ValidateFilePath(string filePath, string paramName = "path", bool allowAbsolutePaths = false)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException($"{paramName} cannot be null or empty");
        }

        if (!IsSafeFilePath(filePath, allowAbsolutePaths))
        {
            throw new ArgumentException($"{paramName} contains invalid characters or path traversal attempts");
        }

        return filePath;
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

        // Limit length
        if (pattern.Length > MaxFileNameLength)
        {
            pattern = pattern.Substring(0, MaxFileNameLength);
        }

        // Remove path separators
        var sanitized = pattern.Replace("\\", "_").Replace("/", "_");
        
        // Remove path traversal attempts
        sanitized = sanitized.Replace("..", "");
        sanitized = sanitized.Replace(":", "_");
        
        // Remove leading/trailing dots and spaces
        sanitized = sanitized.Trim('.', ' ');

        // Ensure it's not empty
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "file_{index}";
        }

        return sanitized;
    }

    /// <summary>
    /// Validates array size to prevent resource exhaustion
    /// </summary>
    /// <param name="array">Array to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxSize">Maximum allowed size (default: MaxArraySize)</param>
    /// <exception cref="ArgumentException">Thrown if array is too large</exception>
    public static void ValidateArraySize<T>(IEnumerable<T> array, string paramName = "array", int? maxSize = null)
    {
        if (array == null)
        {
            return;
        }

        var count = array.Count();
        var limit = maxSize ?? MaxArraySize;
        
        if (count > limit)
        {
            throw new ArgumentException($"{paramName} exceeds maximum allowed size of {limit}");
        }
    }

    /// <summary>
    /// Validates string length to prevent resource exhaustion
    /// </summary>
    /// <param name="value">String to validate</param>
    /// <param name="paramName">Parameter name for error message</param>
    /// <param name="maxLength">Maximum allowed length (default: MaxStringLength)</param>
    /// <exception cref="ArgumentException">Thrown if string is too long</exception>
    public static void ValidateStringLength(string value, string paramName = "value", int? maxLength = null)
    {
        if (value == null)
        {
            return;
        }

        var limit = maxLength ?? MaxStringLength;
        
        if (value.Length > limit)
        {
            throw new ArgumentException($"{paramName} exceeds maximum allowed length of {limit}");
        }
    }
}

