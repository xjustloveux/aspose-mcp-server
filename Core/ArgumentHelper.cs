using System.Text.Json.Nodes;

namespace AsposeMcpServer.Core;

/// <summary>
/// Helper class for argument parsing, type conversion, and path validation
/// </summary>
public static class ArgumentHelper
{
    /// <summary>
    /// Safely converts a JSON node to an integer, handling both string and number types
    /// </summary>
    /// <param name="node">JSON node to convert</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Converted integer value</returns>
    /// <exception cref="ArgumentException">Thrown if conversion fails or parameter is required but missing</exception>
    public static int GetInt(JsonNode? node, string paramName, bool required = true)
    {
        if (node == null)
        {
            if (required)
                throw new ArgumentException($"{paramName} is required");
            throw new ArgumentException($"{paramName} is required");
        }

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !int.TryParse(str, out int result))
                throw new ArgumentException($"{paramName} must be a valid integer");
            return result;
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            return node.GetValue<int>();
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid integer");
        }
    }

    /// <summary>
    /// Safely converts a JSON node to a nullable integer, handling both string and number types
    /// </summary>
    /// <param name="node">JSON node to convert</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Converted integer value or null if node is null</returns>
    /// <exception cref="ArgumentException">Thrown if conversion fails</exception>
    public static int? GetIntNullable(JsonNode? node, string paramName)
    {
        if (node == null)
            return null;

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !int.TryParse(str, out int result))
                throw new ArgumentException($"{paramName} must be a valid integer");
            return result;
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            return node.GetValue<int>();
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid integer");
        }
    }

    /// <summary>
    /// Safely converts a JSON node to a double, handling both string and number types (including int)
    /// </summary>
    /// <param name="node">JSON node to convert</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Converted double value</returns>
    /// <exception cref="ArgumentException">Thrown if conversion fails or parameter is required but missing</exception>
    public static double GetDouble(JsonNode? node, string paramName, bool required = true)
    {
        if (node == null)
        {
            if (required)
                throw new ArgumentException($"{paramName} is required");
            throw new ArgumentException($"{paramName} is required");
        }

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !double.TryParse(str, out double result))
                throw new ArgumentException($"{paramName} must be a valid number");
            return result;
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            // Try double first, then int (for compatibility with JSON numbers)
            if (node.AsValue().TryGetValue<double>(out var doubleValue))
            {
                return doubleValue;
            }
            else if (node.AsValue().TryGetValue<int>(out var intValue))
            {
                return intValue;
            }
            else
            {
                return node.GetValue<double>();
            }
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid number");
        }
    }

    /// <summary>
    /// Safely converts a JSON node to a nullable double, handling both string and number types
    /// </summary>
    /// <param name="node">JSON node to convert</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Converted double value or null if node is null</returns>
    /// <exception cref="ArgumentException">Thrown if conversion fails</exception>
    public static double? GetDoubleNullable(JsonNode? node, string paramName)
    {
        if (node == null)
            return null;

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !double.TryParse(str, out double result))
                throw new ArgumentException($"{paramName} must be a valid number");
            return result;
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            return node.GetValue<double>();
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid number");
        }
    }

    /// <summary>
    /// Gets an integer from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Converted integer value</returns>
    public static int GetInt(JsonObject? arguments, string key, string paramName, bool required = true)
    {
        return GetInt(arguments?[key], paramName, required);
    }

    /// <summary>
    /// Gets an integer from JSON arguments, checking multiple parameter names (for compatibility)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name</param>
    /// <param name="alternateName">Alternate parameter name (for compatibility)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Converted integer value</returns>
    public static int GetInt(JsonObject? arguments, string primaryName, string? alternateName, string paramName, bool required = true)
    {
        JsonNode? node = null;
        if (arguments != null)
        {
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);
        }
        return GetInt(node, paramName, required);
    }

    /// <summary>
    /// Gets a nullable integer from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Converted integer value or null if not found</returns>
    public static int? GetIntNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetIntNullable(arguments?[key], paramName);
    }

    /// <summary>
    /// Gets a nullable integer from JSON arguments, checking multiple parameter names (for compatibility)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name</param>
    /// <param name="alternateName">Alternate parameter name (for compatibility)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Converted integer value or null if not found</returns>
    public static int? GetIntNullable(JsonObject? arguments, string primaryName, string? alternateName, string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
        {
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);
        }
        return GetIntNullable(node, paramName);
    }

    /// <summary>
    /// Gets a double from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Converted double value</returns>
    public static double GetDouble(JsonObject? arguments, string key, string paramName, bool required = true)
    {
        return GetDouble(arguments?[key], paramName, required);
    }

    /// <summary>
    /// Gets a nullable double from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Converted double value or null if not found</returns>
    public static double? GetDoubleNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetDoubleNullable(arguments?[key], paramName);
    }

    /// <summary>
    /// Gets a string value from JSON arguments with validation
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>String value</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing</exception>
    public static string GetString(JsonObject? arguments, string key, string paramName, bool required = true)
    {
        var value = arguments?[key]?.GetValue<string>();
        if (required && string.IsNullOrEmpty(value))
            throw new ArgumentException($"{paramName} is required");
        return value ?? string.Empty;
    }

    /// <summary>
    /// Gets a nullable string value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <returns>String value or null if not found</returns>
    public static string? GetStringNullable(JsonObject? arguments, string key)
    {
        return arguments?[key]?.GetValue<string>();
    }

    /// <summary>
    /// Gets a boolean value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="defaultValue">Default value if not found (default: false)</param>
    /// <returns>Boolean value</returns>
    public static bool GetBool(JsonObject? arguments, string key, string paramName, bool defaultValue = false)
    {
        var node = arguments?[key];
        if (node == null)
            return defaultValue;

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.True || node.GetValueKind() == System.Text.Json.JsonValueKind.False)
        {
            return node.GetValue<bool>();
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (bool.TryParse(str, out bool result))
                return result;
            throw new ArgumentException($"{paramName} must be a valid boolean");
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            // Support 0/1 as boolean
            var num = node.GetValue<int>();
            return num != 0;
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid boolean");
        }
    }

    /// <summary>
    /// Gets a nullable boolean value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">Parameter key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Boolean value or null if not found</returns>
    public static bool? GetBoolNullable(JsonObject? arguments, string key, string paramName)
    {
        var node = arguments?[key];
        if (node == null)
            return null;

        if (node.GetValueKind() == System.Text.Json.JsonValueKind.True || node.GetValueKind() == System.Text.Json.JsonValueKind.False)
        {
            return node.GetValue<bool>();
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (bool.TryParse(str, out bool result))
                return result;
            throw new ArgumentException($"{paramName} must be a valid boolean");
        }
        else if (node.GetValueKind() == System.Text.Json.JsonValueKind.Number)
        {
            var num = node.GetValue<int>();
            return num != 0;
        }
        else
        {
            throw new ArgumentException($"{paramName} must be a valid boolean");
        }
    }

    /// <summary>
    /// Gets output path from arguments, defaulting to input path if not specified
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="inputPath">Input file path</param>
    /// <returns>Output path</returns>
    public static string GetOutputPath(JsonObject? arguments, string inputPath)
    {
        return arguments?["outputPath"]?.GetValue<string>() ?? inputPath;
    }

    /// <summary>
    /// Gets and validates input path from arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="paramName">Parameter name for error messages (default: "path")</param>
    /// <returns>Validated input path</returns>
    /// <exception cref="ArgumentException">Thrown if path is missing or invalid</exception>
    public static string GetAndValidatePath(JsonObject? arguments, string paramName = "path")
    {
        var path = arguments?[paramName]?.GetValue<string>();
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException($"{paramName} is required");
        SecurityHelper.ValidateFilePath(path, paramName);
        return path;
    }

    /// <summary>
    /// Gets and validates output path from arguments, defaulting to input path if not specified
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="inputPath">Input file path</param>
    /// <param name="paramName">Parameter name for error messages (default: "outputPath")</param>
    /// <returns>Validated output path</returns>
    /// <exception cref="ArgumentException">Thrown if output path is invalid</exception>
    public static string GetAndValidateOutputPath(JsonObject? arguments, string inputPath, string paramName = "outputPath")
    {
        var outputPath = arguments?[paramName]?.GetValue<string>() ?? inputPath;
        SecurityHelper.ValidateFilePath(outputPath, paramName);
        return outputPath;
    }
}

