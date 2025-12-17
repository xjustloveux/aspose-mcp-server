using System.Text.Json;
using System.Text.Json.Nodes;

namespace AsposeMcpServer.Core;

/// <summary>
///     Helper class for argument parsing, type conversion, and path validation
/// </summary>
public static class ArgumentHelper
{
    #region GetInt - Integer Methods

    /// <summary>
    ///     Gets a required integer from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Integer value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static int GetInt(JsonObject? arguments, string key)
    {
        return GetInt(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional integer from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>Nullable integer value, or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static int? GetInt(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetInt(arguments, key) for required parameters");
        return GetIntNullable(arguments, key);
    }

    /// <summary>
    ///     Gets an integer from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="defaultValue">Default value to return if key is missing or null</param>
    /// <returns>Integer value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static int GetInt(JsonObject? arguments, string key, int defaultValue)
    {
        return GetInt(arguments, key, null, key, false, defaultValue);
    }

    /// <summary>
    ///     Gets a required integer from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Integer value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static int GetInt(JsonObject? arguments, string key, string paramName)
    {
        return GetInt(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets an optional integer from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <returns>Integer value or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static int? GetInt(JsonObject? arguments, string key, string paramName, bool required)
    {
        return GetInt(arguments, key, null, paramName, required);
    }

    /// <summary>
    ///     Gets an integer from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <param name="defaultValue">Default value to return if key is missing (only used when required is false)</param>
    /// <returns>Integer value</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static int GetInt(JsonObject? arguments, string primaryName, string? alternateName, string paramName,
        bool required = true, int? defaultValue = null)
    {
        var result = GetIntNullable(arguments, primaryName, alternateName, paramName);
        if (result.HasValue)
            return result.Value;

        if (required)
            throw new ArgumentException($"{paramName} is required");
        if (defaultValue.HasValue)
            return defaultValue.Value;
        throw new ArgumentException($"{paramName} is required");
    }

    /// <summary>
    ///     Gets a nullable integer from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Nullable integer value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static int? GetIntNullable(JsonObject? arguments, string key)
    {
        return GetIntNullable(arguments, key, null, key);
    }

    /// <summary>
    ///     Gets a nullable integer from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable integer value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static int? GetIntNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetIntNullable(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets a nullable integer from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable integer value, or null if both keys are missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static int? GetIntNullable(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
            return null;

        if (node.GetValueKind() == JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !int.TryParse(str, out var result))
                throw new ArgumentException($"{paramName} must be a valid integer");
            return result;
        }

        if (node.GetValueKind() == JsonValueKind.Number) return node.GetValue<int>();

        throw new ArgumentException($"{paramName} must be a valid integer");
    }

    #endregion

    #region GetDouble - Double Methods

    /// <summary>
    ///     Gets a required double from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Double value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static double GetDouble(JsonObject? arguments, string key)
    {
        return GetDouble(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional double from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>Nullable double value, or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double? GetDouble(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetDouble(arguments, key) for required parameters");
        return GetDoubleNullable(arguments, key);
    }

    /// <summary>
    ///     Gets a double from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="defaultValue">Default value to return if key is missing or null</param>
    /// <returns>Double value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double GetDouble(JsonObject? arguments, string key, double defaultValue)
    {
        return GetDouble(arguments, key, key, false, defaultValue);
    }

    /// <summary>
    ///     Gets a double from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="defaultValue">Default value to return if key is missing or null</param>
    /// <returns>Double value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double GetDouble(JsonObject? arguments, string key, string paramName, double defaultValue)
    {
        return GetDouble(arguments, key, paramName, false, defaultValue);
    }

    /// <summary>
    ///     Gets a required double from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Double value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static double GetDouble(JsonObject? arguments, string key, string paramName)
    {
        return GetDouble(arguments, key, paramName, true);
    }

    /// <summary>
    ///     Gets an optional double from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <param name="defaultValue">Default value to return if key is missing (only used when required is false)</param>
    /// <returns>Double value</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static double GetDouble(JsonObject? arguments, string key, string paramName, bool required,
        double? defaultValue = null)
    {
        var result = GetDoubleNullable(arguments, key, null, paramName);
        if (result.HasValue)
            return result.Value;

        if (required)
            throw new ArgumentException($"{paramName} is required");
        if (defaultValue.HasValue)
            return defaultValue.Value;
        throw new ArgumentException($"{paramName} is required");
    }

    /// <summary>
    ///     Gets a nullable double from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Nullable double value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double? GetDoubleNullable(JsonObject? arguments, string key)
    {
        return GetDoubleNullable(arguments, key, null, key);
    }

    /// <summary>
    ///     Gets a nullable double from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable double value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double? GetDoubleNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetDoubleNullable(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets a nullable double from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable double value, or null if both keys are missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static double? GetDoubleNullable(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
            return null;

        if (node.GetValueKind() == JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !double.TryParse(str, out var result))
                throw new ArgumentException($"{paramName} must be a valid number");
            return result;
        }

        if (node.GetValueKind() == JsonValueKind.Number)
        {
            // Try double first, then int (for compatibility with JSON numbers)
            if (node.AsValue().TryGetValue<double>(out var doubleValue)) return doubleValue;

            if (node.AsValue().TryGetValue<int>(out var intValue)) return intValue;

            return node.GetValue<double>();
        }

        throw new ArgumentException($"{paramName} must be a valid number");
    }

    #endregion

    #region GetFloat - Float Methods

    /// <summary>
    ///     Gets a required float from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Float value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static float GetFloat(JsonObject? arguments, string key)
    {
        return GetFloat(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional float from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>Nullable float value, or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float? GetFloat(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetFloat(arguments, key) for required parameters");
        return GetFloatNullable(arguments, key);
    }

    /// <summary>
    ///     Gets a float from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="defaultValue">Default value to return if key is missing or null</param>
    /// <returns>Float value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float GetFloat(JsonObject? arguments, string key, float defaultValue)
    {
        return GetFloat(arguments, key, key, false, defaultValue);
    }

    /// <summary>
    ///     Gets a float from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="defaultValue">Default value to return if key is missing or null</param>
    /// <returns>Float value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float GetFloat(JsonObject? arguments, string key, string paramName, float defaultValue)
    {
        return GetFloat(arguments, key, paramName, false, defaultValue);
    }

    /// <summary>
    ///     Gets a required float from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Float value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static float GetFloat(JsonObject? arguments, string key, string paramName)
    {
        return GetFloat(arguments, key, paramName, true);
    }

    /// <summary>
    ///     Gets an optional float from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <param name="defaultValue">Default value to return if key is missing (only used when required is false)</param>
    /// <returns>Float value</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static float GetFloat(JsonObject? arguments, string key, string paramName, bool required,
        float? defaultValue = null)
    {
        var result = GetFloatNullable(arguments, key, null, paramName);
        if (result.HasValue)
            return result.Value;

        if (required)
            throw new ArgumentException($"{paramName} is required");
        if (defaultValue.HasValue)
            return defaultValue.Value;
        throw new ArgumentException($"{paramName} is required");
    }

    /// <summary>
    ///     Gets a nullable float from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Nullable float value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float? GetFloatNullable(JsonObject? arguments, string key)
    {
        return GetFloatNullable(arguments, key, null, key);
    }

    /// <summary>
    ///     Gets a nullable float from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable float value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float? GetFloatNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetFloatNullable(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets a nullable float from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable float value, or null if both keys are missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static float? GetFloatNullable(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
            return null;

        if (node.GetValueKind() == JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (string.IsNullOrEmpty(str) || !float.TryParse(str, out var result))
                throw new ArgumentException($"{paramName} must be a valid number");
            return result;
        }

        if (node.GetValueKind() == JsonValueKind.Number)
        {
            if (node.AsValue().TryGetValue<float>(out var floatValue)) return floatValue;

            if (node.AsValue().TryGetValue<double>(out var doubleValue)) return (float)doubleValue;

            if (node.AsValue().TryGetValue<int>(out var intValue)) return intValue;

            return node.GetValue<float>();
        }

        throw new ArgumentException($"{paramName} must be a valid number");
    }

    #endregion

    #region GetString - String Methods

    /// <summary>
    ///     Gets a required string value from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>String value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or empty</exception>
    public static string GetString(JsonObject? arguments, string key)
    {
        return GetString(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional string value from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>String value or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if required is true</exception>
    public static string? GetString(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetString(arguments, key) for required parameters");
        return GetStringNullable(arguments, key);
    }

    /// <summary>
    ///     Gets a string value from JSON arguments with a default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="defaultValue">Default value to return if key is missing, null, or empty</param>
    /// <returns>String value or defaultValue</returns>
    public static string GetString(JsonObject? arguments, string key, string defaultValue)
    {
        return GetString(arguments, key, key, false, defaultValue);
    }

    /// <summary>
    ///     Gets an optional string value from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <param name="defaultValue">Default value to return if key is missing or empty (only used when required is false)</param>
    /// <returns>String value, or empty string/defaultValue if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or empty</exception>
    public static string GetString(JsonObject? arguments, string key, string paramName, bool required,
        string? defaultValue = null)
    {
        return GetString(arguments, key, null, paramName, required, defaultValue);
    }

    /// <summary>
    ///     Gets a string value from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <param name="defaultValue">Default value to return if key is missing or empty (only used when required is false)</param>
    /// <returns>String value, or empty string/defaultValue if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or empty</exception>
    public static string GetString(JsonObject? arguments, string primaryName, string? alternateName, string paramName,
        bool required = true, string? defaultValue = null)
    {
        var result = GetStringNullable(arguments, primaryName, alternateName, paramName);
        if (!string.IsNullOrEmpty(result))
            return result;

        if (required)
            throw new ArgumentException($"{paramName} is required");
        return defaultValue ?? string.Empty;
    }

    /// <summary>
    ///     Gets a nullable string value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <returns>Nullable string value, or null if key is missing</returns>
    public static string? GetStringNullable(JsonObject? arguments, string key)
    {
        return GetStringNullable(arguments, key, null, key);
    }

    /// <summary>
    ///     Gets a nullable string value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable string value, or null if key is missing</returns>
    public static string? GetStringNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetStringNullable(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets a nullable string value from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable string value, or null if both keys are missing</returns>
    public static string? GetStringNullable(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        return node?.GetValue<string>();
    }

    #endregion

    #region GetBool - Boolean Methods

    /// <summary>
    ///     Gets a required boolean value from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Boolean value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static bool GetBool(JsonObject? arguments, string key)
    {
        return GetBool(arguments, key, key);
    }

    /// <summary>
    ///     Gets a boolean value from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="defaultValue">Default value to return if key is missing</param>
    /// <returns>Boolean value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static bool GetBool(JsonObject? arguments, string key, bool defaultValue)
    {
        return GetBool(arguments, key, key, defaultValue);
    }

    /// <summary>
    ///     Gets a required boolean value from JSON arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Boolean value (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is invalid</exception>
    public static bool GetBool(JsonObject? arguments, string key, string paramName)
    {
        var result = GetBoolNullable(arguments, key, null, paramName);
        if (!result.HasValue)
            throw new ArgumentException($"{paramName} is required");
        return result.Value;
    }

    /// <summary>
    ///     Gets a boolean value from JSON arguments with default value
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="defaultValue">Default value to return if key is missing</param>
    /// <returns>Boolean value or defaultValue</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static bool GetBool(JsonObject? arguments, string key, string paramName, bool defaultValue)
    {
        var result = GetBoolNullable(arguments, key, null, paramName);
        return result ?? defaultValue;
    }

    /// <summary>
    ///     Gets a nullable boolean value from JSON arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>Nullable boolean value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static bool? GetBoolNullable(JsonObject? arguments, string key)
    {
        return GetBoolNullable(arguments, key, null, key);
    }

    /// <summary>
    ///     Gets a nullable boolean value from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable boolean value, or null if key is missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static bool? GetBoolNullable(JsonObject? arguments, string key, string paramName)
    {
        return GetBoolNullable(arguments, key, null, paramName);
    }

    /// <summary>
    ///     Gets a nullable boolean value from JSON arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>Nullable boolean value, or null if both keys are missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is invalid (but not missing)</exception>
    public static bool? GetBoolNullable(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
            return null;

        if (node.GetValueKind() == JsonValueKind.True || node.GetValueKind() == JsonValueKind.False)
            return node.GetValue<bool>();

        if (node.GetValueKind() == JsonValueKind.String)
        {
            var str = node.GetValue<string>();
            if (bool.TryParse(str, out var result))
                return result;
            throw new ArgumentException($"{paramName} must be a valid boolean");
        }

        if (node.GetValueKind() == JsonValueKind.Number)
        {
            var num = node.GetValue<int>();
            return num != 0;
        }

        throw new ArgumentException($"{paramName} must be a valid boolean");
    }

    #endregion

    #region GetArray - Array Methods

    /// <summary>
    ///     Gets a required JSON array from arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>JsonArray (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is not an array</exception>
    public static JsonArray GetArray(JsonObject? arguments, string key)
    {
        return GetArray(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional JSON array from arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>JsonArray or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is not an array</exception>
    public static JsonArray? GetArray(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetArray(arguments, key) for required parameters");
        return GetArray(arguments, key, key, false);
    }

    /// <summary>
    ///     Gets a required JSON array from arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>JsonArray (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is not an array</exception>
    public static JsonArray GetArray(JsonObject? arguments, string key, string paramName)
    {
        return GetArray(arguments, key, null, paramName) ?? throw new ArgumentException($"{paramName} is required");
    }

    /// <summary>
    ///     Gets an optional JSON array from arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <returns>JsonArray or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is not an array</exception>
    public static JsonArray? GetArray(JsonObject? arguments, string key, string paramName, bool required)
    {
        return GetArray(arguments, key, null, paramName, required);
    }

    /// <summary>
    ///     Gets a JSON array from arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>JsonArray or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is not an array</exception>
    public static JsonArray? GetArray(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName, bool required = true)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
        {
            if (required)
                throw new ArgumentException($"{paramName} is required");
            return null;
        }

        if (node is JsonArray array) return array;

        throw new ArgumentException($"{paramName} must be an array");
    }

    /// <summary>
    ///     Gets an array of integers from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Array of integers or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static int[]? GetIntArray(JsonObject? arguments, string key, string paramName, bool required = true)
    {
        var array = GetArray(arguments, key, paramName, required);
        if (array == null)
            return null;

        try
        {
            return array.Select(x =>
                    x?.GetValue<int>() ?? throw new ArgumentException($"{paramName} contains invalid integer value"))
                .ToArray();
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException($"{paramName} must be an array of integers", ex);
        }
    }

    /// <summary>
    ///     Gets an array of strings from JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>Array of strings or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is invalid</exception>
    public static string[]? GetStringArray(JsonObject? arguments, string key, string paramName, bool required = true)
    {
        var array = GetArray(arguments, key, paramName, required);
        if (array == null)
            return null;

        try
        {
            return array
                .Select(x =>
                    x?.GetValue<string>() ?? throw new ArgumentException($"{paramName} contains invalid string value"))
                .ToArray();
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException($"{paramName} must be an array of strings", ex);
        }
    }

    #endregion

    #region GetObject - Object Methods

    /// <summary>
    ///     Gets a required JSON object from arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <returns>JsonObject (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is not an object</exception>
    public static JsonObject GetObject(JsonObject? arguments, string key)
    {
        return GetObject(arguments, key, key);
    }

    /// <summary>
    ///     Gets an optional JSON object from arguments (simplified version where key is used as paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key (also used as paramName in error messages)</param>
    /// <param name="required">Must be false for this overload</param>
    /// <returns>JsonObject or null if missing</returns>
    /// <exception cref="ArgumentException">Thrown if value is not an object</exception>
    public static JsonObject? GetObject(JsonObject? arguments, string key, bool required)
    {
        if (required)
            throw new ArgumentException("Use GetObject(arguments, key) for required parameters");
        return GetObject(arguments, key, key, false);
    }

    /// <summary>
    ///     Gets a required JSON object from arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <returns>JsonObject (never null)</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is missing or value is not an object</exception>
    public static JsonObject GetObject(JsonObject? arguments, string key, string paramName)
    {
        return GetObject(arguments, key, null, paramName) ?? throw new ArgumentException($"{paramName} is required");
    }

    /// <summary>
    ///     Gets an optional JSON object from arguments (full version with custom paramName)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="key">JSON property key</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required</param>
    /// <returns>JsonObject or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is not an object</exception>
    public static JsonObject? GetObject(JsonObject? arguments, string key, string paramName, bool required)
    {
        return GetObject(arguments, key, null, paramName, required);
    }

    /// <summary>
    ///     Gets a JSON object from arguments, checking multiple parameter names (full version)
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="primaryName">Primary parameter name to check</param>
    /// <param name="alternateName">Alternate parameter name to check (can be null)</param>
    /// <param name="paramName">Parameter name for error messages</param>
    /// <param name="required">Whether the parameter is required (default: true)</param>
    /// <returns>JsonObject or null if not required and missing</returns>
    /// <exception cref="ArgumentException">Thrown if required parameter is missing or value is not an object</exception>
    public static JsonObject? GetObject(JsonObject? arguments, string primaryName, string? alternateName,
        string paramName, bool required = true)
    {
        JsonNode? node = null;
        if (arguments != null)
            node = arguments[primaryName] ?? (alternateName != null ? arguments[alternateName] : null);

        if (node == null)
        {
            if (required)
                throw new ArgumentException($"{paramName} is required");
            return null;
        }

        if (node is JsonObject obj) return obj;

        if (node is JsonValue value && value.GetValueKind() == JsonValueKind.Null)
        {
            if (required)
                throw new ArgumentException($"{paramName} is required");
            return null;
        }

        throw new ArgumentException($"{paramName} must be an object");
    }

    #endregion

    #region Path Validation Methods

    /// <summary>
    ///     Gets output path from arguments, defaulting to input path if not specified
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="inputPath">Input file path to use as default</param>
    /// <returns>Output path or inputPath if outputPath is not specified</returns>
    public static string GetOutputPath(JsonObject? arguments, string inputPath)
    {
        return arguments?["outputPath"]?.GetValue<string>() ?? inputPath;
    }

    /// <summary>
    ///     Gets and validates input path from arguments
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
    ///     Gets and validates output path from arguments, defaulting to input path if not specified
    /// </summary>
    /// <param name="arguments">JSON arguments object</param>
    /// <param name="inputPath">Input file path</param>
    /// <param name="paramName">Parameter name for error messages (default: "outputPath")</param>
    /// <returns>Validated output path</returns>
    /// <exception cref="ArgumentException">Thrown if output path is invalid</exception>
    public static string GetAndValidateOutputPath(JsonObject? arguments, string inputPath,
        string paramName = "outputPath")
    {
        var outputPath = arguments?[paramName]?.GetValue<string>() ?? inputPath;
        SecurityHelper.ValidateFilePath(outputPath, paramName);
        return outputPath;
    }

    #endregion
}