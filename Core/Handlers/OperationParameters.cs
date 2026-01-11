using System.Text.Json;

namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Encapsulates all parameters passed to an operation handler.
///     Provides type-safe access to common and operation-specific parameters
///     with support for nullable types and JSON element conversion.
/// </summary>
/// <remarks>
///     <para>
///         This class serves as a thin wrapper over a dictionary, providing
///         convenient methods for parameter retrieval with proper type conversion.
///     </para>
///     <para>
///         Supports the following type conversions:
///         - Primitive types (int, double, bool, string)
///         - Nullable types (int?, double?, bool?, etc.)
///         - Enum types
///         - JsonElement values from MCP requests
///     </para>
/// </remarks>
public class OperationParameters
{
    private readonly Dictionary<string, object?> _values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    ///     Sets a parameter value.
    /// </summary>
    /// <param name="name">The parameter name (case-insensitive).</param>
    /// <param name="value">The parameter value.</param>
    public void Set(string name, object? value)
    {
        _values[name] = value;
    }

    /// <summary>
    ///     Gets a required parameter value.
    /// </summary>
    /// <typeparam name="T">The expected type.</typeparam>
    /// <param name="name">The parameter name.</param>
    /// <returns>The parameter value converted to type T.</returns>
    /// <exception cref="ArgumentException">Thrown when the parameter is missing or null.</exception>
    public T GetRequired<T>(string name)
    {
        if (!_values.TryGetValue(name, out var value) || value == null)
            throw new ArgumentException($"{name} is required");

        return ConvertValue<T>(value, name);
    }

    /// <summary>
    ///     Gets an optional parameter value with a default.
    /// </summary>
    /// <typeparam name="T">The expected type.</typeparam>
    /// <param name="name">The parameter name.</param>
    /// <param name="defaultValue">The default value if parameter is missing or null.</param>
    /// <returns>The parameter value or default.</returns>
    public T GetOptional<T>(string name, T defaultValue = default!)
    {
        if (!_values.TryGetValue(name, out var value) || value == null)
            return defaultValue;

        return ConvertValue<T>(value, name);
    }

    /// <summary>
    ///     Checks if a parameter exists and is not null.
    /// </summary>
    /// <param name="name">The parameter name.</param>
    /// <returns>True if the parameter exists and is not null.</returns>
    public bool Has(string name)
    {
        return _values.TryGetValue(name, out var value) && value != null;
    }

    /// <summary>
    ///     Gets the raw value of a parameter without conversion.
    /// </summary>
    /// <param name="name">The parameter name.</param>
    /// <returns>The raw value or null if not found.</returns>
    public object? GetRaw(string name)
    {
        _values.TryGetValue(name, out var value);
        return value;
    }

    /// <summary>
    ///     Converts a value to the target type, handling nullable types,
    ///     enums, and JsonElement values.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="value">The value to convert.</param>
    /// <param name="parameterName">The parameter name for error messages.</param>
    /// <returns>The converted value.</returns>
    private static T ConvertValue<T>(object value, string parameterName)
    {
        var targetType = typeof(T);
        var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

        if (value is JsonElement jsonElement)
            return ConvertJsonElement<T>(jsonElement, underlyingType, parameterName);

        if (targetType.IsInstanceOfType(value))
            return (T)value;

        if (underlyingType.IsEnum)
            return ConvertToEnum<T>(value, underlyingType, parameterName);

        try
        {
            var converted = Convert.ChangeType(value, underlyingType);
            return (T)converted;
        }
        catch (Exception ex) when (ex is InvalidCastException or FormatException or OverflowException)
        {
            throw new ArgumentException(
                $"Cannot convert parameter '{parameterName}' value '{value}' to type {targetType.Name}", ex);
        }
    }

    /// <summary>
    ///     Converts a JsonElement to the target type.
    /// </summary>
    private static T ConvertJsonElement<T>(JsonElement element, Type underlyingType, string parameterName)
    {
        try
        {
            if (underlyingType == typeof(string))
                return (T)(object)element.GetString()!;

            if (underlyingType == typeof(int))
                return (T)(object)element.GetInt32();

            if (underlyingType == typeof(long))
                return (T)(object)element.GetInt64();

            if (underlyingType == typeof(double))
                return (T)(object)element.GetDouble();

            if (underlyingType == typeof(float))
                return (T)(object)(float)element.GetDouble();

            if (underlyingType == typeof(bool))
                return (T)(object)element.GetBoolean();

            if (underlyingType == typeof(decimal))
                return (T)(object)element.GetDecimal();

            if (underlyingType.IsEnum)
            {
                if (element.ValueKind == JsonValueKind.String)
                {
                    var stringValue = element.GetString();
                    if (stringValue != null && Enum.TryParse(underlyingType, stringValue, true, out var enumValue))
                        return (T)enumValue;
                }
                else if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out var intValue))
                {
                    return (T)Enum.ToObject(underlyingType, intValue);
                }
            }

            return element.Deserialize<T>()!;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Cannot convert JSON parameter '{parameterName}' to type {typeof(T).Name}", ex);
        }
    }

    /// <summary>
    ///     Converts a value to an enum type.
    /// </summary>
    private static T ConvertToEnum<T>(object value, Type enumType, string parameterName)
    {
        if (value is string stringValue)
        {
            if (Enum.TryParse(enumType, stringValue, true, out var result))
                return (T)result;

            throw new ArgumentException(
                $"Cannot convert parameter '{parameterName}' value '{stringValue}' to enum {enumType.Name}");
        }

        try
        {
            return (T)Enum.ToObject(enumType, value);
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Cannot convert parameter '{parameterName}' value '{value}' to enum {enumType.Name}", ex);
        }
    }
}
