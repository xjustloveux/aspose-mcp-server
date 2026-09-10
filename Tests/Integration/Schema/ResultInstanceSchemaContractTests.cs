using System.Collections;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tests.Integration.Schema;

/// <summary>
///     Covers TEST-06. The published outputSchema is generated from the result types and sets
///     <c>additionalProperties: false</c>, so a property a result actually serialises but the
///     schema does not declare makes a strict client reject the response. Existing tests checked
///     the schema's shape; nothing checked that a real result instance satisfies it. These tests
///     build a populated instance of every result type, serialise it the way the schema generator
///     reads the type, and hold the two against each other.
/// </summary>
[Trait("Category", "Integration")]
public class ResultInstanceSchemaContractTests
{
    /// <summary>
    ///     Lowest number of result types the discovery must keep finding. The repository
    ///     declares 125; the floor sits just below so ordinary additions do not trip it while a
    ///     discovery that breaks and silently returns nothing does.
    /// </summary>
    private const int MinimumResultTypes = 120;

    private static readonly Assembly TargetAssembly = typeof(OutputSchemaGenerator).Assembly;

    private static readonly JsonSerializerOptions CamelCase = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>Collects the distinct result types declared by handlers.</summary>
    /// <returns>Every type named by a <see cref="ResultTypeAttribute" />.</returns>
    private static List<Type> ResultTypes()
    {
        return TargetAssembly.GetTypes()
            .Where(t => t is { IsClass: true, IsAbstract: false })
            .Select(t => t.GetCustomAttribute<ResultTypeAttribute>(true))
            .Where(a => a != null)
            .Select(a => a!.ResultType)
            .Distinct()
            .OrderBy(t => t.FullName, StringComparer.Ordinal)
            .ToList();
    }

    [Fact]
    public void ResultTypeDiscovery_ShouldStillFindTheHandlerResults()
    {
        var discovered = ResultTypes();

        Assert.True(discovered.Count >= MinimumResultTypes,
            $"Only {discovered.Count} result types were discovered, below the floor of "
            + $"{MinimumResultTypes}. The contract checks below iterate this list, so a broken "
            + "discovery would make them pass without checking anything.");
    }

    [Fact]
    public void EverySerializedProperty_ShouldBeDeclaredInTheSchema()
    {
        var failures = new List<string>();

        foreach (var resultType in ResultTypes())
        {
            var instance = TryCreatePopulatedInstance(resultType, 0);
            if (instance == null)
            {
                failures.Add($"{resultType.Name}: could not be instantiated for the contract check");
                continue;
            }

            var serialized = JsonNode.Parse(JsonSerializer.Serialize(instance, resultType, CamelCase))
                ?.AsObject();
            Assert.NotNull(serialized);

            var dataSchema = DataSchemaFor(resultType);
            if (dataSchema?["properties"] is not JsonObject declared)
            {
                failures.Add($"{resultType.Name}: schema has no properties object");
                continue;
            }

            failures.AddRange(serialized
                .Select(property => property.Key)
                .Where(name => !declared.ContainsKey(name))
                .Select(name =>
                    $"{resultType.Name}: serialises {name}, which the schema does not declare "
                    + "while additionalProperties is false"));
        }

        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    [Fact]
    public void EveryRequiredSchemaProperty_ShouldBeSerialized()
    {
        var failures = new List<string>();

        foreach (var resultType in ResultTypes())
        {
            var instance = TryCreatePopulatedInstance(resultType, 0);
            if (instance == null) continue;

            var serialized = JsonNode.Parse(JsonSerializer.Serialize(instance, resultType, CamelCase))
                ?.AsObject();
            if (serialized == null) continue;

            var dataSchema = DataSchemaFor(resultType);
            if (dataSchema?["required"] is not JsonArray required) continue;

            failures.AddRange(required
                .Select(node => node?.GetValue<string>())
                .Where(name => name != null && !serialized.ContainsKey(name))
                .Select(name =>
                    $"{resultType.Name}: schema requires {name}, which a populated instance does not emit"));
        }

        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>Extracts the data branch of the generated FinalizedResult schema.</summary>
    /// <param name="resultType">The result type to generate a schema for.</param>
    /// <returns>The data schema node, or null when it is missing.</returns>
    private static JsonNode? DataSchemaFor(Type resultType)
    {
        var schema = OutputSchemaGenerator.GenerateForTypes([resultType]);
        return JsonNode.Parse(schema.GetRawText())?["properties"]?["data"];
    }

    /// <summary>
    ///     Creates an instance with every writable property filled, so the serialisation exercises
    ///     the whole surface rather than the defaults.
    /// </summary>
    /// <param name="type">The type to instantiate.</param>
    /// <param name="depth">Current recursion depth, bounded to avoid cyclic models.</param>
    /// <returns>The populated instance, or null when the type cannot be constructed.</returns>
    private static object? TryCreatePopulatedInstance(Type type, int depth)
    {
        if (depth > 3) return null;
        if (type.GetConstructor(Type.EmptyTypes) == null) return null;

        object instance;
        try
        {
            instance = Activator.CreateInstance(type)!;
        }
        catch (Exception)
        {
            return null;
        }

        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanWrite || property.GetIndexParameters().Length > 0) continue;

            var value = SampleValue(property.PropertyType, depth);
            if (value == null) continue;

            try
            {
                property.SetValue(instance, value);
            }
            catch (Exception)
            {
                // A property that rejects the sample keeps its default; the schema check still runs.
            }
        }

        return instance;
    }

    /// <summary>Produces a representative value for a property type.</summary>
    /// <param name="type">The property type.</param>
    /// <param name="depth">Current recursion depth.</param>
    /// <returns>A sample value, or null when none can be produced.</returns>
    private static object? SampleValue(Type type, int depth)
    {
        var target = Nullable.GetUnderlyingType(type) ?? type;

        if (target == typeof(string)) return "sample";
        if (target == typeof(bool)) return true;
        if (target == typeof(int)) return 1;
        if (target == typeof(long)) return 1L;
        if (target == typeof(double)) return 1.5d;
        if (target == typeof(float)) return 1.5f;
        if (target == typeof(decimal)) return 1.5m;
        if (target == typeof(DateTime)) return new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        if (target == typeof(DateTimeOffset)) return DateTimeOffset.UnixEpoch;
        if (target.IsEnum) return Enum.GetValues(target).GetValue(0);
        if (target == typeof(object)) return "sample";

        if (target.IsArray)
        {
            var elementType = target.GetElementType()!;
            var element = SampleValue(elementType, depth + 1);
            var array = Array.CreateInstance(elementType, element == null ? 0 : 1);
            if (element != null) array.SetValue(element, 0);
            return array;
        }

        if (target.IsGenericType)
        {
            var definition = target.GetGenericTypeDefinition();
            if (definition == typeof(List<>) || definition == typeof(IList<>) ||
                definition == typeof(IReadOnlyList<>) || definition == typeof(IEnumerable<>) ||
                definition == typeof(ICollection<>) || definition == typeof(IReadOnlyCollection<>))
            {
                var elementType = target.GetGenericArguments()[0];
                var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(elementType))!;
                var element = SampleValue(elementType, depth + 1);
                if (element != null) list.Add(element);
                return list;
            }

            if (definition == typeof(Dictionary<,>) || definition == typeof(IDictionary<,>) ||
                definition == typeof(IReadOnlyDictionary<,>))
            {
                var arguments = target.GetGenericArguments();
                var dictionary = (IDictionary)Activator.CreateInstance(
                    typeof(Dictionary<,>).MakeGenericType(arguments[0], arguments[1]))!;
                var key = SampleValue(arguments[0], depth + 1);
                var value = SampleValue(arguments[1], depth + 1);
                if (key != null && value != null) dictionary.Add(key, value);
                return dictionary;
            }
        }

        return target.IsClass ? TryCreatePopulatedInstance(target, depth + 1) : null;
    }
}
