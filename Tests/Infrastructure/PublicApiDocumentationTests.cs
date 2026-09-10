using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Keeps the compiler's own documentation checks switched on.
///     <para>
///         Signatures gained parameters without their documentation following: the PowerPoint
///         media, shape and table tools, the Excel cell tool and the external-reference scanner all
///         took arguments that no <c>param</c> tag mentioned (R3-DOC10). Nothing failed, because
///         CS1573 and CS1591 are only raised when the project generates a documentation file, and
///         this one did not.
///     </para>
///     <para>
///         It does now, and the build runs at zero warnings, so the compiler is the guard: it
///         checks every member, resolves overloads exactly and understands <c>inheritdoc</c>, none
///         of which a test scanning sources or reflection metadata does well. What a test can add
///         is making sure the guard cannot be turned off quietly — by clearing the property, or by
///         adding these warning numbers to <c>NoWarn</c> — and that the file it produces is real.
///     </para>
/// </summary>
public class PublicApiDocumentationTests
{
    /// <summary>
    ///     Locates the repository root by walking up to the production project file.
    /// </summary>
    /// <returns>The repository root directory.</returns>
    private static string RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir.FullName;
    }

    /// <summary>Reads the production project file.</summary>
    /// <returns>Its full text.</returns>
    private static string ProjectFile()
    {
        return File.ReadAllText(Path.Combine(RepositoryRoot(), "AsposeMcpServer.csproj"), Encoding.UTF8);
    }

    [Fact]
    public void TheProject_ShouldGenerateItsDocumentationFile()
    {
        var project = ProjectFile();

        Assert.Matches(new Regex(@"<GenerateDocumentationFile>\s*true\s*</GenerateDocumentationFile>",
            RegexOptions.IgnoreCase), project);
    }

    /// <summary>
    ///     CS1591 is a missing member comment and CS1573 is a missing parameter comment. Both are
    ///     the point of generating the file at all, so suppressing either would leave the property
    ///     set and the check gone.
    /// </summary>
    /// <param name="warning">A warning number that must not be suppressed.</param>
    [Theory]
    [InlineData("1591")]
    [InlineData("1573")]
    public void TheDocumentationWarnings_ShouldNotBeSuppressed(string warning)
    {
        var project = ProjectFile();

        foreach (Match block in Regex.Matches(project, "<NoWarn>(.*?)</NoWarn>", RegexOptions.Singleline))
        {
            var suppressed = block.Groups[1].Value
                .Split([';', ',', ' ', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim().TrimStart('C', 'S', 'c', 's'));

            Assert.DoesNotContain(warning, suppressed);
        }
    }

    /// <summary>
    ///     The property could be set and still produce nothing useful — a wrong output path, or a
    ///     build that never ran — so the file that ships is checked too.
    /// </summary>
    [Fact]
    public void TheGeneratedDocumentation_ShouldShipBesideTheAssembly()
    {
        var assembly = typeof(RenderBudget).Assembly;
        var xmlPath = Path.ChangeExtension(assembly.Location, ".xml");

        Assert.True(File.Exists(xmlPath), $"No XML documentation beside {assembly.Location}.");

        var members = XDocument.Load(xmlPath).Descendants("member")
            .Select(m => m.Attribute("name")?.Value ?? string.Empty)
            .Where(n => n.Length > 0)
            .ToList();

        Assert.True(members.Count > 200,
            $"Only {members.Count} documented members were found; the file looks wrong.");
        Assert.Contains(members,
            n => n.StartsWith("M:AsposeMcpServer.Helpers.RenderBudget.", StringComparison.Ordinal));
    }

    /// <summary>
    ///     Every documented method that returns something must say what.
    ///     <para>
    ///         CS1591 and CS1573 cover a missing member comment and a missing parameter tag, and
    ///         nothing covers <c>returns</c>: a signature that changed from <c>void</c> to a value
    ///         kept its old comment, so the value it now hands back was undocumented and callers
    ///         had no reason to think it mattered — which is exactly how a resolved path came to
    ///         be thrown away at two call sites (R3-S03, R4-DOC03).
    ///     </para>
    /// </summary>
    [Fact]
    public void EveryDocumentedMethodThatReturnsAValue_ShouldSayWhatItReturns()
    {
        var assembly = typeof(RenderBudget).Assembly;
        var xmlPath = Path.ChangeExtension(assembly.Location, ".xml");
        var document = XDocument.Load(xmlPath);

        var undocumented = new List<string>();
        var unresolved = new List<string>();

        foreach (var member in document.Descendants("member"))
        {
            var name = member.Attribute("name")?.Value;
            if (name == null || !name.StartsWith("M:", StringComparison.Ordinal)) continue;

            // Inherited documentation carries the base member's returns.
            if (member.Elements("inheritdoc").Any()) continue;
            if (member.Elements("returns").Any()) continue;

            // Constructors and anything the reflection pass cannot match are out of scope; only a
            // resolvable, value-returning method is asked for a returns tag.
            if (name.Contains("#ctor", StringComparison.Ordinal)) continue;

            // The regex source generator emits its helper types as file-local, so their metadata
            // names are mangled and nothing in the key can be matched to them. Their documentation
            // is the generator's own, as the exclusion below says.
            if (name.StartsWith("M:System.Text.RegularExpressions.Generated.", StringComparison.Ordinal))
                continue;

            var method = ResolveMethod(assembly, name);
            if (method == null)
            {
                unresolved.Add(name[2..]);
                continue;
            }

            if (method.ReturnType == typeof(void)) continue;

            // Source-generated members carry documentation the generator wrote, not this
            // repository: [GeneratedRegex] emits the implementing half of a partial method and
            // the emitted XML is what ships. There is no hand-written contract to check.
            if (method.GetCustomAttributes(false)
                .Any(a => a.GetType().Name.StartsWith("Generated", StringComparison.Ordinal)))
                continue;

            undocumented.Add(name[2..]);
        }

        // A key this test cannot match to a member is skipped, so a resolver that quietly gives up
        // makes the whole check vacuous — which is what returning null for every overload set did.
        Assert.True(unresolved.Count == 0,
            "These documentation keys could not be matched to a member, so nothing about them was " +
            "checked:" + Environment.NewLine +
            string.Join(Environment.NewLine, unresolved.Order(StringComparer.Ordinal)));

        Assert.True(undocumented.Count == 0,
            "These documented methods return a value but do not say what it is:" +
            Environment.NewLine + string.Join(Environment.NewLine, undocumented.Order(StringComparer.Ordinal)));
    }

    /// <summary>
    ///     Text known to have been copied from another member is not a contract.
    ///     <para>
    ///         Seven members described what they return as <c>The string found, in order.</c> — text
    ///         belonging to a search helper, pasted onto document converters and a path list. The
    ///         compiler cannot tell a wrong contract from a right one, and the tag-presence check
    ///         above counts it as documented, so a regression that pastes it back is caught here
    ///         instead (R4-DOC03).
    ///     </para>
    /// </summary>
    [Fact]
    public void NoDocumentedContract_ShouldRepeatKnownCopiedText()
    {
        string[] copied = ["The string found, in order."];

        var assembly = typeof(RenderBudget).Assembly;
        var document = XDocument.Load(Path.ChangeExtension(assembly.Location, ".xml"));

        var offenders = new List<string>();

        foreach (var member in document.Descendants("member"))
        {
            var name = member.Attribute("name")?.Value;
            if (name == null) continue;

            foreach (var tag in member.Elements("returns"))
            {
                var text = string.Join(" ", tag.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

                if (text.Length == 0)
                    offenders.Add($"{name[2..]}: empty returns");
                else if (copied.Contains(text, StringComparer.Ordinal))
                    offenders.Add($"{name[2..]}: \"{text}\"");
            }
        }

        Assert.True(offenders.Count == 0,
            "These members carry text copied from an unrelated member instead of their own contract:" +
            Environment.NewLine + string.Join(Environment.NewLine, offenders.Order(StringComparer.Ordinal)));
    }

    /// <summary>
    ///     Finds the method a documentation key names.
    ///     <para>
    ///         The key carries the parameter list, so an overload set is told apart by decoding it.
    ///         Returning <c>null</c> for every name with more than one overload skipped the whole
    ///         set instead of one member of it, and an undocumented return value anywhere in an
    ///         overload set went unnoticed (R4-DOC03).
    ///     </para>
    /// </summary>
    /// <param name="assembly">Assembly to search.</param>
    /// <param name="documentationKey">A key of the form <c>M:Namespace.Type.Method(args)</c>.</param>
    /// <returns>The method, or <c>null</c> when the key names something this assembly does not declare.</returns>
    private static MethodInfo? ResolveMethod(Assembly assembly, string documentationKey)
    {
        var withoutPrefix = documentationKey[2..];

        // A conversion operator's key appends the return type after a tilde; it is not part of
        // the name or of the parameter list.
        var conversion = withoutPrefix.IndexOf('~');
        if (conversion >= 0) withoutPrefix = withoutPrefix[..conversion];

        var signatureStart = withoutPrefix.IndexOf('(');
        var fullName = signatureStart < 0 ? withoutPrefix : withoutPrefix[..signatureStart];
        var signature = signatureStart < 0
            ? []
            : SplitArguments(withoutPrefix[(signatureStart + 1)..^1]);

        var lastDot = fullName.LastIndexOf('.');
        if (lastDot < 0) return null;

        var typeName = fullName[..lastDot];
        var methodName = fullName[(lastDot + 1)..];

        // A generic method's key states its arity after the name.
        var arity = 0;
        var backtick = methodName.IndexOf('`');
        if (backtick >= 0)
        {
            arity = int.Parse(methodName[(backtick + 1)..].TrimStart('`'),
                CultureInfo.InvariantCulture);
            methodName = methodName[..backtick];
        }

        var type = FindType(assembly, typeName);
        if (type == null) return null;

        var candidates = type
            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance |
                        BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Where(m => m.Name == methodName)
            .Where(m => m.GetGenericArguments().Length == arity)
            .Where(m => m.GetParameters().Select(p => DocumentationId(p.ParameterType))
                .SequenceEqual(signature, StringComparer.Ordinal))
            .ToList();

        return candidates.Count == 1 ? candidates[0] : null;
    }

    /// <summary>
    ///     Finds the type a documentation key names.
    /// </summary>
    /// <param name="assembly">Assembly to search.</param>
    /// <param name="typeName">
    ///     The dotted name from the key. A key separates a nested type from the type that declares
    ///     it with a dot, exactly as it separates a namespace, so which dots are nesting is found
    ///     by trying them from the right.
    /// </param>
    /// <returns>The type, or <c>null</c> when this assembly declares no such type.</returns>
    private static Type? FindType(Assembly assembly, string typeName)
    {
        var candidate = typeName;
        while (true)
        {
            var found = assembly.GetType(candidate);
            if (found != null) return found;

            var lastDot = candidate.LastIndexOf('.');
            if (lastDot < 0) return null;
            candidate = string.Concat(candidate[..lastDot], "+", candidate[(lastDot + 1)..]);
        }
    }

    /// <summary>
    ///     Splits a documentation key's argument list on its top-level commas.
    /// </summary>
    /// <param name="arguments">The text between the key's parentheses.</param>
    /// <returns>One entry per parameter, generic arguments kept with the type that takes them.</returns>
    private static List<string> SplitArguments(string arguments)
    {
        var parts = new List<string>();
        if (arguments.Length == 0) return parts;

        var depth = 0;
        var start = 0;
        for (var i = 0; i < arguments.Length; i++)
            switch (arguments[i])
            {
                case '{' or '[':
                    depth++;
                    break;
                case '}' or ']':
                    depth--;
                    break;
                case ',' when depth == 0:
                    parts.Add(arguments[start..i]);
                    start = i + 1;
                    break;
            }

        parts.Add(arguments[start..]);
        return parts;
    }

    /// <summary>
    ///     Writes a type the way a documentation key spells it.
    /// </summary>
    /// <param name="type">The parameter type.</param>
    /// <returns>Its documentation-comment identifier.</returns>
    private static string DocumentationId(Type type)
    {
        if (type.IsByRef) return DocumentationId(type.GetElementType()!) + "@";
        if (type.IsPointer) return DocumentationId(type.GetElementType()!) + "*";

        if (type.IsArray)
        {
            var rank = type.GetArrayRank();
            var dimensions = rank == 1 ? "[]" : "[" + string.Join(",", Enumerable.Repeat("0:", rank)) + "]";
            return DocumentationId(type.GetElementType()!) + dimensions;
        }

        // A generic parameter is written by position: one backtick for the type's, two for the
        // method's own.
        if (type.IsGenericParameter)
            return (type.DeclaringMethod == null ? "`" : "``") + type.GenericParameterPosition;

        // FullName is null for a type that still contains generic parameters, which is every
        // parameter type of a generic method or of a method on a generic type, so the name is
        // built rather than read.
        var name = QualifiedName(type);

        if (!type.IsGenericType) return name;

        var arity = name.IndexOf('`');
        if (arity >= 0) name = name[..arity];
        return name + "{" + string.Join(",", type.GetGenericArguments().Select(DocumentationId)) + "}";
    }

    /// <summary>
    ///     Writes a type's namespace-qualified name, nested types separated by dots as a
    ///     documentation key writes them.
    /// </summary>
    /// <param name="type">The type to name.</param>
    /// <returns>Its qualified name, generic arity suffix included.</returns>
    private static string QualifiedName(Type type)
    {
        if (type.DeclaringType != null)
            return QualifiedName(type.DeclaringType) + "." + type.Name;

        return string.IsNullOrEmpty(type.Namespace) ? type.Name : type.Namespace + "." + type.Name;
    }
}
