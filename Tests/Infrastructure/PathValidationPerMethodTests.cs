using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Covers A-02's meta-test gap. <see cref="PathValidationCoverageTests" /> asks whether a
///     handler <em>file</em> contains any validation call at all, so one validated path made every
///     other sink in the same file look covered — which is how several handlers kept a lexical-only
///     input path while their output path was resolved. This checks the same rule one method at a
///     time, and follows the value rather than the vocabulary: the path a sink actually receives
///     must trace back to the allowlist resolver or to a directory the server itself owns. Asking
///     only whether the word "resolved" appeared somewhere in the method passed any method that
///     resolved one path and then wrote through a second, still-raw one (R7-T01).
/// </summary>
public class PathValidationPerMethodTests
{
    /// <summary>
    ///     Calls that touch the filesystem with a path the caller supplied. Every entry is proven
    ///     detectable by <see cref="EverySinkKind_ShouldBeFlaggedWhenItsPathWasNeverResolved" />, so
    ///     a pattern that no longer matches anything cannot sit here looking like coverage.
    /// </summary>
    private static readonly string[] SinkPatterns =
    [
        "File.Copy(", "File.Delete(", "File.Move(", "File.Replace(",
        "File.WriteAllText(", "File.WriteAllBytes(", "File.WriteAllLines(",
        "File.AppendAllText(", "File.ReadAllText(", "File.ReadAllBytes(",
        "File.ReadAllLines(", "File.Open(", "File.OpenRead(", "File.OpenWrite(",
        "File.Create(", "new FileStream(", "Directory.Delete(",
        "Directory.CreateDirectory(", "Directory.GetFiles(", "ZipFile.OpenRead(",
        "ZipFile.CreateFromDirectory(", "ZipFile.ExtractToDirectory(", ".Save("
    ];

    /// <summary>
    ///     Which argument positions of a sink carry a path, for the ones that do not simply take
    ///     one first.
    ///     <para>
    ///         Every sink but <c>.Save(</c> used to be judged on argument zero alone, so
    ///         <c>File.Copy(resolved, raw)</c> passed: the first argument was resolved, and the
    ///         destination — the one being written — was never looked at. The negative fixtures all
    ///         put the raw path first, so nothing caught it (R17-T02).
    ///     </para>
    ///     <para>
    ///         An entry of <c>null</c> means every argument carries a path, which is how
    ///         <c>.Save(</c> has always been treated: it has a stream overload, and which argument
    ///         holds the path differs between them.
    ///     </para>
    /// </summary>
    private static readonly Dictionary<string, int[]?> PathArguments = new(StringComparer.Ordinal)
    {
        ["File.Copy("] = [0, 1],
        ["File.Move("] = [0, 1],
        ["File.Replace("] = [0, 1, 2],

        // Both ends of a ZIP operation are paths: the directory read from and the archive
        // written, or the archive read and the directory written into. Listed in SinkPatterns
        // and missing here, so they fell back to argument zero — no handler calls them today,
        // which makes this a guard that would have been wrong for the first one that did
        // (R18-T02).
        ["ZipFile.CreateFromDirectory("] = [0, 1],
        ["ZipFile.ExtractToDirectory("] = [0, 1],

        [".Save("] = null
    };

    /// <summary>
    ///     Library types that open a path, and the methods of theirs that take one.
    ///     <para>
    ///         Everything in <see cref="SinkPatterns" /> is a BCL file API or <c>.Save(</c>, so a
    ///         path handed to a vendor type was not a sink at all: the OCR preprocessing handlers
    ///         passed the caller's own spelling to <c>OcrInput.Add</c> and nothing here objected
    ///         (R19-OCR01). The blind spot was the inventory, not the parser -- the method-start
    ///         pattern matches those wrapped signatures perfectly well.
    ///     </para>
    ///     <para>
    ///         Keyed by the constructed type, because the call is written on a variable
    ///         (<c>input.Add(path)</c>) and only the construction says what that variable is.
    ///     </para>
    /// </summary>
    private static readonly Dictionary<string, string[]> LibraryPathSinks =
        new(StringComparer.Ordinal)
        {
            ["OcrInput"] = ["Add"]
        };

    /// <summary>A resolver call that is the value of an assignment.</summary>
    private static readonly Regex ResolverExpression = new(
        @"^\s*(?:[A-Za-z_][A-Za-z0-9_]*\s*\.\s*)*"
        + @"(?:ResolveAndEnsureWithinAllowlist|ValidateUserPath)\s*\(",
        RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>
    ///     A path the server minted for itself. These carry no caller-supplied component, so the
    ///     allowlist has nothing to check them against.
    /// </summary>
    private static readonly string[] ServerOwnedOrigins =
    [
        "Path.GetTempPath()", "Path.GetRandomFileName()", "Path.GetTempFileName()",
        "TempFileManager", "AppContext.BaseDirectory"
    ];

    /// <summary>A local holding a stream, so a Save through it is not a path sink.</summary>
    private static readonly Regex StreamDeclaration = new(
        @"(?:var|[A-Za-z_][A-Za-z0-9_]*Stream)\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*"
        + @"(?:new\s+[A-Za-z_][A-Za-z0-9_]*Stream|File\.(?:Create|Open|OpenRead|OpenWrite)\()",
        RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>A lambda parameter, which is how the bounded publishers hand a stream to a sink.</summary>
    private static readonly Regex LambdaParameter = new(
        @"([A-Za-z_][A-Za-z0-9_]*)\s*=>", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>A string literal, ordinary or verbatim, so its contents are never read as names.</summary>
    private static readonly Regex StringLiteral = new(
        @"@""(?:[^""]|"""")*""|""(?:[^""\\]|\\.)*""",
        RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>
    ///     An identifier that is neither a member being accessed nor a call: a bare variable, or the
    ///     receiver of a member access.
    /// </summary>
    /// <remarks>
    ///     Receivers count. `request` in `request.UserInput` was excluded as "only a receiver",
    ///     and `Path.Combine(resolvedRoot, request.UserInput)` then held one known variable,
    ///     trusted, and was trusted whole (R22-TST01). What a member carries is whatever its
    ///     receiver carries, until a resolver says otherwise. Types (`Path`) and methods
    ///     (`Combine`) are still not variables: the caller keeps only names the method knows.
    /// </remarks>
    private static readonly Regex VariableOrReceiver = new(
        @"(?<![\w.])([A-Za-z_][A-Za-z0-9_]*)(?!\s*\()", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>Any identifier, used to read the names out of an argument expression.</summary>
    private static readonly Regex Identifier = new(
        "[A-Za-z_][A-Za-z0-9_]*", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>Matches a method signature well enough to find where each body starts.</summary>
    private static readonly Regex MethodStart = new(
        @"^\s*(?:\[[^\]]*\]\s*)*(?:public|private|protected|internal)\s+[^;{}()]*\([^;{}]*\)\s*$",
        RegexOptions.Compiled | RegexOptions.Multiline, TimeSpan.FromSeconds(10));

    /// <summary>A `return` of a single name, which is how a resolving helper hands back its result.</summary>
    private static readonly Regex ReturnedName = new(
        @"\breturn\s+([A-Za-z_][A-Za-z0-9_]*)\s*;", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>An interpolation hole, whose identifier is as much a leaf as one outside a string.</summary>
    private static readonly Regex InterpolationHole = new(
        @"\{([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>The argument positions of a sink that carry a path.</summary>
    /// <param name="sink">The sink call prefix.</param>
    /// <param name="arguments">Its arguments as written.</param>
    /// <returns>The arguments to judge.</returns>
    private static List<string> PathBearing(string sink, List<string> arguments)
    {
        if (!PathArguments.TryGetValue(sink, out var positions)) return arguments.Take(1).ToList();
        if (positions == null) return arguments;

        return positions.Where(index => index < arguments.Count)
            .Select(index => arguments[index]).ToList();
    }

    /// <summary>Finds a receiver built from one of the library types above.</summary>
    /// <param name="type">The constructed type name.</param>
    /// <returns>A pattern capturing the variable it was assigned to.</returns>
    private static Regex LibraryReceiver(string type)
    {
        return new Regex(
            "(?:var|" + Regex.Escape(type) + @")\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*new\s+"
            + Regex.Escape(type) + @"\s*\(",
            RegexOptions.Compiled, TimeSpan.FromSeconds(10));
    }

    /// <summary>The library sink calls one method body actually makes.</summary>
    /// <param name="body">The method body.</param>
    /// <returns>Call prefixes such as <c>input.Add(</c>, ready to scan for like any other sink.</returns>
    private static List<string> LibrarySinksIn(string body)
    {
        var sinks = new List<string>();

        foreach (var (type, methods) in LibraryPathSinks)
        foreach (Match match in LibraryReceiver(type).Matches(body))
            sinks.AddRange(methods.Select(method => $"{match.Groups[1].Value}.{method}("));

        return sinks;
    }

    /// <summary>
    ///     Walks up to the repository root, identified by the project file. Looking for a
    ///     directory named "Handlers" instead would stop at Tests/, which has one of its own.
    /// </summary>
    /// <returns>The production Handlers directory.</returns>
    private static string HandlersRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        var handlers = Path.Combine(dir.FullName, "Handlers");
        Assert.True(Directory.Exists(handlers), $"Production handlers not found at {handlers}");
        return handlers;
    }

    /// <summary>
    ///     Splits a source file into method bodies by brace matching, which is enough for this
    ///     codebase's formatting and avoids taking a compiler dependency on a meta-test.
    /// </summary>
    /// <param name="source">The file's text.</param>
    /// <returns>Each method body found, in order.</returns>
    private static List<string> MethodBodies(string source)
    {
        var bodies = new List<string>();

        foreach (Match match in MethodStart.Matches(source))
        {
            var open = source.IndexOf('{', match.Index + match.Length);
            if (open < 0) continue;

            // The parameter list, kept with the body so the analyser can tell a parameter from
            // a local (R21-TST02). Written as a comment the body's own scan ignores.
            var signature = match.Value;
            var parameters = signature[(signature.IndexOf('(') + 1)..signature.LastIndexOf(')')];
            var name = Identifier.Matches(signature[..signature.IndexOf('(')]).LastOrDefault()?.Value ?? "";

            var depth = 0;
            for (var i = open; i < source.Length; i++)
            {
                if (source[i] == '{') depth++;
                else if (source[i] == '}') depth--;

                if (depth != 0) continue;

                bodies.Add("/*method:" + name + " params:" + parameters.Replace("*/", "") + "*/"
                           + source[open..(i + 1)]);
                break;
            }
        }

        return bodies;
    }

    /// <summary>
    ///     Splits the argument list starting at <paramref name="openParen" />, at top level only so
    ///     a nested call stays in one piece.
    /// </summary>
    /// <param name="body">The method body being scanned.</param>
    /// <param name="openParen">Index of the call's opening parenthesis.</param>
    /// <returns>The argument expressions, in order.</returns>
    private static List<string> Arguments(string body, int openParen)
    {
        var arguments = new List<string>();
        var depth = 0;
        var start = openParen + 1;

        for (var i = openParen; i < body.Length; i++)
        {
            var c = body[i];
            if (c is '(' or '[')
            {
                depth++;
            }
            else if (c is ')' or ']')
            {
                depth--;
                if (depth == 0)
                {
                    arguments.Add(body[start..i]);
                    return arguments;
                }
            }
            else if (c == ',' && depth == 1)
            {
                arguments.Add(body[start..i]);
                start = i + 1;
            }
        }

        arguments.Add(body[start..]);
        return arguments;
    }

    /// <summary>
    ///     The methods in a source whose every returned value is one they can vouch for.
    /// </summary>
    /// <param name="bodies">The source's method bodies.</param>
    /// <returns>The names of the resolving helpers.</returns>
    /// <remarks>
    ///     `var resolvedImagePath = ValidateParameters(p, allowlist)` is a resolved path because
    ///     the helper resolves and returns it; the caller cannot see that, but this analyser can,
    ///     and trusting the name alone was what R22-TST01 removed. A helper counts only when it
    ///     returns names and every one of them is trusted inside it; a helper that returns an
    ///     expression, or nothing, does not.
    /// </remarks>
    private static HashSet<string> ResolvingHelpers(IEnumerable<string> bodies)
    {
        var helpers = new HashSet<string>(StringComparer.Ordinal);
        foreach (var body in bodies)
        {
            var name = MethodName(body);
            if (name.Length == 0) continue;
            var returned = ReturnedName.Matches(WithoutStrings(body)).ToList();
            if (returned.Count == 0) continue;
            if (returned.All(match => TrustedOrigins(body, before: match.Index)
                    .Contains(match.Groups[1].Value)))
                helpers.Add(name);
        }

        return helpers;
    }

    /// <summary>The names a body treats as its own variables: its parameters and what it assigns.</summary>
    /// <param name="body">The method body.</param>
    /// <returns>The variable names.</returns>
    private static HashSet<string> KnownVariables(string body)
    {
        var known = new HashSet<string>(ParameterNames(body), StringComparer.Ordinal);
        foreach (var assignment in AssignmentEvents(body)) known.Add(assignment.Name);
        return known;
    }

    /// <summary>Reads declarations, all C# assignment forms, and ref/out mutations with Roslyn.</summary>
    /// <param name="body">The method body.</param>
    /// <returns>Assignment events in source order.</returns>
    private static List<AssignmentEvent> AssignmentEvents(string body)
    {
        const string prefix = "class Fixture { void Method() ";
        var root = CSharpSyntaxTree.ParseText(prefix + body + " }").GetRoot();
        var assignments = new List<AssignmentEvent>();

        foreach (var declarator in root.DescendantNodes().OfType<VariableDeclaratorSyntax>())
        {
            if (declarator.Initializer == null) continue;
            assignments.Add(new AssignmentEvent(declarator.Identifier.ValueText,
                declarator.Initializer.Value.ToString(), declarator.SpanStart - prefix.Length,
                false, ControlRegion(declarator)));
        }

        foreach (var assignment in root.DescendantNodes().OfType<AssignmentExpressionSyntax>())
        {
            var names = assignment.Left.DescendantNodesAndSelf().OfType<IdentifierNameSyntax>()
                .Select(name => name.Identifier.ValueText).Distinct(StringComparer.Ordinal);
            foreach (var name in names)
                assignments.Add(new AssignmentEvent(name, assignment.Right.ToString(),
                    assignment.SpanStart - prefix.Length,
                    !assignment.IsKind(SyntaxKind.SimpleAssignmentExpression),
                    ControlRegion(assignment)));
        }

        foreach (var argument in root.DescendantNodes().OfType<ArgumentSyntax>()
                     .Where(argument => argument.RefOrOutKeyword.IsKind(SyntaxKind.RefKeyword)
                                        || argument.RefOrOutKeyword.IsKind(SyntaxKind.OutKeyword)))
            if (argument.Expression is IdentifierNameSyntax name)
                assignments.Add(new AssignmentEvent(name.Identifier.ValueText, "",
                    argument.SpanStart - prefix.Length, false, ControlRegion(argument)));

        return assignments.OrderBy(assignment => assignment.Index).ToList();
    }

    /// <summary>Identifies the nearest branch or loop arm that contains an assignment.</summary>
    /// <param name="node">The assignment syntax.</param>
    /// <returns>A stable arm key, or null for unconditional code.</returns>
    private static string? ControlRegion(SyntaxNode node)
    {
        foreach (var ancestor in node.Ancestors())
        {
            if (ancestor is ElseClauseSyntax clause)
                return $"else:{clause.Parent?.SpanStart}";
            if (ancestor is IfStatementSyntax statement)
                return $"then:{statement.SpanStart}";
            if (ancestor is SwitchSectionSyntax section)
                return $"switch:{section.SpanStart}";
            if (ancestor is ForStatementSyntax or ForEachStatementSyntax or WhileStatementSyntax
                or DoStatementSyntax)
                return $"loop:{ancestor.SpanStart}";
            if (ancestor is ConditionalExpressionSyntax conditional)
                return $"conditional:{conditional.SpanStart}:"
                       + (conditional.WhenTrue.Span.Contains(node.Span) ? "true" : "false");
        }

        return null;
    }

    /// <summary>
    ///     The identifiers an expression names: those outside its string literals, and those in
    ///     the holes of its interpolated strings. `$"{resolvedRoot}/{userName}"` used to read as
    ///     one literal with no names in it (R23-TST01).
    /// </summary>
    /// <param name="expression">The expression.</param>
    /// <returns>The identifiers.</returns>
    private static List<string> Names(string expression)
    {
        var names = Identifier.Matches(StringLiteral.Replace(expression, " ")).Select(m => m.Value).ToList();
        foreach (Match literal in StringLiteral.Matches(expression))
        {
            var start = literal.Index;
            if (start > 0 && expression[start - 1] == '$')
                names.AddRange(InterpolationHole.Matches(literal.Value).Select(m => m.Groups[1].Value));
        }

        return names;
    }

    /// <summary>Masks string contents without changing source offsets used by point-in-time analysis.</summary>
    /// <param name="source">The source text.</param>
    /// <returns>Text of the same length with string literals replaced by spaces.</returns>
    private static string WithoutStrings(string source)
    {
        return StringLiteral.Replace(source, match => new string(' ', match.Length));
    }

    private static bool VouchedFor(IReadOnlyList<string> names, IReadOnlySet<string> known,
        IReadOnlySet<string> trusted)
    {
        var leaves = names.Where(known.Contains).Distinct().ToList();
        return leaves.Count > 0 ? leaves.All(trusted.Contains) : names.Any(trusted.Contains);
    }

    /// <summary>Evaluates which local names are trusted at one textual point in a method.</summary>
    /// <param name="body">The method body.</param>
    /// <param name="resolvingHelpers">Helpers whose every return is resolved at that return.</param>
    /// <param name="before">Exclusive source offset; later assignments cannot affect this point.</param>
    /// <returns>The names trusted immediately before the requested offset.</returns>
    private static HashSet<string> TrustedOrigins(string body, IReadOnlySet<string>? resolvingHelpers = null,
        int before = int.MaxValue)
    {
        // A parameter named for what it holds is the contract between caller and helper, and the
        // signature is where that contract is stated. Only parameters: a *local* named
        // `resolvedSuffix` is whatever its right-hand side made it, and trusting the name let
        // `var resolvedX = userPath + ".bak"` pass (R21-TST02).
        var parameters = ParameterNames(body);
        var trusted = new HashSet<string>(StringComparer.Ordinal);
        var conditionallyTainted = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var parameter in parameters)
            if (parameter.Contains("resolved", StringComparison.OrdinalIgnoreCase))
                trusted.Add(parameter);

        var assignments = AssignmentEvents(body).Where(item => item.Index < before).ToList();

        // The variables this method knows: its parameters and everything it assigns. Trust flows
        // through an expression only when *every* variable in it is trusted — the directory of a
        // resolved path is inside the allowlist, `Path.Combine(resolvedRoot, userInput)` is not
        // (R21-TST02). Names that are not variables (types, methods) do not count either way.
        var known = new HashSet<string>(parameters, StringComparer.Ordinal);
        foreach (var assignment in assignments) known.Add(assignment.Name);

        foreach (var assignment in assignments)
        {
            var (name, rhs, _, compound, region) = assignment;
            if (IsNeutralInitialiser(rhs))
            {
                trusted.Remove(name);
                if (region == null) conditionallyTainted.Remove(name);
                continue;
            }

            var vouched = !compound && (ResolverExpression.IsMatch(rhs)
                                        || (resolvingHelpers != null
                                            && resolvingHelpers.Any(h => Regex.IsMatch(rhs,
                                                @"^\s*(?:this\.)?" + Regex.Escape(h) + @"\s*\("))));

            if (!vouched)
            {
                // Literals are not variables: `Path.Combine(directory, "page.png")` mentions
                // no `page` the method knows. Types and methods are not leaves; receivers are.
                var variables = VariableOrReceiver.Matches(StringLiteral.Replace(rhs, " "))
                    .Select(m => m.Groups[1].Value)
                    .Where(known.Contains).Distinct().ToList();
                if (compound) variables.Add(name);
                vouched = variables.Count > 0
                    ? variables.All(trusted.Contains)
                    : ServerOwnedOrigins.Any(o => rhs.Contains(o, StringComparison.Ordinal));
            }

            // A raw branch poisons later branch-local resolver assignments. A neutral declaration
            // followed by one resolver assignment remains useful inside that same branch, while
            // `if raw; else resolved` cannot be laundered by the resolver appearing later in text.
            if (region == null)
            {
                conditionallyTainted.Remove(name);
            }
            else if (!vouched)
            {
                if (!conditionallyTainted.TryGetValue(name, out var regions))
                    conditionallyTainted[name] = regions = new HashSet<string>(StringComparer.Ordinal);
                regions.Add(region);
            }
            else if (conditionallyTainted.TryGetValue(name, out var regions))
            {
                // A later resolver in the same arm overwrites the earlier value on every path to
                // a following sink in that arm. Taint from a mutually exclusive arm remains.
                regions.Remove(region);
                if (regions.Count == 0) conditionallyTainted.Remove(name);
            }

            if (vouched && !conditionallyTainted.ContainsKey(name)) trusted.Add(name);
            else trusted.Remove(name);
        }

        return trusted;
    }

    /// <summary>Whether a declaration postpones choosing a path rather than assigning one.</summary>
    /// <param name="rhs">The declaration's right-hand side.</param>
    /// <returns><c>true</c> for null or default initialisers.</returns>
    private static bool IsNeutralInitialiser(string rhs)
    {
        var value = rhs.Trim();
        return value is "null" or "default" || value.StartsWith("default(", StringComparison.Ordinal);
    }

    /// <summary>The parameter names declared in the signature a body belongs to.</summary>
    /// <param name="body">The method body; its signature precedes it in the source.</param>
    /// <returns>The names, or none when the signature cannot be found.</returns>
    /// <remarks>
    ///     <see cref="MethodBodies" /> hands over the braces and everything inside; the signature
    ///     is recovered by the analyser from the source it was cut from, so a body carries its
    ///     parameters along as a leading comment line the splitter writes (see there).
    /// </remarks>
    private static IReadOnlyList<string> ParameterNames(string body)
    {
        return ParameterDeclarations(body).Select(d => d.Name).ToList();
    }

    /// <summary>The parameters declared in the signature a body belongs to, with their types.</summary>
    /// <param name="body">The method body.</param>
    /// <returns>Each parameter's declared type text and name, in order.</returns>
    private static IReadOnlyList<(string Type, string Name)> ParameterDeclarations(string body)
    {
        var marker = body.IndexOf(" params:", StringComparison.Ordinal);
        if (marker < 0) return [];

        var end = body.IndexOf("*/", marker, StringComparison.Ordinal);
        if (end < 0) return [];

        var declarations = new List<(string, string)>();
        foreach (var declaration in body[(marker + 8)..end]
                     .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var words = Identifier.Matches(declaration).Select(m => m.Value).ToList();
            if (words.Count == 0) continue;
            var name = words[^1];
            var type = words.Count >= 2 ? words[^2] : "";
            declarations.Add((type, name));
        }

        return declarations;
    }

    /// <summary>The method name a body carries in its marker.</summary>
    /// <param name="body">The method body.</param>
    /// <returns>The name, or empty when the marker is absent.</returns>
    private static string MethodName(string body)
    {
        var marker = body.IndexOf("/*method:", StringComparison.Ordinal);
        if (marker < 0) return "";
        var end = body.IndexOf(" params:", marker, StringComparison.Ordinal);
        return end < 0 ? "" : body[(marker + 9)..end].Trim();
    }

    /// <summary>
    ///     For each method in a source that has a parameter named for a resolved path, the
    ///     positions of those parameters.
    /// </summary>
    /// <param name="bodies">The source's method bodies.</param>
    /// <returns>Method name to the zero-based positions of its resolved* parameters.</returns>
    /// <remarks>
    ///     A parameter named resolved* is trusted inside its method by contract (R21-TST02); the
    ///     contract is only as good as its callers, and nothing held them to it — a helper's
    ///     `resolvedPath` fed `userPath` was trusted on the name alone (R22-TST01). The call
    ///     sites are checked instead, by <see cref="UnresolvedSinks" />.
    /// </remarks>
    private static Dictionary<string, List<int>> ResolvedParameterPositions(IEnumerable<string> bodies)
    {
        var positions = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (var body in bodies)
        {
            var name = MethodName(body);
            if (name.Length == 0) continue;
            // Strings only: `List<TabStop> resolved` is resolved tab stops, not a path.
            var parameters = ParameterDeclarations(body);
            var resolved = Enumerable.Range(0, parameters.Count)
                .Where(i => parameters[i].Type == "string"
                            && parameters[i].Name.Contains("resolved", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (resolved.Count == 0) continue;
            // Overloads with different shapes would need a signature match; the codebase's
            // helpers with resolved* parameters are not overloaded on that position, and a
            // union of positions across overloads only ever adds checks.
            if (!positions.TryGetValue(name, out var existing)) positions[name] = existing = [];
            existing.AddRange(resolved.Except(existing));
        }

        return positions;
    }

    internal static List<string> UnresolvedSinks(string source, string label = "")
    {
        var offenders = new List<string>();
        var bodies = MethodBodies(source);
        var contracts = ResolvedParameterPositions(bodies);
        var helpers = ResolvingHelpers(bodies);

        foreach (var body in bodies)
        {
            var known = KnownVariables(body);

            // The caller's side of the resolved* contract: an argument handed to a parameter
            // named for a resolved path must be one this method can vouch for (R22-TST01).
            foreach (var (callee, slots) in contracts)
            {
                var at = 0;
                var pattern = callee + "(";
                while ((at = body.IndexOf(pattern, at, StringComparison.Ordinal)) >= 0)
                {
                    var start = at;
                    at += pattern.Length;
                    if (start > 0 && (char.IsLetterOrDigit(body[start - 1]) || body[start - 1] == '_')) continue;
                    // Its own declaration marker is not a call.
                    if (start >= 9 && body[(start - 9)..start] == "/*method:") continue;

                    var arguments = Arguments(body, start + pattern.Length - 1);
                    foreach (var slot in slots)
                    {
                        if (slot >= arguments.Count) continue;
                        var argument = arguments[slot];
                        var names = Names(argument);
                        var trusted = TrustedOrigins(body, helpers, start);
                        if (names.Count == 0 || VouchedFor(names, known, trusted)) continue;

                        offenders.Add(
                            $"{label} -> {callee}({argument.Trim()}) [resolved* parameter fed an unresolved value]");
                    }
                }
            }

            var streams = new HashSet<string>(
                StreamDeclaration.Matches(body).Select(m => m.Groups[1].Value)
                    .Concat(LambdaParameter.Matches(body).Select(m => m.Groups[1].Value)),
                StringComparer.Ordinal);

            foreach (var sink in SinkPatterns.Concat(LibrarySinksIn(body)))
            {
                var position = 0;
                while ((position = body.IndexOf(sink, position, StringComparison.Ordinal)) >= 0)
                {
                    var start = position;
                    var arguments = Arguments(body, position + sink.Length - 1);
                    position += sink.Length;
                    var trusted = TrustedOrigins(body, helpers, start);

                    // ZipFile.OpenRead( ends with File.OpenRead(, so without a name boundary one
                    // call was reported under two patterns.
                    if (char.IsLetter(sink[0]) && start > 0
                                               && (char.IsLetterOrDigit(body[start - 1]) || body[start - 1] == '_'))
                        continue;

                    // Which arguments hold a path is a property of the API, not an assumption.
                    // Save has a stream overload, and only the path one is a sink; Copy, Move and
                    // Replace take a path in every position (R17-T02).
                    var relevant = PathBearing(sink, arguments);
                    if (relevant.Count == 0) continue;

                    if (sink == ".Save(")
                    {
                        // The stream overload is recognised over every argument; the path is
                        // judged on the argument that carries it, so a format or an options
                        // object beside it is not a leaf (R23-TST01). Aspose.OCR's
                        // `ImageProcessing.Save(input, folder)` carries its path second.
                        var all = relevant.SelectMany(Names).ToList();
                        if (all.Any(streams.Contains) || all.Count == 0) continue;
                        var receiverEnd = start;
                        var receiverStart = receiverEnd;
                        while (receiverStart > 0 && (char.IsLetterOrDigit(body[receiverStart - 1]) ||
                                                     body[receiverStart - 1] == '_'))
                            receiverStart--;
                        var receiver = body[receiverStart..receiverEnd];
                        var pathArgument = receiver == "ImageProcessing" && relevant.Count > 1
                            ? relevant[1]
                            : relevant[0];
                        var pathNames = Names(pathArgument);
                        if (pathNames.Count == 0 || VouchedFor(pathNames, known, trusted)) continue;

                        offenders.Add($"{label} -> {sink}{pathArgument.Trim()}");
                        continue;
                    }

                    // Each path-bearing argument on its own: a resolved source does not make an
                    // unresolved destination safe, and asking about them together is what let one
                    // vouch for the other.
                    foreach (var argument in relevant)
                    {
                        var names = Names(argument);
                        if (names.Count == 0 || VouchedFor(names, known, trusted)) continue;

                        offenders.Add($"{label} -> {sink}{argument.Trim()}");
                    }
                }
            }
        }

        return offenders;
    }

    /// <summary>Runs the analyser over one method built from its statements.</summary>
    /// <param name="statements">The method body, one statement per line.</param>
    /// <param name="signature">The method's parameters.</param>
    /// <returns>What the analyser objected to.</returns>
    /// <remarks>
    ///     Built in the shape `MethodBodies` recognises — an access modifier and braces on their
    ///     own lines — rather than as a one-liner. A fixture the analyser does not parse reports no
    ///     offenders, which reads exactly like a clean result.
    /// </remarks>
    private static List<string> Analyse(string signature, params string[] statements)
    {
        return UnresolvedSinks(
            $"public void Run({signature})\n{{\n"
            + string.Concat(statements.Select(s => $"    {s}\n"))
            + "}\n",
            "Fixture.cs");
    }

    [Theory]
    [InlineData("File.Copy(")]
    [InlineData("File.Move(")]
    [InlineData("ZipFile.CreateFromDirectory(")]
    [InlineData("ZipFile.ExtractToDirectory(")]
    public void ARawDestinationBehindAResolvedSource_ShouldBeReported(string sink)
    {
        // The shape the first-argument rule could not see: argument zero is resolved, so the check
        // was satisfied, and the argument being *written* was never looked at (R17-T02).
        var offenders = Analyse("string a, string b",
            "var resolvedSource = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            $"{sink}resolvedSource, b);");

        Assert.NotEmpty(offenders);
        Assert.Contains(offenders, offender => offender.EndsWith('b'));
    }

    [Theory]
    [InlineData("File.Copy(")]
    [InlineData("File.Move(")]
    [InlineData("ZipFile.CreateFromDirectory(")]
    [InlineData("ZipFile.ExtractToDirectory(")]
    public void BothPathsRaw_ShouldBeReportedForEachOfThem(string sink)
    {
        var offenders = Analyse("string a, string b", $"{sink}a, b);");

        Assert.Equal(2, offenders.Count);
    }

    [Theory]
    [InlineData("File.Copy(")]
    [InlineData("File.Move(")]
    [InlineData("ZipFile.CreateFromDirectory(")]
    [InlineData("ZipFile.ExtractToDirectory(")]
    public void BothPathsResolved_ShouldBeAccepted(string sink)
    {
        // The control. Reporting a correctly written call would teach a reader to ignore the gate,
        // which is worse than not having it.
        var offenders = Analyse("string a, string b",
            "var resolvedSource = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            "var resolvedTarget = SecurityHelper.ResolveAndEnsureWithinAllowlist(b, bases, \"b\");",
            $"{sink}resolvedSource, resolvedTarget);");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AllThreePathsOfFileReplace_ShouldBeJudged()
    {
        // Replace takes source, destination and backup. Only the first was ever looked at.
        var offenders = Analyse("string a, string b, string c",
            "var resolvedSource = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            "File.Replace(resolvedSource, b, c);");

        Assert.Equal(2, offenders.Count);
    }

    [Fact]
    public void AStreamSave_ShouldStillNotBeAPathSink()
    {
        // The overload distinction that already worked, kept working: `Save(stream)` writes
        // nowhere by itself. Written the way production does — a stream declared from a resolved
        // path — because that is the shape the analyser recognises as a stream.
        var offenders = Analyse("string a",
            "var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            "using var outputStream = new MemoryStream();",
            "document.Save(outputStream);");

        Assert.Empty(offenders);
    }

    /// <summary>Builds a method that resolves one path and then reaches a sink with another.</summary>
    /// <param name="sink">The sink call prefix under test.</param>
    /// <returns>Source text for the analyser.</returns>
    private static string MethodWithOneResolvedAndOneRawPath(string sink)
    {
        return "public void Run(string userPath, string other)\n"
               + "{\n"
               + "    var resolvedOther = SecurityHelper.ResolveAndEnsureWithinAllowlist(other, bases, \"o\");\n"
               + "    File.WriteAllText(resolvedOther, \"ok\");\n"
               + $"    {sink}userPath);\n"
               + "}\n";
    }

    /// <summary>Every declared sink kind, one case each.</summary>
    /// <returns>The sink patterns as theory data.</returns>
    public static TheoryData<string> SinkKinds()
    {
        var data = new TheoryData<string>();
        foreach (var sink in SinkPatterns) data.Add(sink);
        return data;
    }

    [Fact]
    public void EveryMethodThatTouchesTheFilesystem_ShouldWorkFromAResolvedPath()
    {
        var root = HandlersRoot();
        var offenders = new List<string>();
        var methodsChecked = 0;

        foreach (var file in Directory.GetFiles(root, "*.cs", SearchOption.AllDirectories).OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);

            foreach (var body in MethodBodies(source))
                if (SinkPatterns.Any(p => body.Contains(p, StringComparison.Ordinal)))
                    methodsChecked++;

            offenders.AddRange(UnresolvedSinks(source, Path.GetRelativePath(root, file)));
        }

        Assert.True(methodsChecked > 0,
            "No filesystem-touching method was found, so this guard checked nothing. "
            + "The method splitter or the sink list is probably broken.");

        Assert.True(offenders.Count == 0,
            $"{offenders.Count} sink(s) receive a path that was never resolved:\n  "
            + string.Join("\n  ", offenders));
    }

    [Theory]
    [MemberData(nameof(SinkKinds))]
    public void EverySinkKind_ShouldBeFlaggedWhenItsPathWasNeverResolved(string sink)
    {
        var offenders = UnresolvedSinks(MethodWithOneResolvedAndOneRawPath(sink), "synthetic");

        Assert.True(offenders.Count == 1,
            $"{sink} with a raw caller path should be reported exactly once, got "
            + $"{offenders.Count}: {string.Join(" | ", offenders)}");
        Assert.Contains("userPath", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void APathDerivedFromAResolvedOne_ShouldBeAccepted()
    {
        const string source = "public void Run(string userPath)\n"
                              + "{\n"
                              + "    var resolved = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");\n"
                              + "    var directory = Path.GetDirectoryName(resolved);\n"
                              + "    Directory.CreateDirectory(directory);\n"
                              + "    var page = Path.Combine(directory, \"page.png\");\n"
                              + "    File.WriteAllBytes(page, bytes);\n"
                              + "}\n";

        Assert.Empty(UnresolvedSinks(source, "synthetic"));
    }

    [Fact]
    public void APathTheServerMintedForItself_ShouldBeAccepted()
    {
        const string source = "public void Run()\n"
                              + "{\n"
                              + "    var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());\n"
                              + "    Directory.CreateDirectory(temp);\n"
                              + "    File.Delete(temp);\n"
                              + "}\n";

        Assert.Empty(UnresolvedSinks(source, "synthetic"));
    }

    [Fact]
    public void SavingThroughAStream_ShouldNotCountAsAPathSink()
    {
        const string source = "public void Run(Workbook workbook)\n"
                              + "{\n"
                              + "    using var bounded = new BoundedWriteStream(buffer, cap, \"CSV output\");\n"
                              + "    workbook.Save(bounded, saveOptions);\n"
                              + "}\n";

        Assert.Empty(UnresolvedSinks(source, "synthetic"));
    }

    [Fact]
    public void ADirectoryCreatedBeforeTheAllowlistCheck_ShouldBeFlagged()
    {
        // The shape three Word render handlers actually had: the caller's string reaches
        // Directory.CreateDirectory, and only the later write is resolved. A guard that looked for
        // the word "resolved" anywhere in the method called this compliant (R7-T01).
        const string source = "public void Run(string userPath)\n"
                              + "{\n"
                              + "    var directory = Path.GetDirectoryName(userPath);\n"
                              + "    Directory.CreateDirectory(directory);\n"
                              + "    var resolved = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");\n"
                              + "    doc.Save(resolved, options);\n"
                              + "}\n";

        var offenders = UnresolvedSinks(source, "synthetic");

        Assert.Single(offenders);
        Assert.Contains("Directory.CreateDirectory(", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ALibraryPathSink_ShouldBeFlaggedWhenItsPathWasNeverResolved()
    {
        // R19-OCR01. `OcrInput.Add` opens the path it is given, exactly as `File.OpenRead` does,
        // and the inventory had no entry for it — so the OCR handlers handed it the caller's own
        // spelling and this test reported a clean file.
        const string body = """
                            public void Run(string userPath)
                            {
                                using var input = new OcrInput(InputType.SingleImage, filters);
                                input.Add(userPath);
                            }
                            """;

        var offenders = UnresolvedSinks(body, "synthetic");

        Assert.True(offenders.Count == 1,
            $"a library sink given a raw caller path should be reported exactly once, got "
            + $"{offenders.Count}: {string.Join(" | ", offenders)}");
        Assert.Contains("userPath", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ALibraryPathSink_ShouldBeAcceptedWhenItsPathWasResolved()
    {
        const string body = """
                            public void Run(string userPath)
                            {
                                var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, roots, "path");
                                using var input = new OcrInput(InputType.SingleImage, filters);
                                input.Add(resolvedPath);
                            }
                            """;

        Assert.Empty(UnresolvedSinks(body, "synthetic"));
    }

    [Fact]
    public void EveryLibraryPathSinkType_ShouldStillBeConstructedSomewhereInHandlers()
    {
        // The inventory is the one place a whole class of sink can go unmodelled, so an entry that
        // stops describing this codebase has to fail rather than sit there looking like coverage.
        var sources = Directory
            .EnumerateFiles(HandlersRoot(), "*.cs", SearchOption.AllDirectories)
            .Select(File.ReadAllText)
            .ToList();

        Assert.NotEmpty(LibraryPathSinks);

        foreach (var (type, methods) in LibraryPathSinks)
        {
            Assert.True(sources.Any(s => s.Contains($"new {type}(", StringComparison.Ordinal)),
                $"no handler constructs {type} any more, so its entry describes nothing");
            Assert.NotEmpty(methods);
        }
    }

    [Fact]
    public void ALocalNamedResolved_ShouldNotBeTrustedForItsName()
    {
        // R21-TST02. `resolvedSuffix` is whatever its right-hand side made it.
        var offenders = Analyse("string userPath",
            "var resolvedSuffix = userPath + \".bak\";",
            "File.WriteAllText(resolvedSuffix, \"x\");");

        Assert.Single(offenders);
        Assert.Contains("resolvedSuffix", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AMixOfTrustedAndUntrusted_ShouldLoseTrust()
    {
        var offenders = Analyse("string resolvedRoot, string userInput",
            "var target = Path.Combine(resolvedRoot, userInput);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
        Assert.Contains("target", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AResolvedLocalLaterOverwrittenByARawPath_ShouldLoseTrust()
    {
        // R29-TST01. Trust belongs to the value reaching the sink, not permanently to the local
        // that once held a resolved value.
        var offenders = Analyse("string userPath",
            "var target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");",
            "target = userPath;",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
        Assert.Contains("target", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AResolvedLocalOverwrittenInsideAnInlineBranch_ShouldLoseTrust()
    {
        var offenders = Analyse("string userPath, bool replace",
            "var target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");",
            "if (replace) target = userPath;",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AResolvedLocalMutatedWithAnUntrustedSuffix_ShouldLoseTrust()
    {
        var offenders = Analyse("string userPath, string userSuffix",
            "var target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");",
            "target += userSuffix;",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AResolvedAssignmentOnTheLaterElseBranch_ShouldNotHideTheRawIfBranch()
    {
        var offenders = Analyse("string userPath, bool chooseRaw",
            "string target;",
            "if (chooseRaw) target = userPath;",
            "else target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void ADeconstructionAssignmentContainingARawPath_ShouldLoseTrust()
    {
        var offenders = Analyse("string userPath, string resolvedPath",
            "var target = resolvedPath;",
            "var other = resolvedPath;",
            "(target, other) = (userPath, resolvedPath);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AResolvedPathPassedByReference_ShouldLoseTrust()
    {
        var offenders = Analyse("string userPath",
            "var target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");",
            "Mutate(ref target);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AHelperFedAMixOfTrustedAndUntrusted_ShouldBeReportedAtTheCaller()
    {
        // R23-TST01. One trusted name in the argument vouched for the untrusted one beside it.
        var offenders = UnresolvedSinks(
            "private static void Helper(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string resolvedRoot, Request request)\n{\n"
            + "    Helper(Path.Combine(resolvedRoot, request.UserInput));\n}\n", "Fixture.cs");

        Assert.Single(offenders);
        Assert.Contains("resolved* parameter fed an unresolved value", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ASaveFedAMixOfTrustedAndUntrusted_ShouldBeReported()
    {
        var offenders = Analyse("string resolvedRoot, string userName, Document document",
            "document.Save(Path.Combine(resolvedRoot, userName));");

        Assert.Contains(offenders, o => o.Contains(".Save(", StringComparison.Ordinal));
    }

    [Fact]
    public void ASaveWithAnOptionsObjectBesideATrustedPath_ShouldPass()
    {
        // The options beside the path are not a leaf of the path.
        var offenders = Analyse("string a, Document document",
            "var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            "var options = new SaveOptions();",
            "document.Save(resolvedPath, options);");

        Assert.Empty(offenders);
    }

    [Fact]
    public void ADirectSinkFedAConditionalMix_ShouldBeReported()
    {
        var offenders = Analyse("string resolvedRoot, string userName, bool flag",
            "File.WriteAllText(flag ? resolvedRoot : userName, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void ADirectSinkFedAnInterpolatedMix_ShouldBeReported()
    {
        var offenders = Analyse("string resolvedRoot, string userName",
            "File.WriteAllText($\"{resolvedRoot}/{userName}.txt\", \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void ADirectSinkFedAnIndexedMix_ShouldBeReported()
    {
        var offenders = Analyse("string resolvedRoot, string[] inputs",
            "File.WriteAllText(Path.Combine(resolvedRoot, inputs[0]), \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void ADirectSinkFedAHelperReturnOverAMix_ShouldBeReported()
    {
        var offenders = Analyse("string resolvedRoot, Request request",
            "File.WriteAllText(Path.Combine(resolvedRoot, request.Name()), \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void ADirectSinkFedOnlyTrustedLeaves_ShouldPass()
    {
        // The control: every leaf trusted, whatever surrounds them. A condition variable is a
        // leaf too — the analyser has no types — so the control keeps its leaves to paths.
        var offenders = Analyse("string resolvedRoot, string resolvedName",
            "File.WriteAllText(Path.Combine(resolvedRoot, Path.GetFileName(resolvedName)), \"x\");");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AMemberOfAnUntrustedParameter_ShouldPoisonTheDerivation()
    {
        // R22-TST01. `request` was "only a receiver" and did not count; the one variable that
        // did count was trusted, and so was the whole expression.
        var offenders = Analyse("string resolvedRoot, Request request",
            "var target = Path.Combine(resolvedRoot, request.UserInput);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
        Assert.Contains("target", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ANestedMemberOfAnUntrustedParameter_ShouldPoisonTheDerivation()
    {
        var offenders = Analyse("string resolvedRoot, Request request",
            "var target = Path.Combine(resolvedRoot, request.Body.RelativePath);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AMethodCalledOnAnUntrustedParameter_ShouldPoisonTheDerivation()
    {
        var offenders = Analyse("string resolvedRoot, Request request",
            "var target = Path.Combine(resolvedRoot, request.ResolveName());",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AnIndexIntoAnUntrustedParameter_ShouldPoisonTheDerivation()
    {
        var offenders = Analyse("string resolvedRoot, string[] inputs",
            "var target = Path.Combine(resolvedRoot, inputs[0]);",
            "File.WriteAllText(target, \"x\");");

        Assert.Single(offenders);
    }

    [Fact]
    public void AResolvedParameterFedAnUnresolvedValue_ShouldBeReportedAtTheCaller()
    {
        // R22-TST01, the contract's other half. `Helper` trusts its parameter by name; that is
        // only true if every caller keeps the promise, and this caller does not.
        var offenders = UnresolvedSinks(
            "private static void Helper(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string userPath)\n{\n"
            + "    Helper(userPath);\n}\n", "Fixture.cs");

        Assert.Single(offenders);
        Assert.Contains("Helper(userPath)", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AResolvedParameterFedAResolvedValue_ShouldPass()
    {
        // The control: a caller that resolves first keeps the promise.
        var offenders = UnresolvedSinks(
            "private static void Helper(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string userPath)\n{\n"
            + "    var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");\n"
            + "    Helper(resolvedPath);\n}\n", "Fixture.cs");

        Assert.Empty(offenders);
    }

    [Fact]
    public void ANonStringResolvedParameter_IsNotAPathContract()
    {
        // `List<TabStop> resolved` is resolved tab stops; the contract is about string paths.
        var offenders = UnresolvedSinks(
            "private static void Apply(Document doc, List<TabStop> resolved)\n{\n"
            + "    doc.Use(resolved);\n}\n"
            + "public void Run(string userPath, List<TabStop> stops)\n{\n"
            + "    Apply(doc, stops);\n}\n", "Fixture.cs");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AValueReturnedByAResolvingHelper_ShouldBeTrustedAtTheCaller()
    {
        // The helper resolves and returns; the caller's local is what it returned.
        var offenders = UnresolvedSinks(
            "private static string Validate(string userPath)\n{\n"
            + "    var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");\n"
            + "    return resolvedPath;\n}\n"
            + "private static void Write(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string userPath)\n{\n"
            + "    var target = Validate(userPath);\n"
            + "    Write(target);\n    File.WriteAllText(target, \"y\");\n}\n", "Fixture.cs");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AHelperThatReturnsAnUnresolvedValue_IsNotAResolvingHelper()
    {
        var offenders = UnresolvedSinks(
            "private static string Validate(string userPath)\n{\n"
            + "    return userPath;\n}\n"
            + "private static void Write(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string userPath)\n{\n"
            + "    var target = Validate(userPath);\n"
            + "    Write(target);\n}\n", "Fixture.cs");

        Assert.Single(offenders);
        Assert.Contains("Write(target)", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AHelperWithAnUntrustedEarlyReturn_IsNotAResolvingHelper()
    {
        var offenders = UnresolvedSinks(
            "private static string Validate(string userPath, bool early)\n{\n"
            + "    var target = userPath;\n"
            + "    if (early) return target;\n"
            + "    target = SecurityHelper.ResolveAndEnsureWithinAllowlist(userPath, bases, \"p\");\n"
            + "    return target;\n}\n"
            + "private static void Write(string resolvedTarget)\n{\n"
            + "    File.WriteAllText(resolvedTarget, \"x\");\n}\n"
            + "public void Run(string userPath)\n{\n"
            + "    var target = Validate(userPath, true);\n"
            + "    Write(target);\n}\n", "Fixture.cs");

        Assert.Single(offenders);
        Assert.Contains("Write(target)", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AMemberOfATrustedVariable_ShouldStayTrusted()
    {
        // The control: a receiver that is trusted carries trust to its members too.
        var offenders = Analyse("string a",
            "var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, bases, \"a\");",
            "var info = new FileInfo(resolvedPath);",
            "var target = Path.Combine(info.DirectoryName, \"out.txt\");",
            "File.WriteAllText(target, \"x\");");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AValueDerivedOnlyFromTrustedVariables_ShouldStayTrusted()
    {
        var offenders = Analyse("string resolvedPath",
            "var directory = Path.GetDirectoryName(resolvedPath);",
            "Directory.CreateDirectory(directory);");

        Assert.Empty(offenders);
    }

    [Fact]
    public void AParameterNamedResolved_IsTheContractAndIsTrusted()
    {
        Assert.Empty(Analyse("string resolvedOutputPath",
            "File.WriteAllText(resolvedOutputPath, \"x\");"));
    }

    /// <summary>One local assignment, with the control-flow shape that decides safe re-trust.</summary>
    /// <param name="Name">The local being assigned or passed by reference.</param>
    /// <param name="Rhs">The assigned expression, or an empty value for an opaque ref/out write.</param>
    /// <param name="Index">Its source offset inside the method body.</param>
    /// <param name="Compound">Whether the result also depends on the old local value.</param>
    /// <param name="Region">The nearest branch/loop arm, or null for unconditional code.</param>
    private sealed record AssignmentEvent(
        string Name,
        string Rhs,
        int Index,
        bool Compound,
        string? Region);
}
