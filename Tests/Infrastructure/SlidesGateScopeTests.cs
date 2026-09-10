using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Every place that constructs an <c>Aspose.Slides.Presentation</c> must do so inside a hold on
///     <c>SlidesGate</c>, and the inventory recording how each one is protected must be true.
///     <para>
///         This used to be brace matching over text with comments and string literals blanked out.
///         It got the common cases right and could not decide the rest: whether a gate and an
///         acquisition are in the same method, whether an early return reaches the acquisition
///         before the gate, whether a <c>GatedByCallee</c> claim still corresponds to a call that
///         exists. Those were listed as needing Roslyn (§23.8, §23.27.1), and Roslyn is now here.
///     </para>
///     <para>
///         What the syntax tree settles that text could not: a <c>using</c> declaration's scope is
///         the remainder of its block and a <c>using</c> statement's is the statement it governs,
///         both of which the parser knows exactly. An acquisition inside that scope is reached
///         only after the gate is taken — there is no path around it — and C# releases the hold on
///         every exit from the scope including exceptions. So containment in the syntactic scope
///         <em>is</em> the path-coverage proof; it does not need a separate one.
///     </para>
/// </summary>
public class SlidesGateScopeTests : TestBase
{
    /// <summary>
    ///     Every method that comes to hold a presentation, and what protects it. Written out per
    ///     method rather than per file: "this file has two gates" says nothing about which
    ///     acquisition either of them covers (§23.8).
    /// </summary>
    private static readonly (string File, string Method, Protection How, string Why)[] Inventory =
    [
        ("Core/Conversion/DocumentConverter.cs", "ConvertPowerPointDocument", Protection.SelfGated,
            "the file-conversion boundary"),
        ("Core/Conversion/DocumentConverter.cs", "ConvertPowerPointToStream", Protection.SelfGated,
            "the in-memory conversion boundary"),
        ("Core/Conversion/DocumentConverter.cs", "ConvertToStream", Protection.GatedByCallee,
            "casts the document and hands it to ConvertPowerPointToStream, which gates"),
        ("Core/Conversion/DocumentConverter.cs", "SlideCountOf", Protection.SelfGated,
            "counting slides enters the library, and the conversion has not taken its hold yet"),
        ("Core/Session/DocumentContext.cs", "Create", Protection.HeldByInstance,
            "takes the gate before loading and stores it on the context, which releases it last"),
        ("Core/Session/DocumentContext.cs", "CreateCore", Protection.GatedByCaller,
            "called only by Create, which already holds it"),
        ("Core/Session/DocumentContext.cs", "Dispose", Protection.GatedByCaller,
            "releases the hold Create took, last"),
        ("Core/Session/DocumentSession.cs", "Dispose", Protection.SelfGated,
            "releasing the presentation re-enters the library"),
        ("Core/Session/DocumentSessionManager.cs", "LoadPresentation", Protection.SelfGated,
            "opens a presentation with no session around it"),
        ("Core/Session/DocumentSessionManager.cs", "SaveDocumentToFile", Protection.SelfGated,
            "every save funnels through here"),
        ("Tools/Conversion/ConvertDocumentTool.cs", "ConvertFromFile", Protection.SelfGated,
            "constructs its own presentation"),
        ("Tools/Conversion/ConvertDocumentTool.cs", "ConvertFromSession", Protection.GatedByCallee,
            "hands the session document to ConvertPowerPointDocument, which gates"),
        ("Handlers/PowerPoint/FileOperations/ConvertPresentationHandler.cs", "Execute",
            Protection.SelfGated, "opens its own"),
        ("Handlers/PowerPoint/FileOperations/CreatePresentationHandler.cs", "Execute",
            Protection.SelfGated, "builds its own"),
        ("Handlers/PowerPoint/FileOperations/MergePresentationsHandler.cs", "Execute",
            Protection.SelfGated, "opens the master and each source"),
        ("Handlers/PowerPoint/FileOperations/SplitPresentationHandler.cs", "Execute",
            Protection.SelfGated, "opens the source and builds each output"),
        ("Handlers/PowerPoint/Layout/ApplyThemeHandler.cs", "Execute", Protection.SelfGated,
            "opens the theme presentation")
    ];

    /// <summary>
    ///     Walks up from the test binaries to the repository root so the scan works from any
    ///     working directory the runner chooses.
    /// </summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    /// <summary>Parses one production file into a syntax tree.</summary>
    /// <param name="relativePath">Repository-relative path.</param>
    /// <returns>The file's root syntax node.</returns>
    private static CompilationUnitSyntax Parse(string relativePath)
    {
        var full = Path.Combine(RepositoryRoot().FullName, relativePath);
        Assert.True(File.Exists(full), $"{relativePath} does not exist");

        return CSharpSyntaxTree.ParseText(File.ReadAllText(full)).GetCompilationUnitRoot();
    }

    /// <summary>Finds one method (or local function) declaration by name.</summary>
    /// <param name="root">The parsed file.</param>
    /// <param name="name">The method's name.</param>
    /// <returns>The declarations found; more than one when the method is overloaded.</returns>
    private static List<SyntaxNode> MethodsNamed(CompilationUnitSyntax root, string name)
    {
        // The file's own type, not the helpers nested inside it. `DocumentSession.cs` declares
        // `Dispose()` on the session and again on a nested scope object; matching by name alone
        // judged the nested one too and reported a missing gate it does not need.
        var outerTypes = root.DescendantNodes()
            .OfType<TypeDeclarationSyntax>()
            .Where(type => !type.Ancestors().OfType<TypeDeclarationSyntax>().Any())
            .ToList();

        var methods = outerTypes
            .SelectMany(type => type.Members.OfType<MethodDeclarationSyntax>())
            .Where(m => m.Identifier.ValueText == name)
            .Cast<SyntaxNode>();

        var locals = outerTypes
            .SelectMany(type => type.Members.OfType<MethodDeclarationSyntax>())
            .SelectMany(m => m.DescendantNodes().OfType<LocalFunctionStatementSyntax>())
            .Where(f => f.Identifier.ValueText == name)
            .Cast<SyntaxNode>();

        return methods.Concat(locals).ToList();
    }

    /// <summary>Every <c>new Presentation(...)</c> inside a node.</summary>
    /// <param name="node">The node to search.</param>
    /// <returns>The object-creation expressions.</returns>
    private static List<BaseObjectCreationExpressionSyntax> AcquisitionsIn(SyntaxNode node)
    {
        return node.DescendantNodes().OfType<ObjectCreationExpressionSyntax>()
            .Where(creation => creation.Type is IdentifierNameSyntax { Identifier.ValueText: "Presentation" }
                or QualifiedNameSyntax { Right.Identifier.ValueText: "Presentation" })
            .Cast<BaseObjectCreationExpressionSyntax>()
            .ToList();
    }

    /// <summary>Whether an expression is a call to <c>SlidesGate.Enter()</c>.</summary>
    /// <param name="expression">The expression to test.</param>
    /// <returns><c>true</c> when it enters the gate.</returns>
    private static bool IsGateEntry(ExpressionSyntax? expression)
    {
        if (expression == null) return false;

        // Anywhere in the expression, not only at its root: the session paths write
        // `Document is Presentation ? SlidesGate.Enter() : null`, and a root-only test read that
        // as taking no gate at all.
        return expression.DescendantNodesAndSelf()
            .OfType<InvocationExpressionSyntax>()
            .Any(invocation => invocation.Expression is MemberAccessExpressionSyntax
            {
                Name.Identifier.ValueText: "Enter",
                Expression: IdentifierNameSyntax { Identifier.ValueText: "SlidesGate" }
            });
    }

    /// <summary>
    ///     The syntactic scopes in which a gate is held inside this method.
    ///     <para>
    ///         A <c>using</c> declaration holds from its own statement to the end of the block that
    ///         contains it; a <c>using</c> statement holds for the statement it governs. Both are
    ///         exact, and both release on every exit including an exception — which is why a node
    ///         inside one of these spans is covered on every path that can reach it.
    ///     </para>
    /// </summary>
    /// <param name="method">The method to inspect.</param>
    /// <returns>The text spans over which the gate is held.</returns>
    private static List<TextSpan> GateHoldsIn(SyntaxNode method)
    {
        var holds = new List<TextSpan>();

        foreach (var statement in method.DescendantNodes().OfType<UsingStatementSyntax>())
        {
            var entersHere =
                IsGateEntry(statement.Expression)
                || (statement.Declaration?.Variables
                    .Any(v => IsGateEntry(v.Initializer?.Value)) ?? false);

            if (entersHere) holds.Add(statement.Statement.Span);
        }

        foreach (var declaration in method.DescendantNodes().OfType<LocalDeclarationStatementSyntax>())
        {
            if (declaration.UsingKeyword == default) continue;
            if (!declaration.Declaration.Variables.Any(v => IsGateEntry(v.Initializer?.Value)))
                continue;

            // A using declaration holds until the end of its enclosing block, so the hold runs
            // from just after this statement to that block's closing brace.
            var block = declaration.Ancestors().OfType<BlockSyntax>().FirstOrDefault();
            if (block == null) continue;

            holds.Add(TextSpan.FromBounds(declaration.Span.End, block.Span.End));
        }

        return holds;
    }

    /// <summary>Every method this file calls from within the named method.</summary>
    /// <param name="method">The calling method.</param>
    /// <returns>The names of the methods it invokes.</returns>
    private static HashSet<string> CallsMadeBy(SyntaxNode method)
    {
        var called = new HashSet<string>(StringComparer.Ordinal);

        foreach (var invocation in method.DescendantNodes().OfType<InvocationExpressionSyntax>())
            switch (invocation.Expression)
            {
                case MemberAccessExpressionSyntax member:
                    called.Add(member.Name.Identifier.ValueText);
                    break;
                case IdentifierNameSyntax identifier:
                    called.Add(identifier.Identifier.ValueText);
                    break;
            }

        return called;
    }

    [Fact]
    public void EverySelfGatedAcquisition_ShouldSitInsideItsOwnMethodsHold()
    {
        var wrong = new List<string>();

        foreach (var (file, method, how, why) in Inventory)
        {
            if (how != Protection.SelfGated) continue;

            var root = Parse(file);
            var declarations = MethodsNamed(root, method);

            if (declarations.Count == 0)
            {
                wrong.Add($"{file}: no method named {method} — the inventory is out of date");
                continue;
            }

            foreach (var declaration in declarations)
            {
                // The hold is checked first, and unconditionally. Checking it only when the method
                // *constructs* a presentation let a self-gated method that merely uses one —
                // SlideCountOf reads Slides.Count — lose its gate with this test still green. A
                // guard that skips the case it was written for is the false green this whole
                // exercise keeps producing.
                var holds = GateHoldsIn(declaration);
                if (holds.Count == 0)
                {
                    wrong.Add($"{file}::{method} is recorded as self-gated but takes no gate — "
                              + $"{why}");
                    continue;
                }

                var acquisitions = AcquisitionsIn(declaration);

                foreach (var acquisition in acquisitions)
                    if (!holds.Any(hold => hold.Contains(acquisition.Span)))
                        wrong.Add(
                            $"{file}::{method} constructs a presentation outside every hold it "
                            + "takes, so a path reaches the library without the gate — "
                            + $"{why}");
            }
        }

        Assert.True(wrong.Count == 0,
            "these acquisitions are not covered by a gate on their own path:\n  "
            + string.Join("\n  ", wrong));
    }

    [Fact]
    public void EveryHeldByInstanceClaim_ShouldTakeAGateAndHandItToTheInstance()
    {
        // The hold has no syntactic scope to be contained in, so what is checked instead is that
        // the method takes a gate at all and that the value leaves the method — assigned to a
        // field or a member of the object it returns. A hold taken and dropped on the floor would
        // release at once and protect nothing.
        var wrong = new List<string>();

        foreach (var (file, method, how, why) in Inventory)
        {
            if (how != Protection.HeldByInstance) continue;

            foreach (var declaration in MethodsNamed(Parse(file), method))
            {
                var entries = declaration.DescendantNodes()
                    .OfType<ExpressionSyntax>()
                    .Where(IsGateEntry)
                    .ToList();

                if (entries.Count == 0)
                {
                    wrong.Add($"{file}::{method} is recorded as holding the gate on the instance "
                              + $"but never takes one — {why}");
                    continue;
                }

                // The hold must be stored, not discarded: either into a local that is later
                // assigned onward, or straight into a member.
                var storesIt = declaration.DescendantNodes()
                    .OfType<AssignmentExpressionSyntax>()
                    .Any(assignment => assignment.Left is MemberAccessExpressionSyntax
                                       || assignment.Left is IdentifierNameSyntax);

                var declaresIt = declaration.DescendantNodes()
                    .OfType<VariableDeclaratorSyntax>()
                    .Any(v => IsGateEntry(v.Initializer?.Value));

                if (!storesIt && !declaresIt)
                    wrong.Add($"{file}::{method} takes a gate and does not keep it — {why}");
            }
        }

        Assert.True(wrong.Count == 0,
            "these HeldByInstance claims no longer hold:\n  " + string.Join("\n  ", wrong));
    }

    [Fact]
    public void EveryGatedByCalleeClaim_ShouldStillCorrespondToACallThatGates()
    {
        // The claim is that this method hands the presentation to something that gates. If that
        // call is refactored away the claim becomes a comment that happens to be false, which is
        // exactly what a hand-written inventory cannot notice on its own.
        var selfGated = Inventory
            .Where(entry => entry.How == Protection.SelfGated)
            .Select(entry => entry.Method)
            .ToHashSet(StringComparer.Ordinal);

        var wrong = new List<string>();

        foreach (var (file, method, how, why) in Inventory)
        {
            if (how != Protection.GatedByCallee) continue;

            var declarations = MethodsNamed(Parse(file), method);
            if (declarations.Count == 0)
            {
                wrong.Add($"{file}: no method named {method}");
                continue;
            }

            var reaches = declarations
                .SelectMany(CallsMadeBy)
                .Any(called => selfGated.Contains(called));

            if (!reaches)
                wrong.Add($"{file}::{method} is recorded as gated by a callee, but calls no "
                          + $"self-gated method — {why}");
        }

        Assert.True(wrong.Count == 0,
            "these GatedByCallee claims no longer hold:\n  " + string.Join("\n  ", wrong));
    }

    [Fact]
    public void EveryGatedByCallerClaim_ShouldOnlyBeReachedFromAGatedMethod()
    {
        // The mirror of the claim above: this method does not gate, so every route into it must
        // already hold. Checked against the files the inventory names, which is where such a call
        // would be.
        var gated = Inventory
            .Where(entry => entry.How is Protection.SelfGated or Protection.GatedByCaller
                or Protection.HeldByInstance)
            .Select(entry => entry.Method)
            .ToHashSet(StringComparer.Ordinal);

        var wrong = new List<string>();

        foreach (var (file, method, how, why) in Inventory)
        {
            if (how != Protection.GatedByCaller) continue;

            var root = Parse(file);
            var callers = Inventory
                .Where(entry => entry.File == file)
                .SelectMany(entry => MethodsNamed(root, entry.Method)
                    .Where(declaration => CallsMadeBy(declaration).Contains(method))
                    .Select(_ => entry.Method))
                .ToList();

            if (callers.Count == 0)
            {
                wrong.Add($"{file}::{method} is recorded as gated by its caller, but no method in "
                          + $"the inventory for that file calls it — {why}");
                continue;
            }

            var ungated = callers.Where(caller => !gated.Contains(caller)).ToList();
            if (ungated.Count > 0)
                wrong.Add($"{file}::{method} is reached from {string.Join(", ", ungated)}, which "
                          + $"do not hold the gate — {why}");
        }

        Assert.True(wrong.Count == 0,
            "these GatedByCaller claims no longer hold:\n  " + string.Join("\n  ", wrong));
    }

    [Fact]
    public void TheAnalyser_ShouldTellAGateThatCoversTheAcquisitionFromOneThatDoesNot()
    {
        // The two shapes brace matching could not separate, and the reason this analyser exists.
        // Both take the gate before constructing; only the first still holds it at that point.
        const string covered = """
                               class C
                               {
                                   void M()
                                   {
                                       using var gate = SlidesGate.Enter();
                                       var p = new Presentation();
                                   }
                               }
                               """;

        const string released = """
                                class C
                                {
                                    void M()
                                    {
                                        using (SlidesGate.Enter())
                                        {
                                        }

                                        var p = new Presentation();
                                    }
                                }
                                """;

        static bool IsCovered(string source)
        {
            var root = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
            var method = MethodsNamed(root, "M").Single();
            var holds = GateHoldsIn(method);

            return AcquisitionsIn(method).All(a => holds.Any(hold => hold.Contains(a.Span)));
        }

        Assert.True(IsCovered(covered), "a using declaration must cover what follows it");
        Assert.False(IsCovered(released),
            "a using statement that closes before the acquisition must not count as covering it");
    }

    [Fact]
    public void TheAnalyser_ShouldNotBeFooledByCommentsOrStringLiterals()
    {
        // The text scanner needed comments and literals blanked out by hand. The parser never
        // sees them as code in the first place.
        const string source = """
                              class C
                              {
                                  void M()
                                  {
                                      // using var gate = SlidesGate.Enter();
                                      var note = "using var gate = SlidesGate.Enter();";
                                      var p = new Presentation();
                                  }
                              }
                              """;

        var root = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        var method = MethodsNamed(root, "M").Single();

        Assert.Empty(GateHoldsIn(method));
        Assert.Single(AcquisitionsIn(method));
    }

    [Fact]
    public void TheAnalyser_ShouldSeeAcquisitionsInsideNestedScopes()
    {
        // A hold taken at the top of a method covers a construction several blocks deeper; a
        // check that only looked at the method's immediate statements would miss it.
        const string source = """
                              class C
                              {
                                  void M(bool flag)
                                  {
                                      using var gate = SlidesGate.Enter();

                                      if (flag)
                                      {
                                          foreach (var x in new int[0])
                                          {
                                              var p = new Presentation();
                                          }
                                      }
                                  }
                              }
                              """;

        var root = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        var method = MethodsNamed(root, "M").Single();
        var holds = GateHoldsIn(method);

        Assert.All(AcquisitionsIn(method),
            acquisition => Assert.Contains(holds, hold => hold.Contains(acquisition.Span)));
    }

    [Fact]
    public void TheInventory_ShouldNameEveryFileThatConstructsAPresentation()
    {
        // The set has to stay complete: a new file that constructs one needs its own line here,
        // with its own reason.
        var root = RepositoryRoot();
        var listed = Inventory.Select(entry => entry.File).ToHashSet(StringComparer.Ordinal);

        var unlisted = new[] { "Core", "Handlers", "Helpers", "Tools" }
            .Select(name => Path.Combine(root.FullName, name))
            .Where(Directory.Exists)
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Select(path => Path.GetRelativePath(root.FullName, path).Replace('\\', '/'))
            .Where(relative => !listed.Contains(relative))
            .Where(relative => AcquisitionsIn(Parse(relative)).Count > 0)
            .ToList();

        Assert.True(unlisted.Count == 0,
            "these files construct a presentation but are not in the gate inventory: "
            + string.Join(", ", unlisted));
    }

    /// <summary>How a method that acquires a presentation is protected.</summary>
    private enum Protection
    {
        /// <summary>The method takes the gate itself, and the acquisition is inside that hold.</summary>
        SelfGated,

        /// <summary>The method it hands the presentation to takes the gate.</summary>
        GatedByCallee,

        /// <summary>Whoever created the object holds the gate for this method's whole life.</summary>
        GatedByCaller,

        /// <summary>
        ///     The method stores the hold on the instance and releases it when the instance is
        ///     disposed, so the hold spans the object's life rather than a syntactic scope.
        ///     <para>
        ///         Deliberate: <see cref="Protection.SelfGated" />'s proof is containment in a
        ///         `using` scope, and this pattern has no such scope to be contained in. Judging
        ///         it by that rule reported a defect that was not there, so it is classified
        ///         instead and checked by the rule that does apply (§23.30).
        ///     </para>
        /// </summary>
        HeldByInstance
    }
}
