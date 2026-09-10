using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     A handler that writes many files into a directory the caller named must consult a budget in
///     the same method that does the writing.
///     <para>
///         <see cref="FanOutBudgetInventoryTests" /> asks whether a budget name appears anywhere in
///         the file, with comments stripped. That catches a budget deleted outright and cannot see
///         a budget that has drifted away from the loop it is supposed to bound — named in one
///         method while another writes without it. The limitation was recorded honestly (§22.13.1,
///         §23.8) and left in place because proving otherwise needed a parser.
///     </para>
///     <para>
///         Now it has one. What is checked here is per method: a method that writes to disk inside
///         a loop must itself use a budget. That is weaker than proving the budget dominates the
///         sink — an `if (false)` guard would still satisfy it — and much stronger than "the file
///         contains the word somewhere", which was satisfied by a budget in an unrelated method
///         twenty lines away.
///     </para>
///     <para>
///         Three things it did less well than it read (R13-T02). It scanned <c>Handlers/</c> only,
///         so a loop writing files from anywhere else was never asked the question; it now scans
///         every production source. The budget was looked for with a substring search over the
///         method's text, comments included — a comment saying "no RenderBudget needed here" was
///         indistinguishable from using one; it is now an identifier in the syntax tree. And it
///         asked only whether the budget appeared, not where: a budget consulted after the write,
///         or inside a branch the write is not in — <c>if (false)</c> included — satisfied it.
///     </para>
///     <para>
///         Position is decidable from the syntax tree alone. Statements run in order, and a block
///         introduces no path around what precedes it, so a budget that appears earlier in the
///         method and is not nested inside a branch the write sits outside of is one every path to
///         that write passes through. That is what is required now.
///     </para>
///     <para>
///         Still not a data-flow proof, and two of the shapes §24.12 lists remain undecided here.
///         A budget consulted through a wrapper this cannot see through — a helper that checks it
///         on the method's behalf — reads as no budget at all, which is why
///         <see cref="BoundedByCaller" /> exists and is checked for staleness. And a guard whose
///         condition is always false in some way other than the literal is still a guard by this
///         rule. Both need a semantic model and a real control-flow graph; what is here is a
///         syntactic bound on where the guard sits, stated as that rather than as more.
///     </para>
/// </summary>
public class FanOutBudgetScopeTests : TestBase
{
    /// <summary>The ways a handler in this codebase bounds what it writes.</summary>
    private static readonly string[] KnownBudgets =
    [
        "RenderBudget", "BoundedFileBatch", "BoundedFilePublisher",
        "MaxTotalWorkUnits", "MaxExtractAllBytes", "PixelBudget"
    ];

    /// <summary>Method names whose disk writes are bounded by their caller, with the reason.</summary>
    private static readonly (string File, string Method, string Why)[] BoundedByCaller = [];

    /// <summary>Whether a path is production source rather than a test or a build artefact.</summary>
    /// <param name="path">The file to judge.</param>
    /// <returns><c>true</c> when it is part of the server.</returns>
    private static bool IsProductionSource(string path)
    {
        var separator = Path.DirectorySeparatorChar;
        return !path.Contains($"{separator}Tests{separator}", StringComparison.Ordinal)
               && !path.Contains($"{separator}obj{separator}", StringComparison.Ordinal)
               && !path.Contains($"{separator}bin{separator}", StringComparison.Ordinal);
    }

    /// <summary>Whether a node names one of the budgets as code rather than in prose.</summary>
    /// <param name="node">The node to read.</param>
    /// <returns><c>true</c> when this node is a reference to a budget.</returns>
    /// <remarks>
    ///     Read from the syntax tree, not from the method's text. A substring search counted a
    ///     comment or a string literal that mentioned a budget as using one, which is the failure
    ///     mode of every text-based guard this suite has replaced (R13-T02).
    /// </remarks>
    private static bool IsABudgetReference(SyntaxNode node)
    {
        var name = node switch
        {
            IdentifierNameSyntax identifier => identifier.Identifier.ValueText,
            MemberAccessExpressionSyntax member => member.Name.Identifier.ValueText,
            GenericNameSyntax generic => generic.Identifier.ValueText,
            ObjectCreationExpressionSyntax creation => creation.Type.ToString(),
            _ => string.Empty
        };

        return KnownBudgets.Contains(name, StringComparer.Ordinal);
    }

    /// <summary>Whether a method names one of the budgets anywhere.</summary>
    /// <param name="declaration">The method to inspect.</param>
    /// <returns><c>true</c> when a budget appears as an identifier in the method.</returns>
    private static bool UsesABudget(MethodDeclarationSyntax declaration)
    {
        return declaration.DescendantNodes().Any(IsABudgetReference);
    }

    /// <summary>Whether every disk write in a loop has a budget ahead of it on every path.</summary>
    /// <param name="declaration">The method to inspect.</param>
    /// <returns><c>true</c> when each such write is guarded.</returns>
    private static bool EveryWriteIsGuarded(MethodDeclarationSyntax declaration)
    {
        var budgets = declaration.DescendantNodes().Where(IsABudgetReference).ToList();
        if (budgets.Count == 0) return false;

        var writes = declaration.DescendantNodes().OfType<InvocationExpressionSyntax>()
            .Where(invocation => IsDiskWrite(invocation) && InsideALoop(invocation))
            .ToList();

        return writes.All(write => budgets.Any(budget => Bounds(budget, write)));
    }

    /// <summary>Whether a budget bounds a write: passed to it, or reached before it every time.</summary>
    /// <param name="guard">The budget reference.</param>
    /// <param name="write">The disk write it is supposed to bound.</param>
    /// <returns><c>true</c> when nothing can reach the write without the guard.</returns>
    /// <remarks>
    ///     Two ways to bound a write, and this codebase uses both.
    ///     <para>
    ///         The budget is handed to the writer —
    ///         <c>BoundedFilePublisher.Publish(name, RenderBudget.MaxOutputBytes, ...)</c>. It does
    ///         not precede the call, it is part of it, so a rule about source order alone rejected
    ///         every one of them.
    ///     </para>
    ///     <para>
    ///         Or the budget is checked earlier and the write is only reached if it passes. Source
    ///         order decides the first half of that, because statements run in order; the second
    ///         half is that nothing between the two gives execution a way past the check. Being
    ///         inside an <c>if</c> is not by itself such a way — <c>if (units &gt; Max) throw</c>
    ///         evaluates its condition on every path, and only the <em>body</em> of a branch or a
    ///         loop is what may be skipped.
    ///     </para>
    ///     <para>
    ///         Still syntax, not data flow: a budget read into a variable that is then never
    ///         consulted satisfies this, and a budget consulted inside a helper this cannot see
    ///         through does not.
    ///     </para>
    /// </remarks>
    private static bool Bounds(SyntaxNode guard, SyntaxNode write)
    {
        // Handed to the writer.
        if (write.Span.Contains(guard.Span)) return true;

        if (guard.SpanStart >= write.SpanStart) return false;

        var common = guard.Ancestors().FirstOrDefault(ancestor => ancestor.Span.Contains(write.Span));
        if (common == null) return false;

        for (var node = guard; node.Parent != null && node.Parent != common; node = node.Parent)
            if (IsSkippablePart(node.Parent, node))
                return false;

        return true;
    }

    /// <summary>Whether a child sits in the part of its parent that a path can go around.</summary>
    /// <param name="parent">The enclosing construct.</param>
    /// <param name="child">The part of it the guard is in.</param>
    /// <returns><c>true</c> when execution can reach past the parent without entering the child.</returns>
    private static bool IsSkippablePart(SyntaxNode parent, SyntaxNode child)
    {
        return parent switch
        {
            // The condition runs on every path; the branches do not.
            IfStatementSyntax branch => child == branch.Statement || child == branch.Else,
            ElseClauseSyntax => true,
            SwitchSectionSyntax => true,
            SwitchExpressionArmSyntax => true,
            CatchClauseSyntax => true,
            ConditionalExpressionSyntax ternary => child == ternary.WhenTrue || child == ternary.WhenFalse,

            // A loop the write sits outside of: its body may run no times at all.
            ForStatementSyntax loop => child == loop.Statement,
            ForEachStatementSyntax loop => child == loop.Statement,
            WhileStatementSyntax loop => child == loop.Statement,
            DoStatementSyntax loop => child == loop.Statement,

            // Written here, called somewhere this cannot see.
            AnonymousFunctionExpressionSyntax => true,
            LocalFunctionStatementSyntax => true,
            _ => false
        };
    }

    /// <summary>
    ///     Walks up from the test binaries to the repository root.
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

    /// <summary>Whether an invocation puts bytes on disk at a path.</summary>
    /// <param name="invocation">The call to classify.</param>
    /// <returns><c>true</c> when it writes a file.</returns>
    private static bool IsDiskWrite(InvocationExpressionSyntax invocation)
    {
        if (invocation.Expression is not MemberAccessExpressionSyntax member) return false;

        var name = member.Name.Identifier.ValueText;

        // `Save`/`ToImage` with a stream argument writes nowhere by itself; with a path it does.
        if (name is "Save" or "ToImage")
            return invocation.ArgumentList.Arguments.Count > 0
                   && !invocation.ArgumentList.Arguments[0].ToString().Contains("Stream",
                       StringComparison.OrdinalIgnoreCase);

        // `Copy` only writes a file when it is `File.Copy`; on a range or a collection it is an
        // in-memory operation, and matching the bare name reported a spreadsheet merge as a
        // fan-out that writes files.
        if (name == "Copy")
            return member.Expression is IdentifierNameSyntax { Identifier.ValueText: "File" };

        return name is "WriteAllBytes" or "WriteAllText" or "WriteAllLines" or "ExtractToFile"
            or "Publish";
    }

    /// <summary>Whether a node sits inside a loop.</summary>
    /// <param name="node">The node to place.</param>
    /// <returns><c>true</c> when a loop encloses it.</returns>
    private static bool InsideALoop(SyntaxNode node)
    {
        return node.Ancestors().Any(ancestor =>
            ancestor is ForStatementSyntax or ForEachStatementSyntax or WhileStatementSyntax
                or DoStatementSyntax);
    }

    /// <summary>Every production method that writes to disk from inside a loop.</summary>
    /// <returns>The file, the method, and its declaration.</returns>
    private static List<(string File, string Method, MethodDeclarationSyntax Declaration)> FanOutMethods()
    {
        var root = RepositoryRoot();
        var found = new List<(string, string, MethodDeclarationSyntax)>();

        foreach (var path in Directory
                     .EnumerateFiles(root.FullName, "*.cs", SearchOption.AllDirectories)
                     .Where(IsProductionSource))
        {
            var relative = Path.GetRelativePath(root.FullName, path).Replace('\\', '/');
            var unit = CSharpSyntaxTree.ParseText(File.ReadAllText(path)).GetCompilationUnitRoot();

            foreach (var method in unit.DescendantNodes().OfType<MethodDeclarationSyntax>())
            {
                var writesInALoop = method.DescendantNodes()
                    .OfType<InvocationExpressionSyntax>()
                    .Any(invocation => IsDiskWrite(invocation) && InsideALoop(invocation));

                if (writesInALoop) found.Add((relative, method.Identifier.ValueText, method));
            }
        }

        return found;
    }

    [Fact]
    public void EveryMethodWritingFilesInALoop_ShouldConsultABudgetInThatSameMethod()
    {
        var boundedElsewhere = BoundedByCaller
            .Select(entry => (entry.File, entry.Method))
            .ToHashSet();

        var unbounded = new List<string>();

        foreach (var (file, method, declaration) in FanOutMethods())
        {
            if (boundedElsewhere.Contains((file, method))) continue;

            if (!EveryWriteIsGuarded(declaration)) unbounded.Add($"{file}::{method}");
        }

        Assert.True(unbounded.Count == 0,
            "these methods write files in a loop with no budget named in the method itself: "
            + string.Join(", ", unbounded)
            + ". Give it one of " + string.Join("/", KnownBudgets)
            + ", or list it in BoundedByCaller with the reason.");
    }

    [Fact]
    public void TheBoundedByCallerList_ShouldNotHoldNamesThatNoLongerApply()
    {
        // A list of exceptions that outlives the code it excuses becomes a way to hide new ones.
        var actual = FanOutMethods().Select(entry => (entry.File, entry.Method)).ToHashSet();

        var stale = BoundedByCaller
            .Where(entry => !actual.Contains((entry.File, entry.Method)))
            .Select(entry => $"{entry.File}::{entry.Method}")
            .ToList();

        Assert.True(stale.Count == 0,
            "these are listed as bounded by their caller but no longer write in a loop: "
            + string.Join(", ", stale));
    }

    [Fact]
    public void TheScan_ShouldReachBeyondTheHandlersDirectory()
    {
        // It scanned Handlers/ only, and the widening found one method it had never asked about:
        // DocumentConverter.ConvertExcelToImages, which writes one file per sheet. That one is
        // bounded, so nothing was wrong — but nothing had checked, and narrowing the scan back
        // would return to not checking (R13-T02).
        var outside = FanOutMethods()
            .Where(entry => !entry.File.StartsWith("Handlers/", StringComparison.Ordinal))
            .ToList();

        Assert.NotEmpty(outside);
    }

    [Fact]
    public void TheAnalyser_ShouldNotAcceptABudgetNamedOnlyInAComment()
    {
        // The substring search could not tell a budget from prose about one, which is how a
        // text-based guard passes a method that does nothing.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      // RenderBudget is not needed here.
                                      var note = "RenderBudget";
                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        var unit = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        var writes = unit.DescendantNodes().OfType<MethodDeclarationSyntax>()
            .Single(method => method.Identifier.ValueText == "Writes");

        Assert.Contains("RenderBudget", writes.ToString(), StringComparison.Ordinal);
        Assert.False(UsesABudget(writes));
    }

    /// <summary>Whether the analyser accepts a method as bounded.</summary>
    /// <param name="source">The C# to parse.</param>
    /// <param name="method">The method to judge.</param>
    /// <returns><c>true</c> when every write in a loop is bounded.</returns>
    private static bool Accepts(string source, string method)
    {
        var unit = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        return EveryWriteIsGuarded(unit.DescendantNodes().OfType<MethodDeclarationSyntax>()
            .Single(declaration => declaration.Identifier.ValueText == method));
    }

    [Fact]
    public void ABudgetConsultedAfterTheWrite_ShouldNotCount()
    {
        // A ceiling checked once the files are already on disk is a report, not a bound.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                          RenderBudget.EnsureOutputCount(paths.Length);
                                      }
                                  }
                              }
                              """;

        Assert.False(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetCheckedOnOnlyOneBranch_ShouldNotCount()
    {
        // The write is reached whichever way the branch goes; the check is not.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths, bool careful)
                                  {
                                      if (careful)
                                      {
                                          RenderBudget.EnsureOutputCount(paths.Length);
                                      }

                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        Assert.False(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetInsideAnIfThatIsNeverTaken_ShouldNotCount()
    {
        // `if (false)` is the degenerate case of the branch above, and it is the one §24.12 names.
        // No condition is read here — a guard in the body of any branch the write sits outside of
        // is a guard the write can be reached without.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      if (false)
                                      {
                                          RenderBudget.EnsureOutputCount(paths.Length);
                                      }

                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        Assert.False(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetInsideALambdaThatMayNeverRun_ShouldNotCount()
    {
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      Action check = () => RenderBudget.EnsureOutputCount(paths.Length);

                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        Assert.False(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetInABranchsConditionBeforeTheWrite_ShouldCount()
    {
        // The control, and the codebase's own shape: `if (over the limit) throw`. The condition
        // runs on every path, so treating "inside an if" as skippable would report every one of
        // these as unbounded.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      if (paths.Length > MaxTotalWorkUnits)
                                          throw new ArgumentException("too many");

                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        Assert.True(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetHandedToTheWriter_ShouldCount()
    {
        // The other shape this codebase uses. The budget does not precede the call; it is an
        // argument to it.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      foreach (var path in paths)
                                      {
                                          BoundedFilePublisher.Publish(path, RenderBudget.MaxOutputBytes,
                                              stream => image.Save(stream), "an image", Recovery, []);
                                      }
                                  }
                              }
                              """;

        Assert.True(Accepts(source, "Writes"));
    }

    [Fact]
    public void ABudgetConsultedThroughAHelper_IsNotDecidedHere()
    {
        // Recorded rather than asserted away. A method that delegates its bounding to a helper
        // reads as unbounded, because deciding otherwise needs to resolve the callee — a semantic
        // model this analyser does not build (R13-T02). The consequence is a false positive, not a
        // false negative: such a method must be listed in BoundedByCaller with its reason, and
        // TheBoundedByCallerList_ShouldNotHoldNamesThatNoLongerApply keeps that list honest.
        const string source = """
                              class C
                              {
                                  void Writes(string[] paths)
                                  {
                                      EnsureWithinBudget(paths);

                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        Assert.False(Accepts(source, "Writes"));
    }

    [Fact]
    public void TheAnalyser_ShouldSeeABudgetInOneMethodAsAbsentFromAnother()
    {
        // The exact drift the file-level scan cannot see: the budget is in the file, and the
        // method that writes does not use it.
        const string source = """
                              class C
                              {
                                  void Bounded()
                                  {
                                      RenderBudget.EnsureOutputCount(1);
                                  }

                                  void Writes(string[] paths)
                                  {
                                      foreach (var path in paths)
                                      {
                                          File.WriteAllBytes(path, new byte[0]);
                                      }
                                  }
                              }
                              """;

        var unit = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        var writes = unit.DescendantNodes().OfType<MethodDeclarationSyntax>()
            .Single(m => m.Identifier.ValueText == "Writes");

        Assert.Contains("RenderBudget", source, StringComparison.Ordinal);
        Assert.DoesNotContain("RenderBudget", writes.ToString(), StringComparison.Ordinal);
        Assert.Contains(writes.DescendantNodes().OfType<InvocationExpressionSyntax>(),
            invocation => IsDiskWrite(invocation) && InsideALoop(invocation));
    }

    [Fact]
    public void TheAnalyser_ShouldNotCountAStreamSaveAsADiskWrite()
    {
        // `Save(stream)` hands bytes to the caller; only `Save(path)` puts them on disk. Counting
        // both would fill the inventory with methods that write nothing.
        const string source = """
                              class C
                              {
                                  void M(string[] paths, System.IO.Stream outputStream)
                                  {
                                      foreach (var path in paths)
                                      {
                                          document.Save(outputStream);
                                      }
                                  }
                              }
                              """;

        var unit = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();
        var invocations = unit.DescendantNodes().OfType<InvocationExpressionSyntax>().ToList();

        Assert.DoesNotContain(invocations, IsDiskWrite);
    }

    [Fact]
    public void TheAnalyser_ShouldNotCountAWriteOutsideALoop()
    {
        // One file is not a fan-out. Treating it as one would demand a budget from every handler
        // that saves a single document.
        const string source = """
                              class C
                              {
                                  void M(string path)
                                  {
                                      File.WriteAllBytes(path, new byte[0]);
                                  }
                              }
                              """;

        var unit = CSharpSyntaxTree.ParseText(source).GetCompilationUnitRoot();

        Assert.DoesNotContain(unit.DescendantNodes().OfType<InvocationExpressionSyntax>(),
            invocation => IsDiskWrite(invocation) && InsideALoop(invocation));
    }

    [Fact]
    public void TheScan_ShouldFindSomethingToJudge()
    {
        // A scan matching nothing would make every assertion above vacuous.
        Assert.NotEmpty(FanOutMethods());
    }
}
