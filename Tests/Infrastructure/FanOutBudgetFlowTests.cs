using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.FlowAnalysis;
using Microsoft.CodeAnalysis.Operations;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     R13-T02, the bar §24.12 actually set: a semantic model for the symbols and a control-flow
///     graph for the paths, rather than names and source order.
///     <para>
///         <see cref="FanOutBudgetInventoryTests" /> asks whether a budget appears anywhere in the
///         file. <see cref="FanOutBudgetScopeTests" /> narrows that to the method and to a
///         syntactic position. Both are cheap and both decide by shape: a call named <c>Save</c>
///         is a disk write if its first argument does not read like a stream, a name spelled
///         <c>RenderBudget</c> is a budget, and a guard that appears earlier in the source is
///         assumed to be reached first.
///     </para>
///     <para>
///         Here neither is assumed. Every source file of the server is compiled together, so a
///         sink is a resolved <see cref="IMethodSymbol" /> — <c>Save(string)</c> is a write and
///         <c>Save(Stream)</c> is not because of their <em>parameter types</em>, not because of
///         how the argument was spelled — and a budget is a symbol declared on one of the budget
///         types. The method's control-flow graph then decides reachability: a write is bounded
///         only if the budget is evaluated as part of the write itself, or sits in a basic block
///         that <em>dominates</em> the write's block. Dominance is the property the earlier gates
///         approximate with source order, and unlike source order it is not fooled by a guard on
///         one branch, by an early return, or by a loop that may execute zero times.
///     </para>
///     <para>
///         Indirection is followed to a fixed point, not one level: a method that hands its
///         bounding to a helper, which hands it on again, is bounded. Answers are memoised and a
///         method already under evaluation answers "not yet known to", so a cycle terminates
///         instead of recursing — and that direction can only under-report a budget, which
///         produces a finding to examine rather than a clean result to trust. This is the shape
///         <see cref="FanOutBudgetScopeTests" /> records as undecidable, and it is decided here.
///     </para>
/// </summary>
public class FanOutBudgetFlowTests : TestBase
{
    /// <summary>Types whose members are ceilings on what a handler may write.</summary>
    private static readonly string[] BudgetTypes =
    [
        "RenderBudget", "BoundedFileBatch", "BoundedFilePublisher", "PixelBudget", "StagingBudget"
    ];

    /// <summary>Members that are a ceiling wherever they are declared.</summary>
    private static readonly string[] BudgetMembers = ["MaxTotalWorkUnits", "MaxExtractAllBytes"];

    /// <summary>
    ///     Methods that open a path in a loop to take a lock, not to put bytes on disk.
    ///     <para>
    ///         The rule below decides by shape: a <c>FileStream</c> constructed on a string path is
    ///         a write, because that is how a handler that hands a stream to <c>Save</c> actually
    ///         produces a file. Taking a lock has the same shape and none of the meaning — one
    ///         fixed path, retried until it opens, never written to, and at most one zero-byte file
    ///         for the life of the installation. A budget in front of it would bound nothing.
    ///     </para>
    ///     <para>
    ///         Written down rather than designed around: the alternative was to split the retry
    ///         loop across two methods so the shape stops matching, which hides the exception
    ///         instead of stating it. Kept to a named list, and asserted to stay short, so it
    ///         cannot quietly become the place findings go.
    ///     </para>
    /// </summary>
    private static readonly string[] LocksRatherThanWrites =
    [
        "CrossProcessFileGate.cs::TryAcquire",
        // The staged-input sweep opens each stale copy exclusively to prove nobody holds it
        // before deleting it (R21-REC05). Read access, no bytes written, one open per file.
        "CleanupDebtService.cs::SweepStagedInputs"
    ];

    /// <summary>Methods that put bytes on disk, named by their declaring type and name.</summary>
    private static readonly (string Type, string Method)[] DiskWrites =
    [
        ("System.IO.File", "WriteAllBytes"),
        ("System.IO.File", "WriteAllText"),
        ("System.IO.File", "WriteAllLines"),
        ("System.IO.File", "Copy"),
        ("System.IO.Compression.ZipFileExtensions", "ExtractToFile")
    ];

    /// <summary>What is already known about which methods consult a budget.</summary>
    /// <remarks>
    ///     Memoised across the whole run, which is what makes the transitive walk affordable and
    ///     what makes it terminate: a method under evaluation answers "not yet known to", so a
    ///     cycle resolves instead of recursing. That direction is the safe one — it can only
    ///     under-report a budget, and under-reporting produces a finding to look at rather than a
    ///     clean result to trust.
    /// </remarks>
    private static readonly Dictionary<ISymbol, bool> KnownConsumers =
        new(SymbolEqualityComparer.Default);

    /// <summary>The methods currently being evaluated, so recursion stops at a cycle.</summary>
    private static readonly HashSet<ISymbol> BeingEvaluated = new(SymbolEqualityComparer.Default);

    [Fact]
    public void EveryDiskWriteInALoop_ShouldBeDominatedByABudget()
    {
        var (compilation, unresolved) = ProductionCompilation();

        // A scan whose symbols did not resolve finds no sinks and reports the same clean result as
        // a codebase with none. Checked before anything is concluded from the absence of findings.
        Assert.True(unresolved.Count == 0,
            "these anchor symbols did not resolve, so the analysis would have seen no sinks at "
            + "all: " + string.Join(", ", unresolved));

        var analysed = 0;
        var unbounded = new List<string>();
        var unanalysable = new List<string>();

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            var file = Path.GetFileName(tree.FilePath);

            foreach (var declaration in tree.GetRoot().DescendantNodes()
                         .OfType<MethodDeclarationSyntax>())
            {
                if (declaration.Body == null && declaration.ExpressionBody == null) continue;
                if (model.GetOperation(declaration) is not IMethodBodyOperation body) continue;

                var writes = WritesInALoop(body).ToList();
                if (writes.Count == 0) continue;

                analysed++;
                var name = $"{file}::{declaration.Identifier.ValueText}";

                ControlFlowGraph graph;
                try
                {
                    graph = ControlFlowGraph.Create(body);
                }
                catch (Exception exception)
                {
                    unanalysable.Add($"{name} ({exception.GetType().Name})");
                    continue;
                }

                var carriers = BudgetCarryingLocals(body, compilation);

                if (LocksRatherThanWrites.Contains(name)) continue;

                if (writes.Any(write => !IsBounded(graph, write, compilation, carriers)))
                    unbounded.Add(name);
            }
        }

        Assert.True(analysed > 0,
            "no method was found writing to disk inside a loop, which is what a compilation that "
            + "resolved nothing also reports");

        Assert.True(unanalysable.Count == 0,
            "no control-flow graph could be built for these, so nothing was proved about them: "
            + string.Join(", ", unanalysable));

        Assert.True(unbounded.Count == 0,
            "these methods write files in a loop on a path that does not pass through a budget: "
            + string.Join(", ", unbounded));

        // The exemption list is the one place a finding can be made to disappear, so it is held to
        // a size a reader can check by eye, and every entry has to still exist.
        Assert.True(LocksRatherThanWrites.Length <= 3,
            "the lock-acquisition exemption has grown past the handful it was meant to be: "
            + string.Join(", ", LocksRatherThanWrites));

        Assert.All(LocksRatherThanWrites, exempt =>
            Assert.Contains(exempt.Split("::")[0], compilation.SyntaxTrees
                .Select(tree => Path.GetFileName(tree.FilePath))));
    }

    [Fact]
    public void TheFlowAnalysis_ShouldSeeMoreMethodsThanItReports()
    {
        // The count that makes the clean result above mean something. It is asserted separately so
        // a change that quietly stops finding fan-out methods fails by name rather than by leaving
        // the other test vacuously green.
        var (compilation, _) = ProductionCompilation();

        var found = compilation.SyntaxTrees
            .SelectMany(tree =>
            {
                var model = compilation.GetSemanticModel(tree);
                return tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
                    .Where(declaration => model.GetOperation(declaration) is IMethodBodyOperation body
                                          && WritesInALoop(body).Any())
                    .Select(declaration =>
                        $"{Path.GetFileName(tree.FilePath)}::{declaration.Identifier.ValueText}");
            })
            .ToList();

        // Measured, not guessed. The syntactic gate next door reports nine; this reports those
        // eight. The one it drops is `DocumentConverter.ConvertExcelToImages`, whose only in-loop
        // call is `sr.ToImage(0, stream)` — a write to a stream, counted there because the first
        // argument is spelled `0` and so does not read as one. Its actual publish happens after
        // the loop, through a bounded batch. Falling below eight means this stopped resolving
        // symbols, which is the failure that otherwise looks like a clean result.
        Assert.True(found.Count >= 8,
            $"only {found.Count} fan-out methods were found: {string.Join(", ", found)}");
    }

    /// <summary>Whether the flow analysis accepts a method written for this test.</summary>
    /// <param name="body">The method's source, compiled against the same references.</param>
    /// <returns><c>true</c> when every disk write in a loop is dominated by a budget.</returns>
    private static bool AcceptsSnippet(string body)
    {
        var source = $$"""
                       using System.IO;
                       static class RenderBudget
                       {
                           public const long MaxOutputBytes = 1024;
                           public static void EnsureOutputCount(int count) { }
                       }
                       class Subject
                       {
                           const int MaxTotalWorkUnits = 100;
                           {{body}}
                       }
                       """;

        var (production, _) = ProductionCompilation();
        var tree = CSharpSyntaxTree.ParseText(source, path: "Subject.cs");
        var compilation = CSharpCompilation.Create("FanOutFlowSnippet", [tree],
            production.References, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        var model = compilation.GetSemanticModel(tree);
        var declaration = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
            .Single(method => method.Identifier.ValueText == "Writes");

        var operation = Assert.IsType<IMethodBodyOperation>(model.GetOperation(declaration), false);
        var writes = WritesInALoop(operation).ToList();

        // Every one of these snippets writes in a loop. A snippet where the sink stopped being
        // recognised would otherwise pass every "should be rejected" case for the wrong reason.
        Assert.NotEmpty(writes);

        var graph = ControlFlowGraph.Create(operation);
        var carriers = BudgetCarryingLocals(operation, compilation);

        return writes.All(write => IsBounded(graph, write, compilation, carriers));
    }

    [Fact]
    public void ABudgetConsultedTwoCallsAway_ShouldStillBound()
    {
        // One level decided the common case and stopped there. A helper that hands the check on to
        // another helper is an ordinary shape, and it was the last thing this analysis could not
        // answer — reported as unbounded, which is a finding nobody should have to dismiss.
        Assert.True(AcceptsSnippet("""
                                   void Writes(string[] paths)
                                   {
                                       Outer(paths);
                                       foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                   }

                                   void Outer(string[] paths) => Inner(paths);

                                   void Inner(string[] paths) => RenderBudget.EnsureOutputCount(paths.Length);
                                   """));
    }

    [Fact]
    public void AHelperThatChecksNothing_ShouldStillLeaveTheWriteUnbounded()
    {
        // The control for the walk above. Following calls must not turn every call into a budget.
        Assert.False(AcceptsSnippet("""
                                    void Writes(string[] paths)
                                    {
                                        Outer(paths);
                                        foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                    }

                                    void Outer(string[] paths) => Inner(paths);

                                    void Inner(string[] paths) { }
                                    """));
    }

    [Fact]
    public void MutuallyRecursiveHelpers_ShouldNotHangTheWalk()
    {
        // Following calls to a fixed point means meeting cycles. A method under evaluation answers
        // "not yet known to consult one", which terminates and can only under-report — the
        // direction that produces a finding to look at rather than a clean result to trust.
        Assert.False(AcceptsSnippet("""
                                    void Writes(string[] paths)
                                    {
                                        Ping(paths, 3);
                                        foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                    }

                                    void Ping(string[] paths, int n) { if (n > 0) Pong(paths, n - 1); }

                                    void Pong(string[] paths, int n) { if (n > 0) Ping(paths, n - 1); }
                                    """));
    }

    [Fact]
    public void ACycleThatReachesABudget_ShouldStillBound()
    {
        // And the other half: a cycle is not a reason to stop looking. The budget is reachable, so
        // it is found, whichever order the two are evaluated in.
        Assert.True(AcceptsSnippet("""
                                   void Writes(string[] paths)
                                   {
                                       Ping(paths, 3);
                                       foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                   }

                                   void Ping(string[] paths, int n)
                                   {
                                       RenderBudget.EnsureOutputCount(paths.Length);
                                       if (n > 0) Pong(paths, n - 1);
                                   }

                                   void Pong(string[] paths, int n) { if (n > 0) Ping(paths, n - 1); }
                                   """));
    }

    [Fact]
    public void ABudgetOnOnlyOneBranch_ShouldNotDominateTheWrite()
    {
        Assert.False(AcceptsSnippet("""
                                    void Writes(string[] paths, bool careful)
                                    {
                                        if (careful) RenderBudget.EnsureOutputCount(paths.Length);
                                        foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                    }
                                    """));
    }

    [Fact]
    public void ABudgetAfterAnEarlyReturn_ShouldNotDominateAWriteBeforeIt()
    {
        // The shape source order cannot see: the guard is textually first and is reached on only
        // one of the two paths out of the branch above it.
        Assert.False(AcceptsSnippet("""
                                    void Writes(string[] paths, bool quick)
                                    {
                                        foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                        if (quick) return;
                                        RenderBudget.EnsureOutputCount(paths.Length);
                                    }
                                    """));
    }

    [Fact]
    public void ABudgetConsultedOnlyInsideAnotherLoop_ShouldNotDominate()
    {
        // A loop body may run zero times, so nothing inside it dominates what comes after.
        Assert.False(AcceptsSnippet("""
                                    void Writes(string[] paths)
                                    {
                                        foreach (var first in paths) RenderBudget.EnsureOutputCount(1);
                                        foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                    }
                                    """));
    }

    [Fact]
    public void ABudgetInAConditionBeforeTheWrite_ShouldDominateIt()
    {
        // The control, and this codebase's own shape. A condition is evaluated on every path
        // through the branch, so `if (over the limit) throw` does bound what follows.
        Assert.True(AcceptsSnippet("""
                                   void Writes(string[] paths)
                                   {
                                       if (paths.Length > MaxTotalWorkUnits) throw new System.ArgumentException("too many");
                                       foreach (var path in paths) File.WriteAllBytes(path, new byte[0]);
                                   }
                                   """));
    }

    [Fact]
    public void ABudgetReadIntoALocalBeforeTheLoop_ShouldDominateThroughThatLocal()
    {
        // The two extract-all handlers' shape, and the case that made this analysis worth doing
        // with a semantic model: the budget reference is inside a conditional access, so the block
        // holding it dominates nothing, while the guard that does dominate compares against the
        // local it was read into.
        Assert.True(AcceptsSnippet("""
                                   void Writes(string[] paths, string? config)
                                   {
                                       var cap = config == null ? RenderBudget.MaxOutputBytes : RenderBudget.MaxOutputBytes;
                                       long used = 0;
                                       foreach (var path in paths)
                                       {
                                           if (used > cap) break;
                                           File.WriteAllBytes(path, new byte[0]);
                                           used += 1;
                                       }
                                   }
                                   """));
    }

    /// <summary>Whether a budget is evaluated on every path that reaches this write.</summary>
    /// <param name="graph">The method's control-flow graph.</param>
    /// <param name="write">The write to bound.</param>
    /// <param name="compilation">The compilation, for resolving a helper's own body.</param>
    /// <param name="carriers">Symbols that carry a value read from a budget.</param>
    /// <returns><c>true</c> when the write cannot be reached without a budget.</returns>
    private static bool IsBounded(ControlFlowGraph graph, IOperation write,
        Compilation compilation, IReadOnlySet<ISymbol> carriers)
    {
        // Handed to the writer: `Publish(path, RenderBudget.MaxOutputBytes, …)`. The budget is
        // evaluated as part of building this very call, so there is no path that reaches the call
        // without it — and it does not precede the call in source order at all.
        if (write.Descendants().Any(operation => IsABudget(operation, compilation, carriers)))
            return true;

        var writeBlock = BlockOf(graph, write);
        if (writeBlock == null) return false;

        // In the same block, before the write. A basic block has one entry and one exit, so
        // ordering within it is execution order.
        if (BudgetOperations(writeBlock, compilation, carriers)
            .Any(budget => budget.Syntax.SpanStart < write.Syntax.SpanStart))
            return true;

        var dominators = Dominators(graph);
        return dominators[writeBlock.Ordinal]
            .Where(ordinal => ordinal != writeBlock.Ordinal)
            .Any(ordinal => BudgetOperations(graph.Blocks[ordinal], compilation, carriers).Any());
    }

    /// <summary>Every disk write this method performs from inside a loop.</summary>
    /// <param name="body">The method body operation.</param>
    /// <returns>The write invocations.</returns>
    private static IEnumerable<IOperation> WritesInALoop(IMethodBodyOperation body)
    {
        return body.Descendants()
            .Where(operation => operation switch
            {
                IInvocationOperation invocation => IsADiskWrite(invocation.TargetMethod),
                // A FileStream opened on a path creates the file. This is how a handler that hands
                // a stream to `Save` actually puts bytes on disk, and it is the operation a budget
                // has to sit in front of.
                IObjectCreationOperation creation => OpensAFile(creation),
                _ => false
            })
            .Where(InsideALoop);
    }

    /// <summary>Whether an operation sits inside a loop in the method's own body.</summary>
    /// <param name="operation">The operation to place.</param>
    /// <returns><c>true</c> when a loop encloses it.</returns>
    private static bool InsideALoop(IOperation operation)
    {
        for (var parent = operation.Parent; parent != null; parent = parent.Parent)
            if (parent is ILoopOperation)
                return true;

        return false;
    }

    /// <summary>Whether a resolved method puts bytes on disk at a path.</summary>
    /// <param name="method">The callee.</param>
    /// <returns><c>true</c> when calling it writes a file.</returns>
    private static bool IsADiskWrite(IMethodSymbol? method)
    {
        if (method == null) return false;

        var containing = method.ContainingType?.ToDisplayString() ?? string.Empty;

        if (DiskWrites.Any(sink => sink.Type == containing && sink.Method == method.Name))
            return true;

        // The bounded publishers are sinks in their own right, and bounded by construction.
        if (method is
            {
                Name: "Publish",
                ContainingType.Name: "BoundedFileBatch" or "BoundedFilePublisher"
            })
            return true;

        // The distinction the syntactic gate makes by reading the argument's text: `Save(path)`
        // writes a file, `Save(stream)` hands the bytes back. Decided here by the parameter types,
        // so `Save(fs)` and `ToImage(0, stream)` are correctly not writes, and
        // `ToImage(index, path)` — where the path is the *second* parameter — correctly is.
        if (method.Name is "Save" or "ToImage")
            return method.Parameters.Any(parameter =>
                       parameter.Type.SpecialType == SpecialType.System_String)
                   && !method.Parameters.Any(parameter => IsAStream(parameter.Type));

        return false;
    }

    /// <summary>Whether a constructed object opens a file at a path.</summary>
    /// <param name="creation">The construction to classify.</param>
    /// <returns><c>true</c> when it creates or truncates a file.</returns>
    private static bool OpensAFile(IObjectCreationOperation creation)
    {
        if (creation.Type?.Name is not ("FileStream" or "StreamWriter")) return false;

        return creation.Arguments.Length > 0
               && creation.Arguments[0].Parameter?.Type.SpecialType == SpecialType.System_String;
    }

    /// <summary>Whether a type is a stream.</summary>
    /// <param name="type">The type to test.</param>
    /// <returns><c>true</c> when it is <c>System.IO.Stream</c> or derives from it.</returns>
    private static bool IsAStream(ITypeSymbol type)
    {
        for (var current = type; current != null; current = current.BaseType)
            if (current.ToDisplayString() == "System.IO.Stream")
                return true;

        return false;
    }

    /// <summary>The budget references evaluated in one basic block.</summary>
    /// <param name="block">The block to read.</param>
    /// <param name="compilation">The compilation, for resolving a helper's own body.</param>
    /// <param name="carriers">Symbols that carry a value read from a budget.</param>
    /// <returns>The operations that consult a budget.</returns>
    private static IEnumerable<IOperation> BudgetOperations(BasicBlock block,
        Compilation compilation, IReadOnlySet<ISymbol> carriers)
    {
        var operations = block.Operations.SelectMany(operation =>
            operation.DescendantsAndSelf());

        if (block.BranchValue != null)
            operations = operations.Concat(block.BranchValue.DescendantsAndSelf());

        return operations.Where(operation => IsABudget(operation, compilation, carriers));
    }

    /// <summary>The locals in a method that hold a value read from a budget.</summary>
    /// <param name="body">The method body.</param>
    /// <param name="compilation">The compilation, for resolving a helper's own body.</param>
    /// <returns>Those local symbols.</returns>
    /// <remarks>
    ///     `var cap = config?.MaxExtractAllBytes ?? long.MaxValue;` puts the budget reference
    ///     inside a conditional access, so the block holding it does not dominate anything — while
    ///     the guard that <em>does</em> dominate the write compares against `cap`. Following the
    ///     value from the budget into the local is what makes the two the same fact. Flow-
    ///     insensitive on purpose: it can only recognise more guards, never fewer, and where the
    ///     guard sits is still decided by dominance.
    /// </remarks>
    private static IReadOnlySet<ISymbol> BudgetCarryingLocals(IMethodBodyOperation body,
        Compilation compilation)
    {
        var carriers = new HashSet<ISymbol>(SymbolEqualityComparer.Default);
        var empty = (IReadOnlySet<ISymbol>)new HashSet<ISymbol>(SymbolEqualityComparer.Default);

        foreach (var operation in body.Descendants())
        {
            var (target, value) = operation switch
            {
                IVariableInitializerOperation initializer =>
                    (Local(initializer), initializer.Value),
                ISimpleAssignmentOperation { Target: ILocalReferenceOperation local } assignment =>
                    (local.Local, assignment.Value),
                _ => (null, null)
            };

            if (target == null || value == null) continue;
            if (value.DescendantsAndSelf().Any(node => IsABudget(node, compilation, empty)))
                carriers.Add(target);
        }

        return carriers;
    }

    /// <summary>The single local a variable initializer belongs to, when there is one.</summary>
    /// <param name="initializer">The initializer.</param>
    /// <returns>The local symbol, or null.</returns>
    private static ISymbol? Local(IVariableInitializerOperation initializer)
    {
        return initializer.Parent is IVariableDeclaratorOperation declarator ? declarator.Symbol : null;
    }

    /// <summary>Whether an operation consults a budget, directly or through a helper.</summary>
    /// <param name="operation">The operation to classify.</param>
    /// <param name="compilation">The compilation, for reading a callee's own body.</param>
    /// <param name="carriers">Symbols that carry a value read from a budget.</param>
    /// <returns><c>true</c> when a ceiling is being consulted here.</returns>
    private static bool IsABudget(IOperation operation, Compilation compilation,
        IReadOnlySet<ISymbol> carriers)
    {
        // A local that was given a budget's value carries it: the guard that reads `cap` is the
        // same fact as the one that read `MaxExtractAllBytes`.
        if (operation is ILocalReferenceOperation local && carriers.Contains(local.Local))
            return true;

        var symbol = operation switch
        {
            IInvocationOperation invocation => invocation.TargetMethod,
            IMemberReferenceOperation member => member.Member,
            IObjectCreationOperation creation => creation.Type,
            _ => null
        };

        if (symbol == null) return false;

        var owner = symbol as ITypeSymbol ?? symbol.ContainingType;
        if (owner != null && BudgetTypes.Contains(owner.Name, StringComparer.Ordinal)) return true;
        if (BudgetMembers.Contains(symbol.Name, StringComparer.Ordinal)) return true;

        // One level of indirection: a method in this codebase that consults a budget bounds its
        // caller. Resolvable because the callee is a symbol with a declaration to read, which is
        // exactly what the syntactic gate cannot do.
        return operation is IInvocationOperation call && CalleeConsultsABudget(call, compilation);
    }

    /// <summary>Whether a called method consults a budget, however many calls away it is.</summary>
    /// <param name="call">The call to follow.</param>
    /// <param name="compilation">The compilation holding the callee's source.</param>
    /// <returns><c>true</c> when the callee, or something it calls, consults a ceiling.</returns>
    /// <remarks>
    ///     Followed to a fixed point rather than one level. One level decided the common case — a
    ///     method that hands its bounding to a helper — and said nothing about a helper that hands
    ///     it on again, which is an ordinary shape in this codebase and was the last thing this
    ///     analysis could not answer.
    /// </remarks>
    private static bool CalleeConsultsABudget(IInvocationOperation call, Compilation compilation)
    {
        return ConsultsABudget(call.TargetMethod, compilation);
    }

    /// <summary>Whether a method, or anything it calls, consults a budget.</summary>
    /// <param name="method">The method to inspect.</param>
    /// <param name="compilation">The compilation holding its source.</param>
    /// <returns><c>true</c> when a ceiling is consulted somewhere below it.</returns>
    private static bool ConsultsABudget(IMethodSymbol method, Compilation compilation)
    {
        if (KnownConsumers.TryGetValue(method, out var known)) return known;
        if (!BeingEvaluated.Add(method)) return false;

        try
        {
            var answer = Consults(method, compilation);
            KnownConsumers[method] = answer;
            return answer;
        }
        finally
        {
            BeingEvaluated.Remove(method);
        }
    }

    /// <summary>Reads a method's body for a budget, directly or through what it calls.</summary>
    /// <param name="method">The method to read.</param>
    /// <param name="compilation">The compilation holding its source.</param>
    /// <returns><c>true</c> when a ceiling is consulted.</returns>
    private static bool Consults(IMethodSymbol method, Compilation compilation)
    {
        var declaration = method.DeclaringSyntaxReferences.FirstOrDefault()?.GetSyntax();
        if (declaration == null || !compilation.ContainsSyntaxTree(declaration.SyntaxTree))
            return false;

        var model = compilation.GetSemanticModel(declaration.SyntaxTree);
        if (model.GetOperation(declaration) is not { } body) return false;

        foreach (var operation in body.DescendantsAndSelf())
        {
            var symbol = operation switch
            {
                IMemberReferenceOperation member => member.Member,
                IInvocationOperation invocation => invocation.TargetMethod,
                IObjectCreationOperation creation => creation.Type,
                _ => null
            };

            if (symbol == null) continue;

            var owner = symbol as ITypeSymbol ?? symbol.ContainingType;
            if (owner != null && BudgetTypes.Contains(owner.Name, StringComparer.Ordinal))
                return true;
            if (BudgetMembers.Contains(symbol.Name, StringComparer.Ordinal)) return true;

            if (operation is IInvocationOperation next
                && ConsultsABudget(next.TargetMethod, compilation))
                return true;
        }

        return false;
    }

    /// <summary>The basic block an operation belongs to.</summary>
    /// <param name="graph">The control-flow graph.</param>
    /// <param name="operation">The operation to locate.</param>
    /// <returns>Its block, or null when the graph does not hold it.</returns>
    private static BasicBlock? BlockOf(ControlFlowGraph graph, IOperation operation)
    {
        var span = operation.Syntax.Span;

        return graph.Blocks.FirstOrDefault(block =>
            block.Operations.Any(candidate =>
                candidate.DescendantsAndSelf().Any(node => node.Syntax.Span == span))
            || (block.BranchValue?.DescendantsAndSelf()
                .Any(node => node.Syntax.Span == span) ?? false));
    }

    /// <summary>The blocks that dominate each block, by ordinal.</summary>
    /// <param name="graph">The control-flow graph.</param>
    /// <returns>For each block ordinal, the set of ordinals that dominate it.</returns>
    /// <remarks>
    ///     The standard iterative fixed point: a block is dominated by itself and by everything
    ///     that dominates all of its predecessors. Entry blocks have no predecessors and so are
    ///     dominated only by themselves, which makes an unreachable block dominated by itself
    ///     alone — conservative in the right direction, since it can then only be bounded by a
    ///     budget in its own block.
    /// </remarks>
    private static List<HashSet<int>> Dominators(ControlFlowGraph graph)
    {
        var count = graph.Blocks.Length;
        var all = Enumerable.Range(0, count).ToHashSet();
        var dominators = new List<HashSet<int>>(count);

        for (var i = 0; i < count; i++)
            dominators.Add(graph.Blocks[i].Predecessors.Length == 0 ? [i] : [..all]);

        bool changed;
        do
        {
            changed = false;

            for (var i = 0; i < count; i++)
            {
                var block = graph.Blocks[i];
                if (block.Predecessors.Length == 0) continue;

                HashSet<int>? intersection = null;
                foreach (var predecessor in block.Predecessors)
                {
                    var incoming = dominators[predecessor.Source.Ordinal];
                    if (intersection == null) intersection = [..incoming];
                    else intersection.IntersectWith(incoming);
                }

                intersection ??= [];
                intersection.Add(i);

                if (intersection.SetEquals(dominators[i])) continue;

                dominators[i] = intersection;
                changed = true;
            }
        } while (changed);

        return dominators;
    }

    /// <summary>Compiles every production source together, with the assemblies it references.</summary>
    /// <returns>The compilation, and any anchor symbol that failed to resolve.</returns>
    internal static (CSharpCompilation Compilation, List<string> Unresolved) ProductionCompilation()
    {
        var root = RepositoryRoot();
        var separator = Path.DirectorySeparatorChar;

        var sources = Directory
            .EnumerateFiles(root.FullName, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.Contains($"{separator}Tests{separator}", StringComparison.Ordinal)
                           && !path.Contains($"{separator}obj{separator}", StringComparison.Ordinal)
                           && !path.Contains($"{separator}bin{separator}", StringComparison.Ordinal))
            .ToList();

        // The project enables implicit usings, so `File`, `Path` and `List<T>` are declared in a
        // file the build generates under `obj/` — which this scan otherwise excludes. Compiling
        // without it leaves `File` an unknown name, `File.WriteAllBytes(...)` binds to nothing,
        // and the analysis sees a codebase with no disk writes in it and calls that clean.
        sources.AddRange(Directory
            .EnumerateFiles(root.FullName, "*.GlobalUsings.g.cs", SearchOption.AllDirectories)
            .Where(path => !path.Contains($"{separator}Tests{separator}", StringComparison.Ordinal))
            .Take(1));

        var trees = sources
            .Select(path => CSharpSyntaxTree.ParseText(File.ReadAllText(path), path: path))
            .ToList();

        // The framework first, then only those assemblies beside the test host that the framework
        // does not already provide — the Aspose and NuGet ones. Referencing both directories whole
        // defines every BCL type twice, and a call to `File.WriteAllBytes` then binds to nothing:
        // it comes back as an invalid operation rather than an invocation, and this analysis sees
        // a codebase with no disk writes in it. Diagnostics are not required to be clean — the
        // project's own NuGet references are not here — but the anchor symbols must resolve, and
        // the caller asserts that.
        // The runtime, and the web framework beside it: this is an ASP.NET Core server, and
        // without those assemblies every file that mentions `Microsoft.AspNetCore` fails to bind.
        var runtime = Path.GetDirectoryName(typeof(object).Assembly.Location)!;
        var framework = Directory.EnumerateFiles(runtime, "*.dll").ToList();
        framework.AddRange(AspNetCoreAssemblies(runtime));
        var provided = framework.Select(Path.GetFileName).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var references = framework
            .Concat(Directory.EnumerateFiles(AppContext.BaseDirectory, "*.dll")
                .Where(path => !provided.Contains(Path.GetFileName(path))))
            .Where(IsManagedAssembly)
            .Select(TryReference)
            .OfType<MetadataReference>()
            .ToList();

        // The project references System.Drawing.Common under an extern alias, because `Encoder`
        // exists in both it and Aspose.Slides. A compilation that does not reproduce the alias
        // gets the collision instead, and the Aspose types stop binding — which is how three
        // handlers that write through `Save(string)` went unseen.
        for (var i = 0; i < references.Count; i++)
            if (references[i] is PortableExecutableReference { FilePath: { } file }
                && Path.GetFileName(file).Equals("System.Drawing.Common.dll", StringComparison.OrdinalIgnoreCase))
                references[i] = references[i].WithAliases(["SysDrawing"]);

        var compilation = CSharpCompilation.Create("FanOutFlowScan", trees, references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        var unresolved = new List<string>();
        foreach (var (type, _) in DiskWrites.DistinctBy(sink => sink.Type))
            if (compilation.GetTypeByMetadataName(type) == null)
                unresolved.Add(type);

        return (compilation, unresolved);
    }

    /// <summary>The ASP.NET Core shared framework that matches the running runtime.</summary>
    /// <param name="runtime">The runtime's own shared-framework directory.</param>
    /// <returns>Its assemblies, or nothing when the framework is not installed beside it.</returns>
    private static IEnumerable<string> AspNetCoreAssemblies(string runtime)
    {
        var shared = Directory.GetParent(runtime)?.Parent;
        var web = shared == null
            ? null
            : new DirectoryInfo(Path.Combine(shared.FullName, "Microsoft.AspNetCore.App"));

        if (web is not { Exists: true }) return [];

        // The highest version present, which is the one a build against this runtime would pick.
        var version = web.GetDirectories()
            .OrderByDescending(directory => directory.Name, StringComparer.Ordinal)
            .FirstOrDefault();

        return version == null ? [] : Directory.EnumerateFiles(version.FullName, "*.dll");
    }

    /// <summary>Whether a file is a managed assembly and can be read as metadata.</summary>
    /// <param name="path">The file to test.</param>
    /// <returns><c>true</c> when it carries managed metadata.</returns>
    /// <remarks>
    ///     The Aspose packages ship native helper DLLs beside the managed ones, and so does the
    ///     shared framework. Referencing one produces a compilation-time error rather than an
    ///     exception at load, so it is filtered here instead of being caught later.
    /// </remarks>
    private static bool IsManagedAssembly(string path)
    {
        try
        {
            AssemblyName.GetAssemblyName(path);
            return true;
        }
        catch (Exception exception) when (exception is BadImageFormatException or IOException
                                              or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>Reads one assembly as a reference, skipping what cannot be read as metadata.</summary>
    /// <param name="path">The assembly path.</param>
    /// <returns>The reference, or null.</returns>
    private static MetadataReference? TryReference(string path)
    {
        try
        {
            return MetadataReference.CreateFromFile(path);
        }
        catch (Exception exception) when (exception is IOException or BadImageFormatException
                                              or UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>Walks up from the test binaries to the repository root.</summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }
}
