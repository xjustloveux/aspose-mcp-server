using System.Text;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Structural guard: every handler that rasterises a document must price the request first.
///     <para>
///         <see cref="AsposeMcpServer.Helpers.RenderBudget" /> was added for the conversion paths only, so the
///         handlers behind <c>word_render</c>, <c>excel_render</c> and <c>excel_data_import_export</c>
///         still accepted any page count at up to 1200 DPI with nothing checking the product
///         (R2-R01). A functional test cannot cover this cheaply — reproducing the failure means
///         actually rendering billions of pixels — so the guard reads the sources instead.
///     </para>
///     <para>
///         The question is asked per entry point rather than per file. A file-wide search for the
///         word <c>RenderBudget</c> was satisfied by one priced branch, so a second branch in the
///         same handler — the single-page and per-page paths of <c>word_render</c> are exactly that
///         shape — could rasterise with nothing checking its cost (R7-T01). Pricing reached through
///         a helper still counts, which is why the two relations are closed over the calls inside
///         the file.
///     </para>
/// </summary>
public class RenderBudgetWiringTests
{
    /// <summary>Assignments that set a raster resolution, i.e. the site where cost is decided.</summary>
    private static readonly Regex SetsResolution = new(
        @"\b(?:Resolution|HorizontalResolution|VerticalResolution)\s*=",
        RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>
    ///     Either way of pricing a render: the up-front estimate, or the running count of real
    ///     page geometry that replaced it where the actual dimensions are available (R3-R03).
    /// </summary>
    private static readonly Regex PricesTheRequest = new(
        @"RenderBudget\s*\.|new PixelBudget\s*\(",
        RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>A method signature, captured with its name and its modifiers.</summary>
    private static readonly Regex MethodStart = new(
        @"^[ \t]*(?:\[[^\]]*\][ \t]*)*(?<mods>(?:public|private|protected|internal)[^;{}()]*?)"
        + @"(?<name>[A-Za-z_][A-Za-z0-9_]*)\([^;{}]*\)[ \t]*$",
        RegexOptions.Compiled | RegexOptions.Multiline, TimeSpan.FromSeconds(10));

    /// <summary>A call, used to follow rasterising and pricing through the file's own helpers.</summary>
    private static readonly Regex Call = new(
        @"\b([A-Za-z_][A-Za-z0-9_]*)\s*\(", RegexOptions.Compiled, TimeSpan.FromSeconds(10));

    /// <summary>
    ///     Locates the production Handlers directory, ignoring the test project's own tree.
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

    /// <summary>Splits a file into its methods, keeping the name and modifiers of each.</summary>
    /// <param name="source">The file's text.</param>
    /// <returns>Each method's name, modifiers and body.</returns>
    private static List<(string Name, string Modifiers, string Body)> Methods(string source)
    {
        var methods = new List<(string, string, string)>();

        foreach (Match match in MethodStart.Matches(source))
        {
            var open = source.IndexOf('{', match.Index + match.Length);
            if (open < 0) continue;

            var depth = 0;
            for (var i = open; i < source.Length; i++)
            {
                if (source[i] == '{') depth++;
                else if (source[i] == '}') depth--;

                if (depth != 0) continue;

                methods.Add((match.Groups["name"].Value, match.Groups["mods"].Value,
                    source[open..(i + 1)]));
                break;
            }
        }

        return methods;
    }

    /// <summary>
    ///     The methods that match <paramref name="marker" /> directly, plus every method that
    ///     reaches one of those through a call inside the same file.
    /// </summary>
    /// <param name="methods">The file's methods.</param>
    /// <param name="marker">What to look for in a body.</param>
    /// <returns>The names of the matching methods.</returns>
    private static HashSet<string> ClosureOver(
        List<(string Name, string Modifiers, string Body)> methods, Regex marker)
    {
        var reached = new HashSet<string>(
            methods.Where(m => marker.IsMatch(m.Body)).Select(m => m.Name), StringComparer.Ordinal);

        bool changed;
        do
        {
            changed = false;
            foreach (var method in methods)
            {
                if (reached.Contains(method.Name)) continue;
                if (!Call.Matches(method.Body).Any(c => reached.Contains(c.Groups[1].Value))) continue;

                reached.Add(method.Name);
                changed = true;
            }
        } while (changed);

        return reached;
    }

    /// <summary>
    ///     Reports every entry point that reaches a raster resolution without reaching a price.
    /// </summary>
    /// <param name="source">The source text to analyse.</param>
    /// <param name="label">How to name the source in the report.</param>
    /// <returns>The offending entry points, and how many were examined.</returns>
    internal static (List<string> Offenders, int EntryPoints) UnpricedRasterisers(
        string source, string label = "")
    {
        var offenders = new List<string>();
        if (!SetsResolution.IsMatch(source)) return (offenders, 0);

        var methods = Methods(source);
        var rasterises = ClosureOver(methods, SetsResolution);
        var prices = ClosureOver(methods, PricesTheRequest);
        var entryPoints = 0;

        foreach (var (name, modifiers, _) in methods)
        {
            if (!rasterises.Contains(name)) continue;

            // A private helper is judged through its caller: the pricing belongs at the entry
            // point, not in the function that happens to fill in an options object.
            if (modifiers.Contains("private", StringComparison.Ordinal)) continue;

            entryPoints++;
            if (!prices.Contains(name)) offenders.Add($"{label} :: {name}");
        }

        return (offenders, entryPoints);
    }

    [Fact]
    public void EveryEntryPointThatRasterises_ShouldPriceTheRequestFirst()
    {
        var offenders = new List<string>();
        var entryPoints = 0;

        foreach (var file in Directory.GetFiles(HandlersRoot(), "*.cs", SearchOption.AllDirectories)
                     .OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            var (found, checkedHere) = UnpricedRasterisers(source, Path.GetFileName(file));

            offenders.AddRange(found);
            entryPoints += checkedHere;
        }

        Assert.True(entryPoints > 0,
            "No rasterising entry point was found; the scan is looking in the wrong place.");
        Assert.True(offenders.Count == 0,
            "These entry points set a raster resolution without pricing the request against "
            + "RenderBudget or PixelBudget, so page count multiplied by DPI is unbounded:"
            + Environment.NewLine + string.Join(Environment.NewLine, offenders));
    }

    [Fact]
    public void ASecondBranchThatSkipsPricing_ShouldBeFlagged()
    {
        // The false green this guard was rewritten for: one priced entry point made the whole file
        // look wired, and the unpriced one beside it was never asked (R7-T01).
        const string source = "internal void RenderOnePage(Document doc)\n"
                              + "{\n"
                              + "    RenderBudget.EnsureWithinLimits(doc.PageCount, dpi);\n"
                              + "    options.Resolution = dpi;\n"
                              + "}\n"
                              + "\n"
                              + "internal void RenderEveryPage(Document doc)\n"
                              + "{\n"
                              + "    options.Resolution = dpi;\n"
                              + "}\n";

        var (offenders, entryPoints) = UnpricedRasterisers(source, "synthetic");

        Assert.Equal(2, entryPoints);
        Assert.Single(offenders);
        Assert.Contains("RenderEveryPage", offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void PricingReachedThroughAHelper_ShouldBeAccepted()
    {
        const string source = "internal void Render(Document doc)\n"
                              + "{\n"
                              + "    Price(doc);\n"
                              + "    options.Resolution = dpi;\n"
                              + "}\n"
                              + "\n"
                              + "private void Price(Document doc)\n"
                              + "{\n"
                              + "    RenderBudget.EnsureWithinLimits(doc.PageCount, dpi);\n"
                              + "}\n";

        Assert.Empty(UnpricedRasterisers(source, "synthetic").Offenders);
    }

    [Fact]
    public void AHelperThatOnlyBuildsOptions_ShouldBeJudgedThroughItsCaller()
    {
        // The shape word_render actually has: the options helper sets the resolution and the entry
        // point does the pricing. Judging the helper on its own would report a defect that is not
        // there, and a guard that cries wolf gets loosened.
        const string source = "internal void Render(Document doc)\n"
                              + "{\n"
                              + "    RenderBudget.EnsureWithinLimits(doc.PageCount, dpi);\n"
                              + "    var options = CreateOptions(dpi);\n"
                              + "}\n"
                              + "\n"
                              + "private static ImageSaveOptions CreateOptions(int dpi)\n"
                              + "{\n"
                              + "    return new ImageSaveOptions { Resolution = dpi };\n"
                              + "}\n";

        var (offenders, entryPoints) = UnpricedRasterisers(source, "synthetic");

        Assert.Empty(offenders);
        Assert.Equal(1, entryPoints);
    }
}
