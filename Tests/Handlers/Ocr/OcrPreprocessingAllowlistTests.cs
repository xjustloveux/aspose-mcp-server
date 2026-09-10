using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Ocr.Preprocessing;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Ocr;

/// <summary>
///     R19-OCR01 and R19-OCR02: the OCR preprocessing paths were never authorised.
///     <para>
///         <c>ExtractCommonParameters</c> called <c>ValidateFilePath</c> and <c>File.Exists</c>.
///         The first says a string is well formed; the second says something is there. Neither says
///         the file is one this server may touch, and the input reached <c>OcrInput.Add</c> on the
///         caller's own spelling — so any readable path on the machine could be read through a
///         server whose allowlist said otherwise.
///     </para>
///     <para>
///         The output <em>was</em> resolved, but only after its parent directory had been created,
///         so a request that was about to be refused had already made a directory wherever the
///         caller pointed it.
///     </para>
///     <para>
///         Every operation is driven, not one. The check lives in the shared base, and the way this
///         repository keeps failing is that the case someone named gets fixed and the ones beside
///         it do not.
///     </para>
/// </summary>
public class OcrPreprocessingAllowlistTests : HandlerTestBase<AsposeOcr>
{
    /// <summary>Every preprocessing handler, so none of them is left untested.</summary>
    /// <returns>One row per operation name.</returns>
    public static TheoryData<string> EveryOperation()
    {
        var data = new TheoryData<string>();
        foreach (var operation in new[]
                     { "auto_skew", "contrast", "denoise", "dewarp", "invert", "scale" })
            data.Add(operation);

        return data;
    }

    /// <summary>The handler for one operation name.</summary>
    /// <param name="operation">The operation.</param>
    /// <returns>The handler.</returns>
    private static OcrPreprocessingHandlerBase HandlerFor(string operation)
    {
        return operation switch
        {
            "auto_skew" => new AutoSkewOcrPreprocessingHandler(),
            "contrast" => new ContrastOcrPreprocessingHandler(),
            "denoise" => new DenoiseOcrPreprocessingHandler(),
            "dewarp" => new DewarpOcrPreprocessingHandler(),
            "invert" => new InvertOcrPreprocessingHandler(),
            "scale" => new ScaleOcrPreprocessingHandler(),
            _ => throw new ArgumentOutOfRangeException(nameof(operation), operation, "unknown operation")
        };
    }

    /// <summary>A context whose allowlist is one directory.</summary>
    /// <param name="allowedRoot">The only directory a request may touch.</param>
    /// <returns>The context.</returns>
    private static OperationContext<AsposeOcr> AContext(string allowedRoot)
    {
        return new OperationContext<AsposeOcr>
        {
            Document = new AsposeOcr(),
            ServerConfig = ServerConfig.LoadFromArgs(["--allowed-path", allowedRoot])
        };
    }

    /// <summary>The allowed directory and a sibling outside it.</summary>
    /// <returns>The allowed root and the outside root.</returns>
    private (string Allowed, string Outside) TwoRoots()
    {
        return (Directory.CreateDirectory(Path.Combine(TestDir, "allowed")).FullName,
            Directory.CreateDirectory(Path.Combine(TestDir, "outside")).FullName);
    }

    /// <summary>Puts a real image at a path.</summary>
    /// <param name="path">Where it goes.</param>
    private void PlaceAnImage(string path)
    {
        File.Copy(CreateTempImageFile(), path, true);
    }

    [Theory]
    [MemberData(nameof(EveryOperation))]
    public void AnInputOutsideTheAllowlist_ShouldBeRefusedBeforeItIsRead(string operation)
    {
        var (allowed, outside) = TwoRoots();

        var input = Path.Combine(outside, "someone_elses.bmp");
        PlaceAnImage(input);

        var output = Path.Combine(allowed, "out.bmp");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", input },
            { "outputPath", output }
        });

        Assert.Throws<ArgumentException>(() =>
            HandlerFor(operation).Execute(AContext(allowed), parameters));

        Assert.False(File.Exists(output), "a refused request still produced an output");
    }

    [Theory]
    [MemberData(nameof(EveryOperation))]
    public void AnOutputOutsideTheAllowlist_ShouldBeRefusedWithoutCreatingItsDirectory(
        string operation)
    {
        var (allowed, outside) = TwoRoots();

        var input = Path.Combine(allowed, "ours.bmp");
        PlaceAnImage(input);

        // R19-OCR02. This directory was created before the output path was resolved, so a refused
        // request left it behind wherever the caller named.
        var planted = Path.Combine(outside, "made-by-a-refused-request");

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", input },
            { "outputPath", Path.Combine(planted, "out.bmp") }
        });

        Assert.Throws<ArgumentException>(() =>
            HandlerFor(operation).Execute(AContext(allowed), parameters));

        Assert.False(Directory.Exists(planted),
            "a refused request created a directory outside the allowlist");
    }

    [SkippableTheory]
    [MemberData(nameof(EveryOperation))]
    public void AnInputReachedThroughALinkOutOfTheAllowlist_ShouldBeRefused(string operation)
    {
        // A directory junction rather than a file symlink: the leaf itself is not a link, so this
        // is the mid-path redirection a leaf-only check cannot see, and it needs no privilege.
        var (allowed, outside) = TwoRoots();

        var real = Path.Combine(outside, "target.bmp");
        PlaceAnImage(real);

        var doorway = Path.Combine(allowed, "doorway");
        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(doorway, outside),
            "This machine cannot create a directory link.");

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(doorway, "target.bmp") },
            { "outputPath", Path.Combine(allowed, "out.bmp") }
        });

        Assert.Throws<ArgumentException>(() =>
            HandlerFor(operation).Execute(AContext(allowed), parameters));
    }

    [SkippableTheory]
    [MemberData(nameof(EveryOperation))]
    public void AnOutputReachedThroughALinkOutOfTheAllowlist_ShouldBeRefused(string operation)
    {
        var (allowed, outside) = TwoRoots();

        var input = Path.Combine(allowed, "ours.bmp");
        PlaceAnImage(input);

        var doorway = Path.Combine(allowed, "out-doorway");
        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(doorway, outside),
            "This machine cannot create a directory link.");

        var escaped = Path.Combine(doorway, "written_outside.bmp");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", input },
            { "outputPath", escaped }
        });

        Assert.Throws<ArgumentException>(() =>
            HandlerFor(operation).Execute(AContext(allowed), parameters));

        Assert.False(File.Exists(Path.Combine(outside, "written_outside.bmp")),
            "the output was written through the link anyway");
    }
}
