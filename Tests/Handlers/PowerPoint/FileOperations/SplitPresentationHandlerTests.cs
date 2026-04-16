using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

[SupportedOSPlatform("windows")]
public class SplitPresentationHandlerTests : PptHandlerTestBase
{
    private readonly SplitPresentationHandler _handler = new();
    private readonly string _inputPath;

    public SplitPresentationHandlerTests()
    {
        _inputPath = Path.Combine(TestDir, "input.pptx");

        using var pres = new Presentation();
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Save(_inputPath, SaveFormat.Pptx);
    }

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Split()
    {
        SkipIfNotWindows();
        Assert.Equal("split", _handler.Operation);
    }

    #endregion

    #region Basic Split Operations

    [SkippableFact]
    public void Execute_SplitsPresentation()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_output");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3", result.Message);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(3, files.Length);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Split file {file} should have content");
        }
    }

    [SkippableFact]
    public void Execute_WithPath_SplitsPresentation()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_path");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.NotEmpty(files);
    }

    [SkippableFact]
    public void Execute_WithSlidesPerFile_SplitsPresentation()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_multi");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
    }

    [SkippableFact]
    public void Execute_WithOutputFileNamePattern_UsesPattern()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_pattern");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "presentation_{index}.pptx" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(Path.Combine(outputDir, "presentation_0.pptx")));
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSource_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDirectory", Path.Combine(TestDir, "output") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithoutOutputDirectory_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideRange_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "split_invalid");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "startSlideIndex", 5 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    // Guard for bug 20260415-ppt-split-output-explosion: when the requested
    // split would produce more than 1000 output files, the handler must fail
    // fast (before writing anything to disk) instead of spamming the output
    // directory. 1001 slides with slidesPerFile=1 yields an output count of
    // 1001, one step above the enforced ceiling.
    [SkippableFact]
    public void Execute_OutputFileCountAboveCap_ThrowsArgumentException()
    {
        SkipIfNotWindows();

        var largeInputPath = Path.Combine(TestDir, "large_input_1001.pptx");
        using (var big = new Presentation())
        {
            // Start with 1 default slide; add 1000 more to reach 1001.
            for (var i = 0; i < 1000; i++)
                big.Slides.AddEmptySlide(big.Slides[0].LayoutSlide);
            big.Save(largeInputPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_over_cap");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", largeInputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("1000", ex.Message);

        // Guard must short-circuit before any output file is produced. The
        // output dir may be created by the path-validation step, but it must
        // be empty.
        if (Directory.Exists(outputDir))
            Assert.Empty(Directory.GetFiles(outputDir, "*.pptx"));
    }

    // Negative control for the above cap: when the requested split stays at
    // or below the 1000-file ceiling, the guard must not fire. We narrow the
    // range with endSlideIndex so only a handful of files are produced, which
    // keeps the test fast while still exercising the exact arithmetic used
    // by the guard on a presentation that *could* otherwise exceed the cap.
    [SkippableFact]
    public void Execute_OutputFileCountAtOrBelowCap_DoesNotTriggerGuard()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var largeInputPath = Path.Combine(TestDir, "large_input_1001_ok.pptx");
        using (var big = new Presentation())
        {
            big.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
            for (var i = 0; i < 1000; i++)
                big.Slides.AddEmptySlide(big.Slides[0].LayoutSlide);
            big.Save(largeInputPath, SaveFormat.Pptx);
        }

        var outputDir = Path.Combine(TestDir, "split_under_cap");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        // Narrow to slides [0..4] so outputFileCount = 5 — well under the
        // 1000 cap even though the source has 1001 slides.
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", largeInputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 1 },
            { "startSlideIndex", 0 },
            { "endSlideIndex", 4 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);

        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(5, files.Length);
    }

    #endregion

    #region outputFileNamePattern Validation (bug 20260415-ppt-split-pattern-validation)

    /// <summary>
    ///     Guard for bug 20260415-ppt-split-pattern-validation: a pattern missing the
    ///     literal <c>{index}</c> placeholder would cause every output file to be written
    ///     to the same path, silently clobbering earlier slides. The pre-loop precondition
    ///     must reject such patterns with <see cref="ArgumentException" /> before any file
    ///     is opened or any output directory is populated.
    /// </summary>
    [SkippableFact]
    public void Execute_WithPatternMissingIndexPlaceholder_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "split_pattern_noindex");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "constant.pptx" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("{index}", ex.Message, StringComparison.Ordinal);

        // Guard must short-circuit before any output file is produced.
        if (Directory.Exists(outputDir))
            Assert.Empty(Directory.GetFiles(outputDir, "*.pptx"));
    }

    /// <summary>
    ///     Guard for the ordering-safe length precondition: a raw pattern exceeding 255
    ///     chars where the <c>{index}</c> placeholder falls past the truncation boundary
    ///     must be rejected BEFORE <see cref="AsposeMcpServer.Helpers.SecurityHelper.SanitizeFileNamePattern" />
    ///     silently truncates it — otherwise the sanitizer would strip <c>{index}</c>
    ///     and the split would regress into the silent-overwrite symptom this fix closes.
    /// </summary>
    [SkippableFact]
    public void Execute_WithExcessivelyLongPatternPlaceholderPastTruncation_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "split_pattern_long");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);

        // 300-char prefix pushes {index} well past the 255-char SanitizeFileNamePattern
        // truncation boundary — post-sanitation the placeholder would be gone.
        var longPattern = new string('a', 300) + "_{index}.pptx";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", longPattern }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("length", ex.Message, StringComparison.OrdinalIgnoreCase);

        if (Directory.Exists(outputDir))
            Assert.Empty(Directory.GetFiles(outputDir, "*.pptx"));
    }

    /// <summary>
    ///     Guard for bug 20260415-ppt-split-pattern-validation: a pattern containing a
    ///     parent-directory escape such as <c>../escape_{index}.pptx</c> must never
    ///     produce an output file outside the bounded <c>outputDirectory</c>. The fix's
    ///     post-<c>Path.Combine</c> containment check throws <see cref="ArgumentException" />
    ///     once sanitation has defused the traversal and the resolved path is still
    ///     found to be outside the output root (belt-and-braces).
    /// </summary>
    /// <param name="maliciousPattern">
    ///     An attacker-controlled pattern attempting to escape
    ///     <c>outputDirectory</c> via relative-traversal or absolute-path prefixes.
    /// </param>
    [SkippableTheory]
    [InlineData("../escape_{index}.pptx")]
    [InlineData("..\\escape_{index}.pptx")]
    [InlineData("../../../tmp/escape_{index}.pptx")]
    [InlineData("/tmp/abs_{index}.pptx")]
    [InlineData("\\etc\\abs_{index}.pptx")]
    public void Execute_WithTraversalPattern_WritesOnlyInsideOutputDirectory(string maliciousPattern)
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_pattern_traversal_" + Guid.NewGuid().ToString("N"));
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", maliciousPattern }
        });

        // Pre-check: record neighbour directory state so any sibling-directory write is visible.
        var parent = Directory.GetParent(outputDir)!.FullName;
        var siblingBefore = Directory.Exists(parent)
            ? Directory.GetFiles(parent, "*.pptx", SearchOption.TopDirectoryOnly).ToHashSet()
            : new HashSet<string>();

        try
        {
            _handler.Execute(context, parameters);
        }
        catch (ArgumentException)
        {
            // Acceptable: rejection at the containment guard.
        }

        var siblingAfter = Directory.Exists(parent)
            ? Directory.GetFiles(parent, "*.pptx", SearchOption.TopDirectoryOnly).ToHashSet()
            : new HashSet<string>();
        Assert.Equal(siblingBefore, siblingAfter);

        if (Directory.Exists(outputDir))
        {
            var root = Path.GetFullPath(outputDir) + Path.DirectorySeparatorChar;
            foreach (var file in Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories))
            {
                var full = Path.GetFullPath(file);
                Assert.StartsWith(root, full, StringComparison.OrdinalIgnoreCase);
            }
        }

        Assert.False(File.Exists("/tmp/abs_0.pptx"), "absolute-path attack leaked into /tmp");
    }

    /// <summary>
    ///     Guard for bug 20260415-ppt-split-pattern-validation: an absolute-path pattern
    ///     (e.g. <c>/tmp/abs_{index}.pptx</c>) must be sanitized or rejected so the
    ///     resolved per-slide path never lands outside <c>outputDirectory</c>. Included
    ///     separately from the theory above to make the absolute-path contract visible
    ///     as a named regression guard.
    /// </summary>
    [SkippableFact]
    public void Execute_WithAbsolutePathPattern_DoesNotEscapeOutputDirectory()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_pattern_abs_" + Guid.NewGuid().ToString("N"));
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "/tmp/abs_{index}.pptx" }
        });

        try
        {
            _handler.Execute(context, parameters);
        }
        catch (ArgumentException)
        {
            // Acceptable: rejection at the containment guard.
        }

        Assert.False(File.Exists("/tmp/abs_0.pptx"),
            "absolute-path pattern leaked an output file into /tmp");

        if (Directory.Exists(outputDir))
        {
            var root = Path.GetFullPath(outputDir) + Path.DirectorySeparatorChar;
            foreach (var file in Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories))
            {
                var full = Path.GetFullPath(file);
                Assert.StartsWith(root, full, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    /// <summary>
    ///     Guard for bug 20260415-ppt-split-pattern-validation: a NUL byte in the pattern
    ///     is not a legal filename character on any supported platform. The handler must
    ///     surface this as an <see cref="ArgumentException" /> (either at the pattern
    ///     precondition or at the <c>Path.Combine</c> / file-save call), never silently
    ///     writing truncated filenames.
    /// </summary>
    [SkippableFact]
    public void Execute_WithNulBytePattern_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_pattern_nul");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "bad\0name_{index}.pptx" }
        });

        Assert.ThrowsAny<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
