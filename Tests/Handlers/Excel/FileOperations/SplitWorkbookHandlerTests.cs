using System.Reflection;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Handlers.Excel.FileOperations;

public class SplitWorkbookHandlerTests : ExcelHandlerTestBase
{
    private static readonly int[] SheetIndices = [0, 2];

    private readonly SplitWorkbookHandler _handler = new();

    private string CreateMultiSheetWorkbook()
    {
        var inputPath = Path.Combine(TestDir, $"input_{Guid.NewGuid()}.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells[0, 0].Value = "Data1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells[0, 0].Value = "Data2";
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets[2].Cells[0, 0].Value = "Data3";
        workbook.Save(inputPath);
        return inputPath;
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Split()
    {
        Assert.Equal("split", _handler.Operation);
    }

    #endregion

    #region Extended Coverage v2 - MCP tool annotation

    /// <summary>
    ///     Split is a file-creating operation and must be exposed with
    ///     <c>Destructive=true</c> so MCP clients require user confirmation.
    ///     This is enforced at the <see cref="ExcelFileOperationsTool.Execute" /> entry point
    ///     (shared by create/merge/split).
    /// </summary>
    [Fact]
    public void ExcelFileOperationsTool_Execute_HasDestructiveAnnotation()
    {
        var method = typeof(ExcelFileOperationsTool).GetMethod(
            nameof(ExcelFileOperationsTool.Execute),
            BindingFlags.Public | BindingFlags.Instance);
        Assert.NotNull(method);

        var attr = method.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(attr);
        Assert.True(attr.Destructive, "excel_file_operations must be marked Destructive=true");
        Assert.False(attr.ReadOnly, "excel_file_operations must not be marked ReadOnly");
    }

    #endregion

    #region Basic Split Operations

    [SkippableFact]
    public void Execute_SplitsAllSheets()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "split_output");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3", result.Message);
        Assert.True(Directory.Exists(outputDir));

        var splitFiles = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(3, splitFiles.Length);
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");

            using var splitWorkbook = new Workbook(splitFile);
            Assert.True(splitWorkbook.Worksheets.Count > 0, "Split workbook should have at least one worksheet");
        }
    }

    [Fact]
    public void Execute_WithPath_SplitsAllSheets()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "split_output_path");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));

        var splitFiles = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.True(splitFiles.Length > 0, "Split files should be created");
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");
        }
    }

    [Fact]
    public void Execute_WithSheetIndices_SplitsSpecificSheets()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "split_specific");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "sheetIndices", SheetIndices }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);

        var splitFiles = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, splitFiles.Length);
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");

            using var splitWorkbook = new Workbook(splitFile);
            Assert.True(splitWorkbook.Worksheets.Count > 0, "Split workbook should have at least one worksheet");
        }
    }

    [Fact]
    public void Execute_WithOutputFileNamePattern_UsesPattern()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "split_pattern");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "workbook_{index}.xlsx" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        var patternFile = Path.Combine(outputDir, "workbook_0.xlsx");
        Assert.True(File.Exists(patternFile));
        var fileInfo = new FileInfo(patternFile);
        Assert.True(fileInfo.Length > 0, "Split file should have content");

        using var splitWorkbook = new Workbook(patternFile);
        Assert.True(splitWorkbook.Worksheets.Count > 0, "Split workbook should have at least one worksheet");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSource_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDirectory", Path.Combine(TestDir, "output") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputDirectory_ThrowsArgumentException()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Path Traversal / Validation (bug 20260415-excel-split-path-traversal)

    public static TheoryData<string> MaliciousSourcePaths()
    {
        return
        [
            "../../../etc/passwd",
            "foo/../../secret.xlsx",
            "foo\\..\\..\\secret.xlsx",
            "~/.ssh/id_rsa",
            "good\\\\..\\\\evil.xlsx",
            "file\0name.xlsx" // NUL byte (invalid path char)
        ];
    }

    [Theory]
    [MemberData(nameof(MaliciousSourcePaths))]
    public void Execute_WithMaliciousInputPath_ThrowsArgumentException(string badPath)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", badPath },
            { "outputDirectory", Path.Combine(TestDir, "out_malicious") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [MemberData(nameof(MaliciousSourcePaths))]
    public void Execute_WithMaliciousPathAlias_ThrowsArgumentException(string badPath)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", badPath },
            { "outputDirectory", Path.Combine(TestDir, "out_malicious_alias") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithAbsoluteSourceOutsideAllowedRoot_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", "/etc/passwd/../../../etc/shadow" }, // absolute + traversal
            { "outputDirectory", Path.Combine(TestDir, "out_abs") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Extended Coverage v2 - outputFileNamePattern attacks (bug 20260415-excel-split-path-traversal)

    /// <summary>
    ///     Attack patterns that try to escape <c>outputDirectory</c> via the pattern itself.
    ///     The fix uses <see cref="AsposeMcpServer.Helpers.SecurityHelper.SanitizeFileNamePattern" /> (strips
    ///     <c>..</c>, replaces path separators) followed by a <c>Path.GetFullPath</c> containment
    ///     check. Sanitization defuses the traversal; the containment check is the belt-and-braces
    ///     guarantee. In all cases the resulting file MUST land inside <c>outputDirectory</c>.
    /// </summary>
    [SkippableTheory]
    [InlineData("../../../tmp/escape_{index}.xlsx")]
    [InlineData("..\\..\\..\\tmp\\escape_{index}.xlsx")]
    [InlineData("/tmp/abs_{index}.xlsx")]
    [InlineData("\\etc\\abs_{index}.xlsx")]
    public void Execute_WithTraversalPattern_WritesOnlyInsideOutputDirectory(string maliciousPattern)
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "pattern_traversal_" + Guid.NewGuid().ToString("N"));
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", maliciousPattern }
        });

        // Pre-check: record a set of forbidden neighbour directories we must NOT touch.
        var parent = Directory.GetParent(outputDir)!.FullName;
        var siblingBefore = Directory.Exists(parent)
            ? Directory.GetFiles(parent, "*.xlsx", SearchOption.TopDirectoryOnly).ToHashSet()
            : new HashSet<string>();

        try
        {
            _handler.Execute(context, parameters);
        }
        catch (ArgumentException)
        {
            // Acceptable outcome: rejection at the pattern-containment guard.
        }

        // Invariant: no file was created in any directory OTHER than outputDir.
        var siblingAfter = Directory.Exists(parent)
            ? Directory.GetFiles(parent, "*.xlsx", SearchOption.TopDirectoryOnly).ToHashSet()
            : new HashSet<string>();
        Assert.Equal(siblingBefore, siblingAfter);

        if (Directory.Exists(outputDir))
        {
            var produced = Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories);
            foreach (var file in produced)
            {
                var full = Path.GetFullPath(file);
                var root = Path.GetFullPath(outputDir) + Path.DirectorySeparatorChar;
                Assert.StartsWith(root, full, StringComparison.OrdinalIgnoreCase);
            }
        }

        // Forbidden locations must remain untouched.
        Assert.False(File.Exists("/tmp/abs_0.xlsx"), "absolute-path attack leaked into /tmp");
    }

    [Fact]
    public void Execute_WithNulBytePattern_ThrowsArgumentException()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "pattern_nul");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "bad\0name_{index}.xlsx" }
        });

        // NUL is not a legal file-name character; the runtime (Path.Combine/File.Create)
        // throws ArgumentException well before any data is written.
        Assert.ThrowsAny<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithPatternMissingIndexPlaceholder_DoesNotEscapeOutputDir()
    {
        // Pre-fix MEDIUM finding: a pattern without {index} collapses all output files
        // onto the same path, clobbering data. Whether the fix rejects or tolerates it,
        // the file(s) MUST still be contained inside outputDirectory.
        SkipInEvaluationMode(AsposeLibraryType.Cells);
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "pattern_noindex_" + Guid.NewGuid().ToString("N"));
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "constant.xlsx" }
        });

        try
        {
            _handler.Execute(context, parameters);
        }
        catch (ArgumentException)
        {
            // Acceptable: a stricter fix may reject patterns missing {index}.
            return;
        }

        Assert.True(Directory.Exists(outputDir));
        foreach (var file in Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories))
        {
            var full = Path.GetFullPath(file);
            var root = Path.GetFullPath(outputDir) + Path.DirectorySeparatorChar;
            Assert.StartsWith(root, full, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Execute_WithExcessivelyLongPattern_DoesNotCrashAndRemainsContained()
    {
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "pattern_long_" + Guid.NewGuid().ToString("N"));
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);

        // 10 KB pattern exercises the MaxFileNameLength truncation path in
        // SanitizeFileNamePattern and must NOT cause unbounded allocation or crash.
        var longPattern = new string('a', 10_000) + "_{index}.xlsx";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", longPattern }
        });

        try
        {
            _handler.Execute(context, parameters);
        }
        catch (ArgumentException)
        {
            // Acceptable: a stricter validator may reject overlong patterns outright.
            return;
        }
        catch (PathTooLongException)
        {
            // Also acceptable: OS path-length limit kicks in after truncation.
            return;
        }

        if (Directory.Exists(outputDir))
            foreach (var file in Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories))
            {
                var full = Path.GetFullPath(file);
                var root = Path.GetFullPath(outputDir) + Path.DirectorySeparatorChar;
                Assert.StartsWith(root, full, StringComparison.OrdinalIgnoreCase);
            }
    }

    #endregion

    #region Extended Coverage v2 - allowAbsolutePaths tightening

    /// <summary>
    ///     When <see cref="ServerConfig.AllowedBasePaths" /> is configured, an absolute path
    ///     INSIDE the allowed base must pass.
    /// </summary>
    [SkippableFact]
    public void Execute_AbsolutePath_InsideAllowedBase_Succeeds()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);
        var inputPath = CreateMultiSheetWorkbook();
        var outputDir = Path.Combine(TestDir, "abs_allowed_" + Guid.NewGuid().ToString("N"));

        // TestDir is the allowed base; both inputPath and outputDir live under it.
        var config = ServerConfig.LoadFromArgs(["--excel", "--allowed-path", TestDir]);
        var workbook = CreateEmptyWorkbook();
        var context = new OperationContext<Workbook>
        {
            Document = workbook,
            ServerConfig = config
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);
        Assert.IsType<SuccessResult>(res);
    }

    /// <summary>
    ///     When an allowlist is configured, an absolute path OUTSIDE every allowed base
    ///     MUST be rejected — even when the path itself is syntactically valid and
    ///     contains no <c>..</c> traversal. This is the new vector closed by the extended fix.
    /// </summary>
    [Fact]
    public void Execute_AbsoluteInputPath_OutsideAllowedBase_ThrowsArgumentException()
    {
        var allowedBase = Path.Combine(TestDir, "allowed_root");
        Directory.CreateDirectory(allowedBase);
        var outsideBase = Path.Combine(TestDir, "other_root");
        Directory.CreateDirectory(outsideBase);
        var outsideFile = Path.Combine(outsideBase, "victim.xlsx");
        File.WriteAllBytes(outsideFile, [0]); // doesn't need to be a real xlsx — rejection happens first

        var config = ServerConfig.LoadFromArgs(["--excel", "--allowed-path", allowedBase]);
        var workbook = CreateEmptyWorkbook();
        var context = new OperationContext<Workbook>
        {
            Document = workbook,
            ServerConfig = config
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", outsideFile },
            { "outputDirectory", Path.Combine(allowedBase, "out") }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("allowed", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AbsoluteOutputDirectory_OutsideAllowedBase_ThrowsArgumentException()
    {
        var allowedBase = Path.Combine(TestDir, "allowed_root_out");
        Directory.CreateDirectory(allowedBase);
        var inputPath = Path.Combine(allowedBase, "in.xlsx");
        var srcWb = new Workbook();
        srcWb.Worksheets[0].Cells[0, 0].Value = "x";
        srcWb.Save(inputPath);

        var outsideOutput = Path.Combine(TestDir, "outside_out");

        var config = ServerConfig.LoadFromArgs(["--excel", "--allowed-path", allowedBase]);
        var workbook = CreateEmptyWorkbook();
        var context = new OperationContext<Workbook>
        {
            Document = workbook,
            ServerConfig = config
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputDirectory", outsideOutput }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("allowed", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(Directory.Exists(outsideOutput), "outputDirectory must not be created before allowlist check");
    }

    #endregion
}
