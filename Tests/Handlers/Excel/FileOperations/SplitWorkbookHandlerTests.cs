using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("3", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("2", result);

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
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
}
