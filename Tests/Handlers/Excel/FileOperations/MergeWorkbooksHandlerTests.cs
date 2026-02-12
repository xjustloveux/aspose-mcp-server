using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.FileOperations;

/// <summary>
///     Represents paths to test input workbooks.
/// </summary>
/// <param name="Path1">The first workbook path.</param>
/// <param name="Path2">The second workbook path.</param>
internal record WorkbookPaths(string Path1, string Path2);

public class MergeWorkbooksHandlerTests : ExcelHandlerTestBase
{
    private readonly MergeWorkbooksHandler _handler = new();

    private WorkbookPaths CreateInputWorkbooks()
    {
        var input1Path = Path.Combine(TestDir, $"input1_{Guid.NewGuid()}.xlsx");
        var workbook1 = new Workbook();
        workbook1.Worksheets[0].Name = "Sheet1";
        workbook1.Worksheets[0].Cells[0, 0].Value = "Data1";
        workbook1.Save(input1Path);

        var input2Path = Path.Combine(TestDir, $"input2_{Guid.NewGuid()}.xlsx");
        var workbook2 = new Workbook();
        workbook2.Worksheets[0].Name = "Sheet2";
        workbook2.Worksheets[0].Cells[0, 0].Value = "Data2";
        workbook2.Save(input2Path);

        return new WorkbookPaths(input1Path, input2Path);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Merge()
    {
        Assert.Equal("merge", _handler.Operation);
    }

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesWorkbooks()
    {
        var inputPaths = CreateInputWorkbooks();
        var outputPath = Path.Combine(TestDir, "merged.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath },
            { "inputPaths", new[] { inputPaths.Path1, inputPaths.Path2 } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        using var mergedWorkbook = new Workbook(outputPath);
        Assert.True(mergedWorkbook.Worksheets.Count >= 2,
            "Merged workbook should contain sheets from both source files");
        var allText = string.Join(" ", mergedWorkbook.Worksheets.Select(ws => ws.Cells[0, 0].StringValue));
        Assert.Contains("Data1", allText);
        Assert.Contains("Data2", allText);
    }

    [Fact]
    public void Execute_WithOutputPath_MergesWorkbooks()
    {
        var inputPaths = CreateInputWorkbooks();
        var outputPath = Path.Combine(TestDir, "merged_output.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "inputPaths", new[] { inputPaths.Path1, inputPaths.Path2 } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        using var mergedWorkbook = new Workbook(outputPath);
        Assert.True(mergedWorkbook.Worksheets.Count > 0, "Merged workbook should have worksheets");
    }

    [Fact]
    public void Execute_WithMergeSheets_MergesSheetsWithSameName()
    {
        var inputPaths = CreateInputWorkbooks();

        var input3Path = Path.Combine(TestDir, $"input3_{Guid.NewGuid()}.xlsx");
        var workbook3 = new Workbook();
        workbook3.Worksheets[0].Name = "Sheet1";
        workbook3.Worksheets[0].Cells[0, 0].Value = "Data3";
        workbook3.Save(input3Path);

        var outputPath = Path.Combine(TestDir, "merged_sheets.xlsx");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath },
            { "inputPaths", new[] { inputPaths.Path1, input3Path } },
            { "mergeSheets", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        using var mergedWorkbook = new Workbook(outputPath);
        Assert.True(mergedWorkbook.Worksheets.Count > 0, "Merged workbook should have worksheets");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrOutputPath_ThrowsArgumentException()
    {
        var inputPaths = CreateInputWorkbooks();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { inputPaths.Path1, inputPaths.Path2 } }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutInputPaths_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "merged.xlsx") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyInputPaths_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "merged.xlsx") },
            { "inputPaths", Array.Empty<string>() }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
