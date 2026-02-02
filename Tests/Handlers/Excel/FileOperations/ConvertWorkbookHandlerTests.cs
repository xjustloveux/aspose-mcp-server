using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.FileOperations;

public class ConvertWorkbookHandlerTests : ExcelHandlerTestBase
{
    private readonly ConvertWorkbookHandler _handler = new();

    private string CreateInputWorkbook()
    {
        var inputPath = Path.Combine(TestDir, $"input_{Guid.NewGuid()}.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells[0, 0].Value = "Test Data";
        workbook.Save(inputPath);
        return inputPath;
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
    }

    #endregion

    #region Basic Convert Operations

    [Fact]
    public void Execute_ConvertsToPdf()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.pdf");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "pdf" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted file should have content");
    }

    [SkippableFact]
    public void Execute_ConvertsToHtml()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.html");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "html" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted file should have content");
        var content = File.ReadAllText(outputPath);
        Assert.Contains("Test Data", content);
    }

    [Fact]
    public void Execute_ConvertsToCsv()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.csv");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "csv" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted file should have content");
        var content = File.ReadAllText(outputPath);
        Assert.Contains("Test Data", content);
    }

    [Fact]
    public void Execute_ConvertsToXls()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.xls");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "xls" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted file should have content");
        using var convertedWorkbook = new Workbook(outputPath);
        Assert.Equal("Test Data", convertedWorkbook.Worksheets[0].Cells[0, 0].StringValue);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutInputPathOrSessionId_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.pdf") },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateInputWorkbook();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutFormat_ThrowsArgumentException()
    {
        var inputPath = CreateInputWorkbook();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.pdf") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var inputPath = CreateInputWorkbook();
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "format", "xyz" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Additional Format Tests

    [Fact]
    public void Execute_ConvertsToOds()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.ods");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "ods" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted file should have content");
    }

    [Fact]
    public void Execute_ConvertsToTxt()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.txt");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "txt" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_ConvertsToTsv()
    {
        var inputPath = CreateInputWorkbook();
        var outputPath = Path.Combine(TestDir, "output.tsv");
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", inputPath },
            { "outputPath", outputPath },
            { "format", "tsv" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    #endregion
}
