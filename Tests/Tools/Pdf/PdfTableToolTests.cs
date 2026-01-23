using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfTableTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfTableToolTests : PdfTestBase
{
    private readonly PdfTableTool _tool;

    public PdfTableToolTests()
    {
        _tool = new PdfTableTool(SessionManager);
    }

    private string CreatePdfDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Sample PDF Text"));
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddTableToPage()
    {
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 3, columns: 3);

        Assert.True(File.Exists(outputPath));
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added table (3 rows x 3 columns) to page 1", data.Message);

        using var outputDoc = new Document(outputPath);
        var tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(outputDoc.Pages[1]);
        Assert.True(tableAbsorber.TableList.Count > 0, "Table should be present in the PDF");
    }

    [Fact]
    public void Add_WithData_ShouldFillTableWithData()
    {
        var pdfPath = CreatePdfDocument("test_add_data.pdf");
        var outputPath = CreateTestFilePath("test_add_data_output.pdf");
        var tableData = new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } };

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2, data: tableData);

        Assert.True(File.Exists(outputPath));
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, rows: 3, columns: 3);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added table (3 rows x 3 columns) to page 1", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);

        var page = document.Pages[1];
        var tableFound = page.Paragraphs?.OfType<Table>().Any() ?? false;
        Assert.True(tableFound, "Table should be present in memory");
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add", sessionId: "invalid_session", pageIndex: 1, rows: 2, columns: 2));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreatePdfDocument("test_path_table.pdf");
        var pdfPath2 = CreatePdfDocument("test_session_table.pdf");
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("add", pdfPath1, sessionId,
            pageIndex: 1, rows: 2, columns: 2);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    #endregion
}
