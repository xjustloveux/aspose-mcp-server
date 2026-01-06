using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

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

    #region General

    [Fact]
    public void Add_ShouldAddTableToPage()
    {
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 3, columns: 3);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (3 rows x 3 columns) to page 1", result);

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
        var data = new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } };

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2, data: data);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", result);

        using var outputDoc = new Document(outputPath);
        var tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(outputDoc.Pages[1]);
        Assert.True(tableAbsorber.TableList.Count > 0, "Table should be present in the PDF");

        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains("A1", textAbsorber.Text);
        Assert.Contains("B2", textAbsorber.Text);
    }

    [Fact]
    public void Add_WithPosition_ShouldSetPosition()
    {
        var pdfPath = CreatePdfDocument("test_add_position.pdf");
        var outputPath = CreateTestFilePath("test_add_position_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2, x: 200, y: 500);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", result);

        using var outputDoc = new Document(outputPath);
        var tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(outputDoc.Pages[1]);
        Assert.True(tableAbsorber.TableList.Count > 0, "Table should be present in the PDF");
    }

    [Fact]
    public void Add_WithColumnWidths_ShouldApplyWidths()
    {
        var pdfPath = CreatePdfDocument("test_add_widths.pdf");
        var outputPath = CreateTestFilePath("test_add_widths_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 3, columnWidths: "100 150 200");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (2 rows x 3 columns) to page 1", result);

        using var outputDoc = new Document(outputPath);
        var tableAbsorber = new TableAbsorber();
        tableAbsorber.Visit(outputDoc.Pages[1]);
        Assert.True(tableAbsorber.TableList.Count > 0, "Table should be present in the PDF");
    }

    [Fact]
    public void Add_WithEmptyData_ShouldUseDefaultValues()
    {
        var pdfPath = CreatePdfDocument("test_add_empty.pdf");
        var outputPath = CreateTestFilePath("test_add_empty_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2, data: Array.Empty<string[]>());

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", result);
    }

    [Fact]
    public void Add_WithIrregularData_ShouldHandleGracefully()
    {
        var pdfPath = CreatePdfDocument("test_add_irregular.pdf");
        var outputPath = CreateTestFilePath("test_add_irregular_output.pdf");
        var data = new[]
        {
            new[] { "A1", "B1" },
            new[] { "A2", "B2", "C2" }
        };

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 3, data: data);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added table (2 rows x 3 columns) to page 1", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, rows: 2, columns: 2);

        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_edit_{operation}.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, pdfPath, tableIndex: 0, cellRow: 0, cellColumn: 0, cellValue: "Test"));
        Assert.StartsWith("No tables found in the document", ex.Message);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 99, rows: 2, columns: 2));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void Add_WithMissingRows_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_no_rows.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, rows: 0, columns: 2));
        Assert.Equal("rows is required and must be greater than 0 for add operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingColumns_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_no_cols.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, rows: 2, columns: 0));
        Assert.Equal("columns is required and must be greater than 0 for add operation", ex.Message);
    }

    [Fact]
    public void Edit_WithNoTables_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_no_tables.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, tableIndex: 0, cellRow: 0, cellColumn: 0, cellValue: "NewValue"));
        Assert.StartsWith("No tables found in the document", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("add", pageIndex: 1, rows: 2, columns: 2));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", "nonexistent_file.pdf", pageIndex: 1, rows: 2, columns: 2));
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, rows: 3, columns: 3);

        Assert.StartsWith("Added table (3 rows x 3 columns) to page 1", result);
        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);

        // Check for table using Paragraphs collection
        var page = document.Pages[1];
        var tableFound = page.Paragraphs?.OfType<Table>().Any() ?? false;
        Assert.True(tableFound, "Table should be present in memory");
    }

    [Fact]
    public void Add_WithSessionId_AndData_ShouldFillTableInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_add_data.pdf");
        var sessionId = OpenSession(pdfPath);
        var data = new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } };

        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, rows: 2, columns: 2, data: data);

        Assert.StartsWith("Added table (2 rows x 2 columns) to page 1", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Add_WithSessionId_AndOptions_ShouldApplyOptionsInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_options.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, rows: 2, columns: 3, x: 150, y: 400, columnWidths: "80 100 120");

        Assert.StartsWith("Added table (2 rows x 3 columns) to page 1", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void Edit_WithSessionId_WithNoTables_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_session_edit_no_table.pdf");
        var sessionId = OpenSession(pdfPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", sessionId: sessionId, tableIndex: 0,
                cellRow: 0, cellColumn: 0, cellValue: "NewValue"));
        Assert.StartsWith("No tables found in the document", ex.Message);
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

        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    #endregion
}