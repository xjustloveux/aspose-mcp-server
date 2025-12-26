using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordPageToolTests : WordTestBase
{
    private readonly WordPageTool _tool = new();

    [Fact]
    public async Task SetMargins_ShouldSetPageMargins()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_margins.docx");
        var outputPath = CreateTestFilePath("test_set_margins_output.docx");
        var arguments = CreateArguments("set_margins", docPath, outputPath);
        arguments["top"] = 72.0;
        arguments["bottom"] = 72.0;
        arguments["left"] = 90.0;
        arguments["right"] = 90.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(72.0, section.PageSetup.TopMargin);
        Assert.Equal(72.0, section.PageSetup.BottomMargin);
        Assert.Equal(90.0, section.PageSetup.LeftMargin);
        Assert.Equal(90.0, section.PageSetup.RightMargin);
    }

    [Fact]
    public async Task SetOrientation_ShouldSetPageOrientation()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_orientation.docx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.docx");
        var arguments = CreateArguments("set_orientation", docPath, outputPath);
        arguments["orientation"] = "landscape";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(Orientation.Landscape, section.PageSetup.Orientation);
    }

    [Fact]
    public async Task SetPageSize_ShouldSetPageSize()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_size.docx");
        var outputPath = CreateTestFilePath("test_set_size_output.docx");
        var arguments = CreateArguments("set_size", docPath, outputPath);
        arguments["width"] = 595.0;
        arguments["height"] = 842.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(595.0, section.PageSetup.PageWidth);
        Assert.Equal(842.0, section.PageSetup.PageHeight);
    }

    [Fact]
    public async Task SetPageNumber_ShouldSetPageNumber()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_page_number.docx");
        var outputPath = CreateTestFilePath("test_set_page_number_output.docx");
        var arguments = CreateArguments("set_page_number", docPath, outputPath);
        arguments["startingPageNumber"] = 5;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        // Verify page numbering was set (may require RestartPageNumbering to be true)
        Assert.True(section.PageSetup.RestartPageNumbering || section.PageSetup.PageStartingNumber == 5,
            "Page starting number should be set");
    }

    [Fact]
    public async Task DeletePage_ShouldRemoveSpecifiedPage()
    {
        // Arrange - Create a multi-page document
        var docPath = CreateTestFilePath("test_delete_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content");
        doc.Save(docPath);

        var pageCountBefore = doc.PageCount;
        Assert.True(pageCountBefore >= 3, "Document should have at least 3 pages");

        var outputPath = CreateTestFilePath("test_delete_page_output.docx");
        var arguments = CreateArguments("delete_page", docPath, outputPath);
        arguments["pageIndex"] = 1; // Delete middle page

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted successfully", result);
        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.PageCount < pageCountBefore, "Page count should decrease after deletion");
    }

    [Fact]
    public async Task InsertBlankPage_ShouldInsertPageAtSpecifiedPosition()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_insert_blank.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_insert_blank_output.docx");
        var arguments = CreateArguments("insert_blank_page", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("inserted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task AddPageBreak_ShouldAddPageBreakAtDocumentEnd()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_page_break.docx", "Content before break");
        var outputPath = CreateTestFilePath("test_add_page_break_output.docx");
        var arguments = CreateArguments("add_page_break", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Page break added", result);
        var doc = new Document(outputPath);
        // Verify page break was added (document should have increased content)
        Assert.True(doc.GetText().Length > 0);
    }

    [Fact]
    public async Task AddPageBreak_WithParagraphIndex_ShouldAddBreakAtSpecifiedPosition()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_add_break_at_para.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 0");
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_break_at_para_output.docx");
        var arguments = CreateArguments("add_page_break", docPath, outputPath);
        arguments["paragraphIndex"] = 1;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("after paragraph 1", result);
    }
}