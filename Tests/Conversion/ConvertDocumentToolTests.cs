using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Conversion;

public class ConvertDocumentToolTests : TestBase
{
    private readonly ConvertDocumentTool _tool = new();

    [Fact]
    public async Task ConvertWordToPdf_ShouldConvertDocument()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_word_to_pdf.docx");
        var doc = new Document();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_pdf_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task ConvertExcelToCsv_ShouldConvertDocument()
    {
        // Arrange
        var workbookPath = CreateTestFilePath("test_convert_excel_to_csv.xlsx");
        var workbook = new Workbook();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_csv_output.csv");
        var arguments = new JsonObject
        {
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "CSV file should be created");
    }

    [Fact]
    public async Task ConvertWordToHtml_ShouldConvertToHtml()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_word_to_html.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("HTML Test Content");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_html_output.html");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "HTML file should be created");
        var htmlContent = await File.ReadAllTextAsync(outputPath);
        Assert.Contains("HTML Test Content", htmlContent);
    }

    [Fact]
    public async Task ConvertExcelToHtml_ShouldConvertToHtml()
    {
        // Arrange
        var workbookPath = CreateTestFilePath("test_convert_excel_to_html.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Excel HTML Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_html_output.html");
        var arguments = new JsonObject
        {
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "HTML file should be created from Excel");
    }

    [Fact]
    public async Task ConvertPowerPointToPdf_ShouldConvertToPdf()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_ppt_to_pdf.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_to_pdf_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from PowerPoint");
    }

    [Fact]
    public async Task ConvertPowerPointToHtml_ShouldConvertToHtml()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_ppt_to_html.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_to_html_output.html");
        var arguments = new JsonObject
        {
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "HTML file should be created from PowerPoint");
    }

    [Fact]
    public async Task ConvertWordToRtf_ShouldConvertToRtf()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_word_to_rtf.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("RTF Conversion Test");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_rtf_output.rtf");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "RTF file should be created");
    }

    [Fact]
    public async Task ConvertWordToText_ShouldConvertToText()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_word_to_txt.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain Text Conversion");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_txt_output.txt");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "TXT file should be created");
        var content = await File.ReadAllTextAsync(outputPath);
        Assert.Contains("Plain Text Conversion", content);
    }

    [Fact]
    public async Task Convert_WithUnsupportedInputFormat_ShouldThrowException()
    {
        // Arrange
        var unsupportedPath = CreateTestFilePath("test_unsupported.xyz");
        await File.WriteAllTextAsync(unsupportedPath, "Test content");

        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = unsupportedPath,
            ["outputPath"] = outputPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(async () => await _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ConvertExcelToOds_ShouldConvertToOds()
    {
        // Arrange
        var workbookPath = CreateTestFilePath("test_convert_excel_to_ods.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "ODS Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_ods_output.ods");
        var arguments = new JsonObject
        {
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "ODS file should be created");
    }
}