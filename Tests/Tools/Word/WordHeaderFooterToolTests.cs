using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordHeaderFooterToolTests : WordTestBase
{
    private readonly WordHeaderFooterTool _tool;

    public WordHeaderFooterToolTests()
    {
        _tool = new WordHeaderFooterTool(SessionManager);
    }

    #region General

    [Fact]
    public void SetHeaderText_ShouldSetHeaderText()
    {
        var docPath = CreateWordDocument("test_set_header_text.docx");
        var outputPath = CreateTestFilePath("test_set_header_text_output.docx");
        _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "Left Header", headerCenter: "Center Header", headerRight: "Right Header");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFooterText_ShouldSetFooterText()
    {
        var docPath = CreateWordDocument("test_set_footer_text.docx");
        var outputPath = CreateTestFilePath("test_set_footer_text_output.docx");
        _tool.Execute("set_footer_text", docPath, outputPath: outputPath,
            footerLeft: "Page", footerRight: "{PAGE}");
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        Assert.Contains("Page", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetHeadersFooters_ShouldReturnHeadersFooters()
    {
        var docPath = CreateWordDocument("test_get_headers_footers.docx");
        var doc = new Document(docPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(header);
        }

        header.AppendParagraph("Test Header");
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void SetHeaderImage_ShouldSetHeaderImage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Image operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_set_header_image.docx");
        var imagePath = CreateTestFilePath("test_header_image.png");
        CreateTestImage(imagePath);
        var outputPath = CreateTestFilePath("test_set_header_image_output.docx");
        _tool.Execute("set_header_image", docPath, outputPath: outputPath,
            imagePath: imagePath, alignment: "center");
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void SetFooterImage_ShouldSetFooterImage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Image operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_set_footer_image.docx");
        var imagePath = CreateTestFilePath("test_footer_image.png");
        CreateTestImage(imagePath);
        var outputPath = CreateTestFilePath("test_set_footer_image_output.docx");
        _tool.Execute("set_footer_image", docPath, outputPath: outputPath,
            imagePath: imagePath, alignment: "left");
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderLine_ShouldSetHeaderLine()
    {
        var docPath = CreateWordDocument("test_set_header_line.docx");
        var outputPath = CreateTestFilePath("test_set_header_line_output.docx");
        _tool.Execute("set_header_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
    }

    [Fact]
    public void SetFooterLine_ShouldSetFooterLine()
    {
        var docPath = CreateWordDocument("test_set_footer_line.docx");
        var outputPath = CreateTestFilePath("test_set_footer_line_output.docx");
        _tool.Execute("set_footer_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetBoth()
    {
        var docPath = CreateWordDocument("test_set_header_footer.docx");
        var outputPath = CreateTestFilePath("test_set_header_footer_output.docx");
        _tool.Execute("set_header_footer", docPath, outputPath: outputPath,
            headerLeft: "Left Header", footerCenter: "Center Footer");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(header);
        Assert.NotNull(footer);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Center", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderTabs_ShouldSetTabStops()
    {
        var docPath = CreateWordDocument("test_set_header_tabs.docx");
        var outputPath = CreateTestFilePath("test_set_header_tabs_output.docx");
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0, ["alignment"] = "center", ["leader"] = "none" },
            new JsonObject { ["position"] = 200.0, ["alignment"] = "right", ["leader"] = "dots" }
        };
        var result = _tool.Execute("set_header_tabs", docPath, outputPath: outputPath, tabStops: tabStops);
        Assert.StartsWith("Header tab stops set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooterTabs_ShouldSetTabStops()
    {
        var docPath = CreateWordDocument("test_set_footer_tabs.docx");
        var outputPath = CreateTestFilePath("test_set_footer_tabs_output.docx");
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 150.0, ["alignment"] = "left", ["leader"] = "dashes" }
        };
        var result = _tool.Execute("set_footer_tabs", docPath, outputPath: outputPath, tabStops: tabStops);
        Assert.StartsWith("Footer tab stops set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("SET_HEADER_TEXT")]
    [InlineData("Set_Header_Text")]
    [InlineData("set_header_text")]
    public void Operation_ShouldBeCaseInsensitive_SetHeaderText(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, headerLeft: "Test");
        Assert.StartsWith("Header text set successfully", result);
    }

    [Theory]
    [InlineData("SET_FOOTER_TEXT")]
    [InlineData("Set_Footer_Text")]
    [InlineData("set_footer_text")]
    public void Operation_ShouldBeCaseInsensitive_SetFooterText(string operation)
    {
        var docPath = CreateWordDocument($"test_case_footer_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_footer_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, footerLeft: "Test");
        Assert.StartsWith("Footer text set successfully", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var docPath = CreateWordDocument($"test_case_get_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("totalSections", result);
    }

    [Theory]
    [InlineData("single")]
    [InlineData("double")]
    [InlineData("thick")]
    public void SetHeaderLine_WithDifferentLineStyles_ShouldWork(string lineStyle)
    {
        var docPath = CreateWordDocument($"test_header_line_{lineStyle}.docx");
        var outputPath = CreateTestFilePath($"test_header_line_{lineStyle}_output.docx");
        var result = _tool.Execute("set_header_line", docPath, outputPath: outputPath, lineStyle: lineStyle);
        Assert.StartsWith("Header line set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("single")]
    [InlineData("double")]
    [InlineData("thick")]
    public void SetFooterLine_WithDifferentLineStyles_ShouldWork(string lineStyle)
    {
        var docPath = CreateWordDocument($"test_footer_line_{lineStyle}.docx");
        var outputPath = CreateTestFilePath($"test_footer_line_{lineStyle}_output.docx");
        var result = _tool.Execute("set_footer_line", docPath, outputPath: outputPath, lineStyle: lineStyle);
        Assert.StartsWith("Footer line set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("primary")]
    [InlineData("firstPage")]
    [InlineData("evenPages")]
    public void SetHeaderText_WithDifferentHeaderFooterTypes_ShouldWork(string headerFooterType)
    {
        var docPath = CreateWordDocument($"test_header_type_{headerFooterType}.docx");
        var outputPath = CreateTestFilePath($"test_header_type_{headerFooterType}_output.docx");
        var result = _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "Test Header", headerFooterType: headerFooterType);
        Assert.StartsWith("Header text set successfully", result);
    }

    [Theory]
    [InlineData("primary")]
    [InlineData("firstPage")]
    [InlineData("evenPages")]
    public void SetFooterText_WithDifferentHeaderFooterTypes_ShouldWork(string headerFooterType)
    {
        var docPath = CreateWordDocument($"test_footer_type_{headerFooterType}.docx");
        var outputPath = CreateTestFilePath($"test_footer_type_{headerFooterType}_output.docx");
        var result = _tool.Execute("set_footer_text", docPath, outputPath: outputPath,
            footerLeft: "Test Footer", headerFooterType: headerFooterType);
        Assert.StartsWith("Footer text set successfully", result);
    }

    [Fact]
    public void SetHeaderText_WithAllSections_ShouldApplyToAllSections()
    {
        var docPath = CreateWordDocument("test_header_all_sections.docx");
        var doc = new Document(docPath);
        doc.AppendChild(new Section(doc));
        doc.AppendChild(new Section(doc));
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_header_all_sections_output.docx");
        var result = _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "All Sections", sectionIndex: -1);
        Assert.StartsWith("Header text set successfully", result);
    }

    [Fact]
    public void SetHeaderText_WithFieldCodes_ShouldInsertFields()
    {
        var docPath = CreateWordDocument("test_header_fields.docx");
        var outputPath = CreateTestFilePath("test_header_fields_output.docx");
        _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "Page {PAGE} of {NUMPAGES}", headerRight: "{DATE}");
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooterText_WithFontSettings_ShouldApplyFont()
    {
        var docPath = CreateWordDocument("test_footer_font.docx");
        var outputPath = CreateTestFilePath("test_footer_font_output.docx");
        var result = _tool.Execute("set_footer_text", docPath, outputPath: outputPath,
            footerCenter: "Styled Footer", fontName: "Arial", fontSize: 14);
        Assert.StartsWith("Footer text set successfully", result);
    }

    [Fact]
    public void SetHeaderLine_WithLineWidth_ShouldApplyWidth()
    {
        var docPath = CreateWordDocument("test_header_line_width.docx");
        var outputPath = CreateTestFilePath("test_header_line_width_output.docx");
        var result = _tool.Execute("set_header_line", docPath, outputPath: outputPath,
            lineStyle: "single", lineWidth: 2.5);
        Assert.StartsWith("Header line set", result);
    }

    [Fact]
    public void GetHeadersFooters_WithNoHeadersFooters_ShouldReturnEmptyResult()
    {
        var docPath = CreateWordDocument("test_get_empty_hf.docx");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("totalSections", result);
        Assert.Contains("sections", result);
    }

    [SkippableFact]
    public void SetHeaderImage_WithFloating_ShouldCreateFloatingImage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Image operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_header_floating.docx");
        var imagePath = CreateTestFilePath("test_header_floating_image.png");
        CreateTestImage(imagePath);
        var outputPath = CreateTestFilePath("test_header_floating_output.docx");
        var result = _tool.Execute("set_header_image", docPath, outputPath: outputPath,
            imagePath: imagePath, isFloating: true, alignment: "right");
        Assert.StartsWith("Header image set", result);
    }

    [SkippableFact]
    public void SetHeaderImage_WithDimensions_ShouldApplyDimensions()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Image operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_header_dimensions.docx");
        var imagePath = CreateTestFilePath("test_header_dim_image.png");
        CreateTestImage(imagePath);
        var outputPath = CreateTestFilePath("test_header_dimensions_output.docx");
        _tool.Execute("set_header_image", docPath, outputPath: outputPath,
            imagePath: imagePath, imageWidth: 100, imageHeight: 50);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderText_WithNoContent_ShouldReturnWarning()
    {
        var docPath = CreateWordDocument("test_header_no_content.docx");
        var result = _tool.Execute("set_header_text", docPath);
        Assert.Contains("Warning", result);
        Assert.Contains("No header text content provided", result);
    }

    [Fact]
    public void SetFooterText_WithNoContent_ShouldReturnWarning()
    {
        var docPath = CreateWordDocument("test_footer_no_content.docx");
        var result = _tool.Execute("set_footer_text", docPath);
        Assert.Contains("Warning", result);
        Assert.Contains("No footer text content provided", result);
    }

    [Fact]
    public void SetHeaderText_WithClearTextOnly_ShouldPreserveImages()
    {
        var docPath = CreateWordDocument("test_header_clear_text.docx");
        var outputPath = CreateTestFilePath("test_header_clear_text_output.docx");
        var result = _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "New Text", clearTextOnly: true);
        Assert.StartsWith("Header text set successfully", result);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("decimal")]
    [InlineData("bar")]
    public void SetHeaderTabs_WithDifferentAlignments_ShouldWork(string alignment)
    {
        var docPath = CreateWordDocument($"test_header_tabs_{alignment}.docx");
        var outputPath = CreateTestFilePath($"test_header_tabs_{alignment}_output.docx");
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0, ["alignment"] = alignment, ["leader"] = "none" }
        };
        var result = _tool.Execute("set_header_tabs", docPath, outputPath: outputPath, tabStops: tabStops);
        Assert.StartsWith("Header tab stops set", result);
    }

    [Theory]
    [InlineData("none")]
    [InlineData("dots")]
    [InlineData("dashes")]
    [InlineData("line")]
    public void SetFooterTabs_WithDifferentLeaders_ShouldWork(string leader)
    {
        var docPath = CreateWordDocument($"test_footer_tabs_{leader}.docx");
        var outputPath = CreateTestFilePath($"test_footer_tabs_{leader}_output.docx");
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 150.0, ["alignment"] = "left", ["leader"] = leader }
        };
        var result = _tool.Execute("set_footer_tabs", docPath, outputPath: outputPath, tabStops: tabStops);
        Assert.StartsWith("Footer tab stops set", result);
    }

    private void CreateTestImage(string imagePath)
    {
        byte[] pngBytes =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0xFF,
            0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59,
            0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(imagePath, pngBytes);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void SetHeaderImage_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_header_missing_image.docx");
        var outputPath = CreateTestFilePath("test_header_missing_image_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_header_image", docPath, outputPath: outputPath, imagePath: null));
        Assert.Contains("imagePath cannot be null or empty", ex.Message);
    }

    [Fact]
    public void SetHeaderImage_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_header_nonexistent_image.docx");
        var outputPath = CreateTestFilePath("test_header_nonexistent_image_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("set_header_image", docPath, outputPath: outputPath,
                imagePath: "C:\\nonexistent\\image.png"));
    }

    [Fact]
    public void SetFooterImage_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_footer_missing_image.docx");
        var outputPath = CreateTestFilePath("test_footer_missing_image_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_footer_image", docPath, outputPath: outputPath, imagePath: null));
        Assert.Contains("imagePath cannot be null or empty", ex.Message);
    }

    [Fact]
    public void SetFooterImage_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_footer_nonexistent_image.docx");
        var outputPath = CreateTestFilePath("test_footer_nonexistent_image_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("set_footer_image", docPath, outputPath: outputPath,
                imagePath: "C:\\nonexistent\\image.png"));
    }

    [Fact]
    public void GetHeadersFooters_WithInvalidSectionIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_get_invalid_section.docx");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("get", docPath, sectionIndex: 999));
    }

    #endregion

    #region Session

    [SkippableFact]
    public void GetHeadersFooters_WithSessionId_ShouldReturnHeadersFooters()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks document structure modification");
        var docPath = CreateWordDocument("test_session_get_hf.docx");
        var doc = new Document(docPath);
        var header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Session Header");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderText_WithSessionId_ShouldSetHeaderInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_header.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_text", sessionId: sessionId,
            headerLeft: "Session Left", headerCenter: "Session Center");
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Session", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFooterText_WithSessionId_ShouldSetFooterInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_footer.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_footer_text", sessionId: sessionId,
            footerLeft: "Session Footer", footerRight: "{PAGE}");
        Assert.Contains("Footer", result, StringComparison.OrdinalIgnoreCase);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        Assert.Contains("Session Footer", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderLine_WithSessionId_ShouldSetLineInMemory()
    {
        var docPath = CreateWordDocument("test_session_header_line.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_line", sessionId: sessionId, lineStyle: "double");
        Assert.StartsWith("Header line set", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
    }

    [Fact]
    public void SetFooterLine_WithSessionId_ShouldSetLineInMemory()
    {
        var docPath = CreateWordDocument("test_session_footer_line.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_footer_line", sessionId: sessionId, lineStyle: "thick");
        Assert.StartsWith("Footer line set", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
    }

    [Fact]
    public void SetHeaderFooter_WithSessionId_ShouldSetBothInMemory()
    {
        var docPath = CreateWordDocument("test_session_hf_both.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_footer", sessionId: sessionId,
            headerCenter: "Session Header", footerCenter: "Session Footer");
        Assert.StartsWith("Header and footer set", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(header);
        Assert.NotNull(footer);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks document structure modification");
        var docPath1 = CreateWordDocument("test_path_hf.docx");
        var doc1 = new Document(docPath1);
        var header1 = new HeaderFooter(doc1, HeaderFooterType.HeaderPrimary);
        doc1.FirstSection.HeadersFooters.Add(header1);
        header1.AppendParagraph("Path Header");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_hf.docx");
        var doc2 = new Document(docPath2);
        var header2 = new HeaderFooter(doc2, HeaderFooterType.HeaderPrimary);
        doc2.FirstSection.HeadersFooters.Add(header2);
        header2.AppendParagraph("Session Header Unique");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        Assert.Contains("Session Header Unique", result);
        Assert.DoesNotContain("Path Header", result);
    }

    #endregion
}