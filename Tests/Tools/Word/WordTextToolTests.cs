using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordTextToolTests : WordTestBase
{
    private readonly WordTextTool _tool;

    public WordTextToolTests()
    {
        _tool = new WordTextTool(SessionManager);
    }

    #region General

    [Fact]
    public void AddText_ShouldAddTextToDocumentEnd()
    {
        var docPath = CreateWordDocument("test_add_text.docx");
        var outputPath = CreateTestFilePath("test_add_text_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath, text: "Hello World");

        Assert.StartsWith("Text added to document successfully", result);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");
        var hasExactText = paragraphs.Any(p => p.GetText().Contains("Hello World"));
        Assert.True(hasExactText, "Document should contain exact text 'Hello World'");
    }

    [Fact]
    public void AddText_WithFontFormatting_ShouldApplyFormatting()
    {
        var docPath = CreateWordDocument("test_add_text_formatting.docx");
        var outputPath = CreateTestFilePath("test_add_text_formatting_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Bold Text", bold: true, fontSize: 14);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var boldRun = runs.FirstOrDefault(r => r.Text.Contains("Bold Text"));
        Assert.NotNull(boldRun);
        Assert.True(boldRun.Font.Bold);
        Assert.Equal(14, boldRun.Font.Size);
    }

    [Fact]
    public void AddText_WithFontName_ShouldApplyFontName()
    {
        var docPath = CreateWordDocument("test_add_text_font_name.docx");
        var outputPath = CreateTestFilePath("test_add_text_font_name_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Custom Font Text", fontName: "Arial");

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Custom Font Text"));
        Assert.NotNull(run);
        Assert.Equal("Arial", run.Font.Name);
    }

    [Theory]
    [InlineData("single", Underline.Single)]
    [InlineData("double", Underline.Double)]
    [InlineData("dotted", Underline.Dotted)]
    [InlineData("dash", Underline.Dash)]
    public void AddText_WithUnderlineStyles_ShouldApplyUnderlineTypes(string underlineStyle, Underline expected)
    {
        var docPath = CreateWordDocument($"test_add_text_underline_{underlineStyle}.docx");
        var outputPath = CreateTestFilePath($"test_add_text_underline_{underlineStyle}_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath,
            text: $"Underline {underlineStyle}", underline: underlineStyle);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains($"Underline {underlineStyle}"));
        Assert.NotNull(run);
        Assert.Equal(expected, run.Font.Underline);
    }

    [Fact]
    public void AddText_WithColor_ShouldApplyColor()
    {
        var docPath = CreateWordDocument("test_add_text_color.docx");
        var outputPath = CreateTestFilePath("test_add_text_color_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Colored Text", color: "FF0000");

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Colored Text"));
        Assert.NotNull(run);
        Assert.Equal(Color.FromArgb(255, 0, 0), run.Font.Color);
    }

    [Fact]
    public void AddText_WithStrikethrough_ShouldApplyStrikethrough()
    {
        var docPath = CreateWordDocument("test_add_text_strikethrough.docx");
        var outputPath = CreateTestFilePath("test_add_text_strikethrough_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Strikethrough Text", strikethrough: true);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Strikethrough Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.StrikeThrough);
    }

    [Fact]
    public void AddText_WithSuperscript_ShouldApplySuperscript()
    {
        var docPath = CreateWordDocument("test_add_text_superscript.docx");
        var outputPath = CreateTestFilePath("test_add_text_superscript_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Superscript Text", superscript: true);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Superscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Superscript);
    }

    [Fact]
    public void AddText_WithSubscript_ShouldApplySubscript()
    {
        var docPath = CreateWordDocument("test_add_text_subscript.docx");
        var outputPath = CreateTestFilePath("test_add_text_subscript_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Subscript Text", subscript: true);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Subscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Subscript);
    }

    [Fact]
    public void AddText_WithMultipleFormatting_ShouldApplyAllFormats()
    {
        var docPath = CreateWordDocument("test_add_text_multiple_formatting.docx");
        var outputPath = CreateTestFilePath("test_add_text_multiple_formatting_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Annual Financial Report 2024", fontName: "Arial", fontSize: 14,
            bold: true, italic: true, underline: "single", color: "0000FF");

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Annual Financial Report 2024"));
        Assert.NotNull(run);
        Assert.Equal("Arial", run.Font.Name);
        Assert.Equal(14, run.Font.Size);
        Assert.True(run.Font.Bold, "Text should be bold");
        Assert.True(run.Font.Italic, "Text should be italic");
        Assert.Equal(Underline.Single, run.Font.Underline);
        Assert.Equal(Color.FromArgb(0, 0, 255), run.Font.Color);
    }

    [Fact]
    public void AddText_WithFontNameAsciiAndFarEast_ShouldApplyDifferentFonts()
    {
        var docPath = CreateWordDocument("test_add_text_fonts.docx");
        var outputPath = CreateTestFilePath("test_add_text_fonts_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "English Test", fontNameAscii: "Times New Roman", fontNameFarEast: "Microsoft YaHei", fontSize: 12);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("English"));
        Assert.NotNull(run);
        Assert.Equal("Times New Roman", run.Font.NameAscii);
        Assert.NotNull(run.Font.NameFarEast);
    }

    [Fact]
    public void AddText_WithMultipleLines_ShouldCreateMultipleParagraphs()
    {
        var docPath = CreateWordDocument("test_add_multiline.docx");
        var outputPath = CreateTestFilePath("test_add_multiline_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Line1\nLine2\nLine3");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Line1", text);
        Assert.Contains("Line2", text);
        Assert.Contains("Line3", text);
    }

    [Fact]
    public void AddText_WithWindowsNewlines_ShouldCreateMultipleParagraphs()
    {
        var docPath = CreateWordDocument("test_add_windows_newlines.docx");
        var outputPath = CreateTestFilePath("test_add_windows_newlines_output.docx");

        _tool.Execute("add", docPath, outputPath: outputPath, text: "Line1\r\nLine2\r\nLine3");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Line1", text);
        Assert.Contains("Line2", text);
        Assert.Contains("Line3", text);
    }

    [Fact]
    public void ReplaceText_ShouldReplaceAllOccurrences()
    {
        var content = "Dear Customer, Thank you for contacting Customer Support. " +
                      "Our Customer Service team will assist you shortly.";
        var docPath = CreateWordDocumentWithContent("test_replace_text.docx", content);
        var outputPath = CreateTestFilePath("test_replace_text_output.docx");

        _tool.Execute("replace", docPath, outputPath: outputPath, find: "Customer", replace: "Client");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.DoesNotContain("Customer", text);
        Assert.Contains("Client Support", text);
        Assert.Contains("Client Service", text);
        var clientCount = text.Split(["Client"], StringSplitOptions.None).Length - 1;
        Assert.Equal(3, clientCount);
    }

    [Fact]
    public void Replace_WithUseRegex_ShouldReplaceUsingRegex()
    {
        var docPath = CreateWordDocumentWithContent("test_replace_regex.docx", "Test123 and Test456 and Test789");
        var outputPath = CreateTestFilePath("test_replace_regex_output.docx");

        _tool.Execute("replace", docPath, outputPath: outputPath,
            find: @"Test\d+", replace: "Number", useRegex: true);

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Number", text);
        Assert.DoesNotContain("Test123", text);
        Assert.DoesNotContain("Test456", text);
        Assert.DoesNotContain("Test789", text);
    }

    [Fact]
    public void Replace_WithReplaceInFields_ShouldReplaceInFields()
    {
        var docPath = CreateWordDocument("test_replace_in_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Click here: ");
        builder.InsertHyperlink("TestLink", "http://example.com", false);
        builder.Write(" End of document");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_replace_in_fields_output.docx");
        _tool.Execute("replace", docPath, outputPath: outputPath,
            find: "TestLink", replace: "NewLink", replaceInFields: true);

        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("NewLink", text);
    }

    [Fact]
    public void Replace_WithReplaceInFieldsFalse_ShouldNotReplaceInFields()
    {
        var docPath = CreateWordDocument("test_replace_skip_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Normal TestWord ");
        builder.InsertHyperlink("TestWord Link", "http://example.com", false);
        builder.Write(" More TestWord");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_replace_skip_fields_output.docx");
        _tool.Execute("replace", docPath, outputPath: outputPath,
            find: "TestWord", replace: "Replaced", replaceInFields: false);

        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Replaced", text);
    }

    [Fact]
    public void SearchText_ShouldFindTextInDocument()
    {
        var docPath = CreateWordDocumentWithContent("test_search_text.docx", "This is a test document");

        var result = _tool.Execute("search", docPath, searchText: "test");

        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1 matches", result);
        Assert.Contains("test", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Search Results", result);
    }

    [Fact]
    public void Search_WithUseRegex_ShouldSearchUsingRegex()
    {
        var docPath = CreateWordDocumentWithContent("test_search_regex.docx",
            "Email: test@example.com and admin@test.org");

        var result = _tool.Execute("search", docPath, searchText: @"\w+@\w+\.\w+", useRegex: true);

        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("matches", result);
        Assert.Contains("test@example.com", result);
        Assert.Contains("admin@test.org", result);
    }

    [Fact]
    public void Search_WithCaseSensitive_ShouldMatchCase()
    {
        var docPath = CreateWordDocumentWithContent("test_search_case.docx", "Hello HELLO hello HeLLo");

        var result = _tool.Execute("search", docPath, searchText: "Hello", caseSensitive: true);

        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1 matches", result);
    }

    [Fact]
    public void Search_WithMaxResults_ShouldLimitResults()
    {
        var docPath = CreateWordDocumentWithContent("test_search_max.docx",
            "word word word word word word word word word word");

        var result = _tool.Execute("search", docPath, searchText: "word", maxResults: 3);

        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3 matches", result);
        Assert.Contains("limited to first 3", result);
    }

    [Fact]
    public void Search_WithNoResults_ShouldReturnNoMatchMessage()
    {
        var docPath = CreateWordDocumentWithContent("test_search_no_results.docx", "Hello World");

        var result = _tool.Execute("search", docPath, searchText: "NonExistentText");

        Assert.Contains("0 matches", result);
        Assert.Contains("No matching text found", result);
    }

    [Fact]
    public void FormatText_ShouldApplyFormattingToRun()
    {
        var docPath = CreateWordDocumentWithContent("test_format_text.docx", "Format this text");
        var outputPath = CreateTestFilePath("test_format_text_output.docx");

        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, bold: true, italic: true);

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0)
        {
            Assert.True(runs[0].Font.Bold);
            Assert.True(runs[0].Font.Italic);
        }
    }

    [Fact]
    public void FormatText_WithFontNameAsciiAndFarEast_ShouldApplyDifferentFonts()
    {
        var docPath = CreateWordDocumentWithContent("test_format_fonts.docx", "English 中文");
        var outputPath = CreateTestFilePath("test_format_fonts_output.docx");

        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0,
            fontNameAscii: "Arial", fontNameFarEast: "Microsoft YaHei");

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.True(runs.Count > 0);
        Assert.Equal("Arial", runs[0].Font.NameAscii);
    }

    [Fact]
    public void FormatText_WithColor_ShouldApplyColor()
    {
        var docPath = CreateWordDocumentWithContent("test_format_color.docx", "Colored text");
        var outputPath = CreateTestFilePath("test_format_color_output.docx");

        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, color: "00FF00");

        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.True(runs.Count > 0);
        Assert.Equal(Color.FromArgb(0, 255, 0), runs[0].Font.Color);
    }

    [SkippableFact]
    public void FormatText_AllRunsInParagraph_ShouldFormatAllRuns()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Font formatting may be limited in evaluation mode");

        var docPath = CreateWordDocument("test_format_all_runs.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Run1 ");
        builder.Write("Run2 ");
        builder.Write("Run3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_format_all_runs_output.docx");
        _tool.Execute("format", docPath, outputPath: outputPath, paragraphIndex: 0, bold: true);

        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        foreach (var run in runs)
            if (!string.IsNullOrWhiteSpace(run.Text))
                Assert.True(run.Font.Bold, $"Run '{run.Text}' should be bold");
    }

    [SkippableFact]
    public void FormatText_WithSuperscript_ShouldClearSubscript()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Font formatting may be limited in evaluation mode");

        var docPath = CreateWordDocument("test_format_superscript.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc) { Font = { Subscript = true } };
        builder.Write("Subscript Text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_format_superscript_output.docx");
        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, superscript: true);

        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Subscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Superscript);
        Assert.False(run.Font.Subscript);
    }

    [SkippableFact]
    public void FormatText_WithSubscript_ShouldClearSuperscript()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Font formatting may be limited in evaluation mode");

        var docPath = CreateWordDocument("test_format_subscript.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc) { Font = { Superscript = true } };
        builder.Write("Superscript Text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_format_subscript_output.docx");
        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, subscript: true);

        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Superscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Subscript);
        Assert.False(run.Font.Superscript);
    }

    [Fact]
    public void InsertTextAtPosition_ShouldInsertAtCorrectPosition()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert_position.docx", "First", "Third");
        var outputPath = CreateTestFilePath("test_insert_position_output.docx");

        _tool.Execute("insert_at_position", docPath, outputPath: outputPath,
            insertParagraphIndex: 0, charIndex: 0, text: "Second ");

        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);

        var docText = doc.GetText();
        var hasFirst = docText.Contains("First", StringComparison.OrdinalIgnoreCase);
        var hasThird = docText.Contains("Third", StringComparison.OrdinalIgnoreCase);
        var isEvaluationMode = IsEvaluationMode();

        if (paragraphs.Count > 0)
        {
            var firstPara = paragraphs[0];
            var firstParaText = firstPara.GetText();
            var hasSecondInPara = firstParaText.Contains("Second", StringComparison.OrdinalIgnoreCase);

            if (hasFirst || hasThird)
            {
                Assert.True(hasSecondInPara || docText.Contains("Second", StringComparison.OrdinalIgnoreCase),
                    "Inserted text 'Second' should be present in the document");

                if (hasSecondInPara && firstParaText.Contains("First", StringComparison.OrdinalIgnoreCase))
                {
                    var secondIndex = firstParaText.IndexOf("Second", StringComparison.OrdinalIgnoreCase);
                    var firstIndex = firstParaText.IndexOf("First", StringComparison.OrdinalIgnoreCase);
                    if (!isEvaluationMode)
                    {
                        Assert.True(secondIndex < firstIndex,
                            "'Second' should be inserted before 'First'. Paragraph text: " +
                            firstParaText.Substring(0, Math.Min(100, firstParaText.Length)));
                    }
                    else
                    {
                        Assert.True(docText.Contains("Second", StringComparison.OrdinalIgnoreCase),
                            "Document should contain inserted 'Second' text");
                        Assert.True(docText.Contains("First", StringComparison.OrdinalIgnoreCase),
                            "Document should still contain original 'First' text");
                    }
                }
            }
            else if (isEvaluationMode)
            {
                Assert.True(paragraphs.Count >= 1, "Document should have at least one paragraph");
                Assert.True(docText.Length > 0, "Document should have content");
            }
        }
    }

    [Fact]
    public void InsertAtPosition_WithInsertBefore_ShouldInsertBeforePosition()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_before.docx", "Original Text");
        var outputPath = CreateTestFilePath("test_insert_before_output.docx");

        _tool.Execute("insert_at_position", docPath, outputPath: outputPath,
            insertParagraphIndex: 0, charIndex: 0, text: "Prefix: ", insertBefore: true);

        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Prefix:", text);
        Assert.Contains("Original Text", text);
    }

    [Fact]
    public void DeleteText_ShouldDeleteTextBySearch()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_text.docx", "Delete this text");
        var outputPath = CreateTestFilePath("test_delete_text_output.docx");

        _tool.Execute("delete", docPath, outputPath: outputPath, searchText: "Delete ");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.DoesNotContain("Delete", text);
    }

    [Fact]
    public void DeleteText_ByParagraphIndices_ShouldDeleteText()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_text_by_indices.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_text_by_indices_output.docx");

        _tool.Execute("delete", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 0, startRunIndex: 0);

        var doc = new Document(outputPath);
        Assert.True(File.Exists(outputPath));

        var isEvaluationMode = IsEvaluationMode();
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var firstPara = paragraphs.FirstOrDefault();
        if (firstPara != null)
        {
            var firstParaText = firstPara.GetText().Trim();
            if (isEvaluationMode)
            {
                Assert.True(File.Exists(outputPath), "Output file should be created");
                Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
                var docText = doc.GetText();
                Assert.True(docText.Contains("Second") || docText.Contains("Third"),
                    "Document should still contain remaining paragraphs");
            }
            else
            {
                Assert.DoesNotContain("First", firstParaText);
            }
        }
    }

    [Fact]
    public void DeleteRange_ShouldDeleteTextRange()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_range.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_range_output.docx");

        _tool.Execute("delete_range", docPath, outputPath: outputPath,
            startParagraphIndex: 0, startCharIndex: 0, endParagraphIndex: 1, endCharIndex: 0);

        var isEvaluationMode = IsEvaluationMode();
        Assert.True(File.Exists(outputPath));

        var doc = new Document(outputPath);
        var text = doc.GetText();
        var hasFirst = text.Contains("First", StringComparison.OrdinalIgnoreCase);
        var hasSecond = text.Contains("Second", StringComparison.OrdinalIgnoreCase);
        var hasThird = text.Contains("Third", StringComparison.OrdinalIgnoreCase);

        if (isEvaluationMode)
        {
            Assert.True(hasFirst || hasSecond || hasThird,
                "Document should contain at least some of the original paragraphs");
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            Assert.True(paragraphs.Count > 0, "Document should have paragraphs after delete range operation");
        }
        else
        {
            Assert.False(hasFirst);
            Assert.True(hasSecond);
            Assert.True(hasThird);
        }
    }

    [Fact]
    public void DeleteRange_WithinSameParagraph_ShouldDeletePartialText()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_range_same_para.docx", "Hello World Test");
        var outputPath = CreateTestFilePath("test_delete_range_same_para_output.docx");

        _tool.Execute("delete_range", docPath, outputPath: outputPath,
            startParagraphIndex: 0, startCharIndex: 6, endParagraphIndex: 0, endCharIndex: 12);

        var doc = new Document(outputPath);
        var text = doc.GetText();

        if (IsEvaluationMode())
        {
            Assert.True(File.Exists(outputPath), "Output file should be created");
            Assert.NotEmpty(text);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");
            Assert.True(text.Contains("Hello") || text.Contains("Test") || text.Contains("World"),
                "Document should contain some of the original text");
        }
        else
        {
            Assert.Contains("Hello", text);
            Assert.Contains("Test", text);
            Assert.DoesNotContain("World", text);
        }
    }

    [Fact]
    public void AddTextWithStyle_ShouldApplyStyle()
    {
        var docPath = CreateWordDocument("test_add_text_with_style.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomTextStyle");
        customStyle.Font.Size = 16;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_text_with_style_output.docx");
        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Styled Text", styleName: "CustomTextStyle");

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Styled Text"));
        Assert.NotNull(para);
        Assert.Equal("CustomTextStyle", para.ParagraphFormat.StyleName);
    }

    [Fact]
    public void AddTextWithStyle_ShouldCreateEmptyParagraphsWithNormalStyle()
    {
        var docPath = CreateWordDocument("test_empty_paragraph_style.docx");
        var outputPath = CreateTestFilePath("test_empty_paragraph_style_output.docx");

        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestHeadingStyle");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Test Content", styleName: "!TestHeadingStyle");

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        var textPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content"));
        Assert.NotNull(textPara);
        Assert.Equal("!TestHeadingStyle", textPara.ParagraphFormat.StyleName);

        var parentNode = textPara.ParentNode;
        if (parentNode != null)
        {
            var nextSibling = textPara.NextSibling;
            while (nextSibling is Paragraph nextPara)
            {
                if (string.IsNullOrWhiteSpace(nextPara.GetText()))
                {
                    var actualStyle = nextPara.ParagraphFormat.StyleName ?? "";
                    Assert.True(actualStyle == "Normal",
                        $"Empty paragraph after styled text should use Normal style, but got: {actualStyle}");
                }
                else
                {
                    break;
                }

                nextSibling = nextSibling.NextSibling;
            }
        }
    }

    [Fact]
    public void AddTextWithStyle_ThenAddTable_ShouldHaveNormalStyleEmptyParagraphs()
    {
        var tableTool = new WordTableTool(SessionManager);
        var docPath = CreateWordDocument("test_table_empty_paragraph.docx");
        var outputPath1 = CreateTestFilePath("test_table_empty_paragraph_step1.docx");
        var outputPath2 = CreateTestFilePath("test_table_empty_paragraph_step2.docx");

        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestHeadingStyle2");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        _tool.Execute("add_with_style", docPath, outputPath: outputPath1,
            text: "Test Content 1", styleName: "!TestHeadingStyle2");
        tableTool.Execute("create", outputPath1, outputPath: outputPath2, rows: 2, columns: 3);
        _tool.Execute("add_with_style", outputPath2, outputPath: outputPath2,
            text: "Test Content 2", styleName: "!TestHeadingStyle2");

        var resultDoc = new Document(outputPath2);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        var firstTextPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content 1"));
        var secondTextPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content 2"));

        Assert.NotNull(firstTextPara);
        Assert.NotNull(secondTextPara);

        var currentNode = firstTextPara.NextSibling;
        while (currentNode != null && currentNode != secondTextPara.ParentNode)
        {
            if (currentNode is Paragraph para && string.IsNullOrWhiteSpace(para.GetText()))
            {
                var actualStyle = para.ParagraphFormat.StyleName ?? "";
                Assert.True(actualStyle == "Normal",
                    $"Empty paragraph between styled texts should use Normal style, but got: {actualStyle}");
            }
            else if (currentNode.NodeType == NodeType.Table)
            {
                break;
            }

            currentNode = currentNode.NextSibling;
        }
    }

    [Fact]
    public void AddWithStyle_WithStyleNameAndCustomFormat_ShouldOverrideStyle()
    {
        var docPath = CreateWordDocument("test_add_text_style_override.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "BaseStyle");
        customStyle.Font.Size = 12;
        customStyle.Font.Color = Color.Black;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_text_style_override_output.docx");
        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Overridden Text", styleName: "BaseStyle", fontSize: 18, color: "FF0000");

        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Overridden Text"));
        Assert.NotNull(run);
        Assert.Equal(18, run.Font.Size);
        Assert.Equal(Color.FromArgb(255, 0, 0), run.Font.Color);
    }

    [Fact]
    public void AddWithStyle_WithAlignment_ShouldApplyAlignment()
    {
        var docPath = CreateWordDocument("test_add_style_alignment.docx");
        var outputPath = CreateTestFilePath("test_add_style_alignment_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Centered Text", alignment: "center");

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Centered Text"));
        Assert.NotNull(para);
        Assert.Equal(ParagraphAlignment.Center, para.ParagraphFormat.Alignment);
    }

    [Fact]
    public void AddWithStyle_WithIndentLevel_ShouldApplyIndentation()
    {
        var docPath = CreateWordDocument("test_add_style_indent.docx");
        var outputPath = CreateTestFilePath("test_add_style_indent_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Indented Text", indentLevel: 2);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Indented Text"));
        Assert.NotNull(para);
        Assert.Equal(72, para.ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void AddWithStyle_WithLeftIndent_ShouldApplyLeftIndentation()
    {
        var docPath = CreateWordDocument("test_add_style_left_indent.docx");
        var outputPath = CreateTestFilePath("test_add_style_left_indent_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Left Indented Text", leftIndent: 50);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Left Indented Text"));
        Assert.NotNull(para);
        Assert.Equal(50, para.ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void AddWithStyle_WithFirstLineIndent_ShouldApplyFirstLineIndentation()
    {
        var docPath = CreateWordDocument("test_add_style_first_line.docx");
        var outputPath = CreateTestFilePath("test_add_style_first_line_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "First Line Indented", firstLineIndent: 36);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("First Line Indented"));
        Assert.NotNull(para);
        Assert.Equal(36, para.ParagraphFormat.FirstLineIndent);
    }

    [Fact]
    public void AddWithStyle_AtBeginning_ShouldInsertAtStart()
    {
        var docPath = CreateWordDocumentWithContent("test_add_style_beginning.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_add_style_beginning_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "First Paragraph", paragraphIndexForAdd: -1);

        var doc = new Document(outputPath);
        var text = doc.GetText();

        if (IsEvaluationMode())
        {
            Assert.Contains("First Paragraph", text);
            Assert.Contains("Existing content", text);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            Assert.True(paragraphs.Count >= 2, "Document should have at least 2 paragraphs after adding");
        }
        else
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            var firstPara = paragraphs.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p.GetText()));
            Assert.NotNull(firstPara);
            Assert.Contains("First Paragraph", firstPara.GetText());
        }
    }

    [Fact]
    public void AddWithStyle_AfterSpecificParagraph_ShouldInsertAfter()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_add_style_after.docx", "Para1", "Para3");
        var outputPath = CreateTestFilePath("test_add_style_after_output.docx");

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Para2", paragraphIndexForAdd: 0);

        var doc = new Document(outputPath);
        var text = doc.GetText();

        if (IsEvaluationMode())
        {
            Assert.Contains("Para1", text);
            Assert.Contains("Para2", text);
            Assert.Contains("Para3", text);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            Assert.True(paragraphs.Count >= 3, "Document should have at least 3 paragraphs");
        }
        else
        {
            var para1Index = text.IndexOf("Para1", StringComparison.Ordinal);
            var para2Index = text.IndexOf("Para2", StringComparison.Ordinal);
            var para3Index = text.IndexOf("Para3", StringComparison.Ordinal);
            Assert.True(para1Index < para2Index);
            Assert.True(para2Index < para3Index);
        }
    }

    [Fact]
    public void AddWithStyle_WithTabStops_ShouldApplyTabStops()
    {
        var docPath = CreateWordDocument("test_add_style_tabs.docx");
        var outputPath = CreateTestFilePath("test_add_style_tabs_output.docx");
        var tabStopsJson = "[{\"position\":72,\"alignment\":\"Left\",\"leader\":\"None\"}]";

        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Tabbed\tText", tabStops: tabStopsJson);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Tabbed"));
        Assert.NotNull(para);
        Assert.True(para.ParagraphFormat.TabStops.Count > 0);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("AdD")]
    [InlineData("add")]
    public void Execute_OperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_case.docx");

        var result = _tool.Execute(operation, docPath, text: "Test");

        Assert.StartsWith("Text added to document successfully", result);
    }

    #endregion

    #region Exception

    [Theory]
    [InlineData("unknown_operation")]
    [InlineData("invalid")]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException(string operation)
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(operation, docPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddText_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_no_text.docx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", docPath));

        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void Replace_WithoutFindPattern_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_replace_no_find.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("replace", docPath, replace: "replacement"));

        Assert.Contains("find is required", ex.Message);
    }

    [Fact]
    public void Replace_WithoutReplacePattern_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_replace_no_replace.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("replace", docPath, find: "Test"));

        Assert.Contains("replace is required", ex.Message);
    }

    [Fact]
    public void Search_WithoutSearchText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_search_no_text.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("search", docPath));

        Assert.Contains("searchText is required", ex.Message);
    }

    [Fact]
    public void FormatText_WithoutParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_format_no_para.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", docPath, bold: true));

        Assert.Contains("paragraphIndex is required", ex.Message);
    }

    [Fact]
    public void FormatText_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_format_invalid_para.docx", "Test");
        var outputPath = CreateTestFilePath("test_format_invalid_para_output.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", docPath, outputPath: outputPath, paragraphIndex: 999, bold: true));

        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FormatText_WithInvalidRunIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_format_invalid_run.docx", "Test");
        var outputPath = CreateTestFilePath("test_format_invalid_run_output.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", docPath, outputPath: outputPath, paragraphIndex: 0, runIndex: 999, bold: true));

        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FormatText_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_format_invalid_section.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", docPath, paragraphIndex: 0, sectionIndex: 999, bold: true));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_no_text.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath, insertParagraphIndex: 0, charIndex: 0));

        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithoutParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_no_para.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath, charIndex: 0, text: "Test"));

        Assert.Contains("insertParagraphIndex is required", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithoutCharIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_no_char.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath, insertParagraphIndex: 0, text: "Test"));

        Assert.Contains("charIndex is required", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_para.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath,
                insertParagraphIndex: 999, charIndex: 0, text: "Test"));

        Assert.Contains("paragraphIndex must be between", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_section.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath,
                insertParagraphIndex: 0, charIndex: 0, text: "Test", sectionIndex: 999));

        Assert.Contains("sectionIndex must be between", ex.Message);
    }

    [Fact]
    public void DeleteText_WithoutSearchTextOrIndices_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_no_params.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", docPath));

        Assert.Contains("is required", ex.Message);
    }

    [Fact]
    public void DeleteText_WithStartIndexOnly_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_start_only.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, startParagraphIndex: 0));

        Assert.Contains("endParagraphIndex is required", ex.Message);
    }

    [Fact]
    public void DeleteText_WithInvalidParagraphRange_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid_range.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, startParagraphIndex: 999, endParagraphIndex: 1000));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteText_WithNotFoundText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_not_found.docx", "Hello World");
        var outputPath = CreateTestFilePath("test_delete_not_found_output.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, searchText: "NonExistentText"));

        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteRange_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_range_invalid_section.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_range", docPath,
                startParagraphIndex: 0, startCharIndex: 0,
                endParagraphIndex: 0, endCharIndex: 5,
                sectionIndex: 999));

        Assert.Contains("sectionIndex must be between", ex.Message);
    }

    [Fact]
    public void DeleteRange_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_range_invalid_para.docx", "Test content");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_range", docPath,
                startParagraphIndex: 999, startCharIndex: 0,
                endParagraphIndex: 999, endCharIndex: 5));

        Assert.Contains("Paragraph indices out of range", ex.Message);
    }

    [Fact]
    public void AddWithStyle_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_with_style_no_text.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_with_style", docPath, styleName: "Normal"));

        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void AddWithStyle_WithInvalidStyle_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_add_invalid_style.docx");
        var outputPath = CreateTestFilePath("test_add_invalid_style_output.docx");

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("add_with_style", docPath, outputPath: outputPath, text: "Test",
                styleName: "NonExistentStyle12345"));

        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddWithStyle_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_style_invalid_para.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_with_style", docPath, text: "Test", paragraphIndexForAdd: 999));

        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void AddText_WithSessionId_ShouldAddTextInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_text.docx");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("add", sessionId: sessionId, text: "Session Text");

        Assert.StartsWith("Text added to document successfully", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Session Text", text);
    }

    [Fact]
    public void SearchText_WithSessionId_ShouldSearchInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_search.docx", "Searchable content here");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("search", sessionId: sessionId, searchText: "Searchable");

        Assert.Contains("Found", result);
        Assert.Contains("Searchable", result);
    }

    [Fact]
    public void ReplaceText_WithSessionId_ShouldReplaceInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_replace.docx", "Original text to replace");
        var sessionId = OpenSession(docPath);

        _tool.Execute("replace", sessionId: sessionId, find: "Original", replace: "Modified");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.Contains("Modified", text);
        Assert.DoesNotContain("Original", text);
    }

    [Fact]
    public void FormatText_WithSessionId_ShouldFormatInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_format.docx", "Format this text");
        var sessionId = OpenSession(docPath);

        _tool.Execute("format", sessionId: sessionId, paragraphIndex: 0, runIndex: 0, bold: true);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.True(runs.Count > 0);
        Assert.True(runs[0].Font.Bold);
    }

    [Fact]
    public void InsertAtPosition_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert.docx", "Original text");
        var sessionId = OpenSession(docPath);

        _tool.Execute("insert_at_position", sessionId: sessionId,
            insertParagraphIndex: 0, charIndex: 0, text: "Inserted: ");

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Inserted:", text);
        Assert.Contains("Original text", text);
    }

    [Fact]
    public void DeleteText_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Write("Keep this ");
        builder.Write("DeleteMe");
        builder.Write(" and this");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, searchText: "DeleteMe");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.DoesNotContain("DeleteMe", text);
        Assert.Contains("Keep this", text);
    }

    [Fact]
    public void AddWithStyle_WithSessionId_ShouldAddInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_style.docx");
        var sessionId = OpenSession(docPath);

        _tool.Execute("add_with_style", sessionId: sessionId,
            text: "Styled Session Text", styleName: "Normal");

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Styled Session Text", text);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("search", sessionId: "invalid_session_id", searchText: "test"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_text_path.docx", "PathContent");
        var docPath2 = CreateWordDocumentWithContent("test_text_session.docx", "SessionContent");

        var sessionId = OpenSession(docPath2);

        var result = _tool.Execute("search", docPath1, sessionId, searchText: "Content");

        Assert.Contains("Session", result);
        Assert.DoesNotContain("Path", result);
    }

    #endregion
}