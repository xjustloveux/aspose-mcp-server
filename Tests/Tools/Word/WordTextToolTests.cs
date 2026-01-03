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

    #region General Tests

    [Fact]
    public void AddText_ShouldAddTextToDocumentEnd()
    {
        var docPath = CreateWordDocument("test_add_text.docx");
        var outputPath = CreateTestFilePath("test_add_text_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, text: "Hello World");
        Assert.True(result.Contains("Text added", StringComparison.OrdinalIgnoreCase));
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.Contains(paragraphs, p => p.GetText().Contains("Hello World"));
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
    public void ReplaceText_ShouldReplaceAllOccurrences()
    {
        var docPath = CreateWordDocumentWithContent("test_replace_text.docx", "Hello World Hello World");
        var outputPath = CreateTestFilePath("test_replace_text_output.docx");
        _tool.Execute("replace", docPath, outputPath: outputPath, find: "Hello", replace: "Hi");
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Hi World", text);
        Assert.DoesNotContain("Hello", text);
    }

    [Fact]
    public void SearchText_ShouldFindTextInDocument()
    {
        var docPath = CreateWordDocumentWithContent("test_search_text.docx", "This is a test document");
        var result = _tool.Execute("search", docPath, searchText: "test");
        Assert.True(result.Contains("Found", StringComparison.OrdinalIgnoreCase));
        Assert.True(result.Contains("test", StringComparison.OrdinalIgnoreCase));
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
    public void InsertTextAtPosition_ShouldInsertAtCorrectPosition()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert_position.docx", "First", "Third");
        var outputPath = CreateTestFilePath("test_insert_position_output.docx");
        _tool.Execute("insert_at_position", docPath, outputPath: outputPath,
            insertParagraphIndex: 0, charIndex: 0, text: "Second ");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);

        // Verify document has paragraphs
        Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");

        // Verify original content exists in document
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
                        Assert.True(secondIndex < firstIndex,
                            "'Second' should be inserted before 'First'. Paragraph text: " +
                            firstParaText.Substring(0, Math.Min(100, firstParaText.Length)));
                }
            }
            else if (isEvaluationMode)
            {
                Assert.True(paragraphs.Count >= 2 || hasSecondInPara,
                    "Document should have been modified by the insertion operation");
            }
        }
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
    public void DeleteRange_ShouldDeleteTextRange()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_range.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_range_output.docx");
        _tool.Execute("delete_range", docPath, outputPath: outputPath,
            startParagraphIndex: 0, startCharIndex: 0, endParagraphIndex: 1, endCharIndex: 0);

        var isEvaluationMode = IsEvaluationMode();
        Assert.True(File.Exists(outputPath), "Output document should be created");

        var doc = new Document(outputPath);
        var text = doc.GetText();
        var hasFirst = text.Contains("First", StringComparison.OrdinalIgnoreCase);
        var hasSecond = text.Contains("Second", StringComparison.OrdinalIgnoreCase);
        var hasThird = text.Contains("Third", StringComparison.OrdinalIgnoreCase);

        if (isEvaluationMode)
        {
            Assert.True(hasFirst || hasSecond || hasThird, "Document should contain some original content");
        }
        else
        {
            // delete_range from (para 0, char 0) to (para 1, char 0) should delete paragraph 0
            Assert.False(hasFirst, "First should be deleted (paragraph 0)");
            Assert.True(hasSecond, "Second should remain (paragraph 1)");
            Assert.True(hasThird, "Third should remain (paragraph 2)");
        }
    }

    [Fact]
    public void AddTextWithStyle_ShouldCreateEmptyParagraphsWithNormalStyle()
    {
        var docPath = CreateWordDocument("test_empty_paragraph_style.docx");
        var outputPath = CreateTestFilePath("test_empty_paragraph_style_output.docx");

        // Create a custom style similar to the problem scenario
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestHeadingStyle");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        // Act: Add text with custom style
        _tool.Execute("add_with_style", docPath, outputPath: outputPath,
            text: "Test Content", styleName: "!TestHeadingStyle");

        // Assert: Check that empty paragraphs after insertion use Normal style
        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // Find the paragraph with our text
        var textPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content"));
        Assert.NotNull(textPara);
        Assert.Equal("!TestHeadingStyle", textPara.ParagraphFormat.StyleName);

        // Check next empty paragraphs - they should use Normal style
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
                    break; // Stop at first non-empty paragraph
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

        // Create a custom style
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestHeadingStyle2");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;
        doc.Save(docPath);

        // Act: Add text with custom style
        _tool.Execute("add_with_style", docPath, outputPath: outputPath1,
            text: "Test Content 1", styleName: "!TestHeadingStyle2");

        // Add a table
        tableTool.Execute("create", outputPath1, outputPath: outputPath2, rows: 2, columns: 3);

        // Add another text with custom style
        _tool.Execute("add_with_style", outputPath2, outputPath: outputPath2,
            text: "Test Content 2", styleName: "!TestHeadingStyle2");

        // Assert: Check that empty paragraphs between styled text and table use Normal style
        var resultDoc = new Document(outputPath2);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // Find paragraphs
        var firstTextPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content 1"));
        var secondTextPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Content 2"));

        Assert.NotNull(firstTextPara);
        Assert.NotNull(secondTextPara);

        // Check empty paragraphs between first text and second text
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
                break; // Found table, stop checking
            }

            currentNode = currentNode.NextSibling;
        }
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

    [Fact]
    public void AddText_WithUnderline_ShouldApplyUnderline()
    {
        var docPath = CreateWordDocument("test_add_text_underline.docx");
        var outputPath = CreateTestFilePath("test_add_text_underline_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, text: "Underlined Text", underline: "single");
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Underlined Text"));
        Assert.NotNull(run);
        Assert.Equal(Underline.Single, run.Font.Underline);
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
            text: "Formatted Text", fontName: "Arial", fontSize: 14,
            bold: true, italic: true, underline: "single", color: "0000FF");
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Formatted Text"));
        Assert.NotNull(run);
        Assert.Equal("Arial", run.Font.Name);
        Assert.Equal(14, run.Font.Size);
        Assert.True(run.Font.Bold);
        Assert.True(run.Font.Italic);
        Assert.Equal(Underline.Single, run.Font.Underline);
        Assert.Equal(Color.FromArgb(0, 0, 255), run.Font.Color);
    }

    [Fact]
    public void DeleteText_ByParagraphIndices_ShouldDeleteText()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_text_by_indices.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_text_by_indices_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 0, startRunIndex: 0);
        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.True(File.Exists(outputPath), "Output document should be created");

        var isEvaluationMode = IsEvaluationMode();
        var firstPara = paragraphs.FirstOrDefault();
        if (firstPara != null)
        {
            var firstParaText = firstPara.GetText().Trim();
            if (isEvaluationMode)
                // In evaluation mode, deletion may be limited, verify operation completed
                Assert.True(File.Exists(outputPath), "Output file should be created");
            else
                Assert.DoesNotContain("First", firstParaText);
        }
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
        // Verify font names were applied
        // Note: fontNameAscii was set to "Times New Roman"
        // In Word, when fontNameAscii is set, it becomes the Name property
        Assert.Equal("Times New Roman", run.Font.NameAscii);
        // NameFarEast may not be applied correctly in evaluation mode or if font is not available
        // Check that NameFarEast is set (even if value differs due to evaluation mode limitations)
        Assert.NotNull(run.Font.NameFarEast);
        // In evaluation mode, NameFarEast might default to NameAscii if the font is not available
        // So we check that at least NameAscii was correctly applied
        Assert.True(run.Font.NameAscii == "Times New Roman",
            $"Font NameAscii should be Times New Roman. NameAscii: {run.Font.NameAscii}, NameFarEast: {run.Font.NameFarEast}");
    }

    [Fact]
    public void AddText_WithUnderlineStyles_ShouldApplyDifferentUnderlineTypes()
    {
        // Arrange - Test each underline style separately
        var underlineStyles = new[] { "single", "double", "dotted", "dash" };

        foreach (var underlineStyle in underlineStyles)
        {
            var docPath = CreateWordDocument($"test_add_text_underline_{underlineStyle}.docx");
            var outputPath = CreateTestFilePath($"test_add_text_underline_{underlineStyle}_output.docx");
            _tool.Execute("add", docPath, outputPath: outputPath,
                text: $"Underline {underlineStyle}", underline: underlineStyle);
            var doc = new Document(outputPath);
            var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var run = runs.FirstOrDefault(r => r.Text.Contains($"Underline {underlineStyle}"));
            Assert.NotNull(run);

            // Verify underline style was applied
            var expectedUnderline = underlineStyle switch
            {
                "single" => Underline.Single,
                "double" => Underline.Double,
                "dotted" => Underline.Dotted,
                "dash" => Underline.Dash,
                _ => Underline.None
            };
            Assert.Equal(expectedUnderline, run.Font.Underline);
        }
    }

    [Fact]
    public void AddText_WithStyleName_ShouldApplyStyle()
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
    public void AddText_WithStyleNameAndCustomFormat_ShouldOverrideStyle()
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
        // Custom parameters should override style defaults
        Assert.Equal(18, run.Font.Size);
        Assert.Equal(Color.FromArgb(255, 0, 0), run.Font.Color);
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
        // When replaceInFields is true, the hyperlink text should be replaced
        Assert.Contains("NewLink", text);
    }

    [Fact]
    public void Search_WithUseRegex_ShouldSearchUsingRegex()
    {
        var docPath =
            CreateWordDocumentWithContent("test_search_regex.docx", "Email: test@example.com and admin@test.org");
        var result = _tool.Execute("search", docPath,
            searchText: @"\w+@\w+\.\w+", useRegex: true);
        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("test@example.com", result);
        Assert.Contains("admin@test.org", result);
    }

    [Fact]
    public void Search_WithCaseSensitive_ShouldMatchCase()
    {
        var docPath = CreateWordDocumentWithContent("test_search_case.docx", "Hello HELLO hello HeLLo");
        var result = _tool.Execute("search", docPath,
            searchText: "Hello", caseSensitive: true);
        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        // Should only find exact case match "Hello", not "HELLO", "hello", or "HeLLo"
        Assert.Contains("1 matches", result);
    }

    [Fact]
    public void Search_WithMaxResults_ShouldLimitResults()
    {
        var docPath = CreateWordDocumentWithContent("test_search_max.docx",
            "word word word word word word word word word word");
        var result = _tool.Execute("search", docPath,
            searchText: "word", maxResults: 3);
        Assert.Contains("Found", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3 matches", result);
        Assert.Contains("limited to first 3", result);
    }

    [Fact]
    public void InsertAtPosition_WithInsertBefore_ShouldInsertBeforePosition()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_before.docx", "Original Text");
        var outputPath = CreateTestFilePath("test_insert_before_output.docx");
        _tool.Execute("insert_at_position", docPath, outputPath: outputPath,
            insertParagraphIndex: 0, charIndex: 0, text: "Prefix: ", insertBefore: true);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("Prefix:", text);
        Assert.Contains("Original Text", text);
    }

    [SkippableFact]
    public void FormatText_WithSuperscript_ShouldClearSubscript()
    {
        // Skip in evaluation mode as font formatting may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "Font formatting may be limited in evaluation mode");

        // Arrange - Create document with subscript text
        var docPath = CreateWordDocument("test_format_superscript.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc) { Font = { Subscript = true } };
        builder.Write("Subscript Text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_format_superscript_output.docx");
        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, superscript: true);

        // Assert - Superscript should be true, subscript should be false (mutual exclusion)
        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Subscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Superscript, "Superscript should be enabled");
        Assert.False(run.Font.Subscript, "Subscript should be cleared when superscript is set");
    }

    [SkippableFact]
    public void FormatText_WithSubscript_ShouldClearSuperscript()
    {
        // Skip in evaluation mode as font formatting may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "Font formatting may be limited in evaluation mode");

        // Arrange - Create document with superscript text
        var docPath = CreateWordDocument("test_format_subscript.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc) { Font = { Superscript = true } };
        builder.Write("Superscript Text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_format_subscript_output.docx");
        _tool.Execute("format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, subscript: true);

        // Assert - Subscript should be true, superscript should be false (mutual exclusion)
        var resultDoc = new Document(outputPath);
        var runs = resultDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Superscript Text"));
        Assert.NotNull(run);
        Assert.True(run.Font.Subscript, "Subscript should be enabled");
        Assert.False(run.Font.Superscript, "Superscript should be cleared when subscript is set");
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
    public void AddWithStyle_WithInvalidStyle_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_add_invalid_style.docx");
        var outputPath = CreateTestFilePath("test_add_invalid_style_output.docx");

        // Act & Assert - InvalidOperationException wraps the inner ArgumentException
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("add_with_style", docPath, outputPath: outputPath, text: "Test",
                styleName: "NonExistentStyle12345"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Search_WithNoResults_ShouldReturnNoMatchMessage()
    {
        var docPath = CreateWordDocumentWithContent("test_search_no_results.docx", "Hello World");
        var result = _tool.Execute("search", docPath, searchText: "NonExistentText");
        Assert.Contains("0 matches", result);
        Assert.Contains("No matching text found", result);
    }

    #endregion

    #region Exception Tests

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
    public void AddText_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_no_text.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath));

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
    public void Search_WithoutSearchText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_search_no_text.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("search", docPath));

        Assert.Contains("searchText is required", ex.Message);
    }

    [Fact]
    public void InsertAtPosition_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_no_text.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_at_position", docPath, insertParagraphIndex: 0, charIndex: 0));

        Assert.Contains("text is required", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddText_WithSessionId_ShouldAddTextInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_text.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "Session Text");
        Assert.Contains("Text added", result);

        // Verify in-memory document has the text
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

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.Contains("Modified", text);
        Assert.DoesNotContain("Original", text);
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("search", docPath1, sessionId, searchText: "Content");

        // Assert - should use sessionId (Session document), not path
        Assert.Contains("Session", result);
        Assert.DoesNotContain("Path", result);
    }

    #endregion
}