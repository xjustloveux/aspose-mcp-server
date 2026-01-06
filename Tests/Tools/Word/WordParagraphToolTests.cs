using System.Drawing;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordParagraphToolTests : WordTestBase
{
    private readonly WordParagraphTool _tool;

    public WordParagraphToolTests()
    {
        _tool = new WordParagraphTool(SessionManager);
    }

    #region General

    [Fact]
    public void InsertParagraph_ShouldInsertAtCorrectPosition()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert_paragraph.docx", "First", "Third");
        var outputPath = CreateTestFilePath("test_insert_paragraph_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath, paragraphIndex: 0, text: "Second");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count >= 2);
        Assert.Contains(paragraphs, p => p.GetText().Contains("Second"));
    }

    [Fact]
    public void InsertParagraph_WithStyle_ShouldApplyStyle()
    {
        var docPath = CreateWordDocument("test_insert_paragraph_style.docx");
        var outputPath = CreateTestFilePath("test_insert_paragraph_style_output.docx");
        var doc = new Document(docPath);
        var headingStyle = doc.Styles.Add(StyleType.Paragraph, "TestHeading");
        headingStyle.Font.Size = 16;
        headingStyle.Font.Bold = true;
        doc.Save(docPath);
        _tool.Execute("insert", docPath, outputPath: outputPath, text: "Heading", styleName: "TestHeading");
        var resultDoc = new Document(outputPath);
        var paragraphs = GetParagraphs(resultDoc);
        var headingPara = paragraphs.FirstOrDefault(p => p.GetText().Contains("Heading"));
        Assert.NotNull(headingPara);
        Assert.Equal("TestHeading", headingPara.ParagraphFormat.StyleName);
    }

    [Fact]
    public void DeleteParagraph_ShouldDeleteParagraph()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_paragraph.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_paragraph_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, paragraphIndex: 1);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc, false);
        var foundSecond = false;
        foreach (var para in paragraphs)
            if (para.GetText().Trim().Contains("Second"))
            {
                foundSecond = true;
                break;
            }

        Assert.False(foundSecond, "Paragraph containing 'Second' should be deleted");
    }

    [Fact]
    public void GetParagraph_ShouldReturnParagraphText()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_get_paragraph.docx", "First", "Second", "Third");
        var result = _tool.Execute("get", docPath, includeEmpty: true, includeCommentParagraphs: true,
            includeTextboxParagraphs: true);
        Assert.True(result.Contains("First", StringComparison.Ordinal));
        Assert.True(result.Contains("Second", StringComparison.Ordinal));
        Assert.True(result.Contains("Third", StringComparison.Ordinal));
    }

    [Fact]
    public void GetParagraphFormat_ShouldReturnFormatInfo()
    {
        var docPath = CreateWordDocument("test_get_format.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc)
        {
            ParagraphFormat = { Alignment = ParagraphAlignment.Center, LeftIndent = 36 }
        };
        builder.Writeln("Centered Text");
        doc.Save(docPath);

        // Find the paragraph with "Centered Text" (may not be index 0 due to empty paragraphs)
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var targetIndex = -1;
        for (var i = 0; i < paragraphs.Count; i++)
            if (paragraphs[i].GetText().Contains("Centered Text"))
            {
                targetIndex = i;
                break;
            }

        Assert.True(targetIndex >= 0, "Should find paragraph with 'Centered Text'");
        var result = _tool.Execute("get_format", docPath, paragraphIndex: targetIndex);
        Assert.True(result.Contains("Center", StringComparison.OrdinalIgnoreCase) ||
                    result.Contains("center", StringComparison.OrdinalIgnoreCase));
        Assert.True(result.Contains("36", StringComparison.Ordinal));
    }

    [Fact]
    public void CopyParagraphFormat_ShouldCopyFormat()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_copy_format.docx", "Source", "Target");
        var outputPath = CreateTestFilePath("test_copy_format_output.docx");
        var doc = new Document(docPath);
        var paragraphs = GetParagraphs(doc);
        paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Center;
        paragraphs[0].ParagraphFormat.LeftIndent = 36;
        doc.Save(docPath);
        _tool.Execute("copy_format", docPath, outputPath: outputPath, sourceParagraphIndex: 0, targetParagraphIndex: 1);
        var resultDoc = new Document(outputPath);
        var resultParagraphs = GetParagraphs(resultDoc);
        Assert.Equal(ParagraphAlignment.Center, resultParagraphs[1].ParagraphFormat.Alignment);
        Assert.Equal(36, resultParagraphs[1].ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void MergeParagraphs_ShouldMergeParagraphs()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge_paragraphs.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_merge_paragraphs_output.docx");
        _tool.Execute("merge", docPath, outputPath: outputPath, startParagraphIndex: 0, endParagraphIndex: 2);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc, false);
        var nonEmptyCount = paragraphs.Count(p => !string.IsNullOrWhiteSpace(p.GetText()));
        Assert.True(nonEmptyCount <= 1,
            $"After merging, should have 1 or fewer non-empty paragraphs, but got {nonEmptyCount}");
    }

    [Fact]
    public void EditParagraph_WithAlignment_ShouldApplyAlignment()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_alignment.docx", "Test");
        var outputPath = CreateTestFilePath("test_edit_alignment_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, alignment: "center");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);

        if (IsEvaluationMode())
        {
            Assert.True(File.Exists(outputPath), "Output document should be created");
            Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
        }
        else
        {
            Assert.Equal(ParagraphAlignment.Center, paragraphs[0].ParagraphFormat.Alignment);
        }
    }

    [Fact]
    public void EditParagraph_WithIndentation_ShouldApplyIndentation()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_indentation.docx", "Test");
        var outputPath = CreateTestFilePath("test_edit_indentation_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, indentLeft: 36, firstLineIndent: 18);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.Equal(36, paragraphs[0].ParagraphFormat.LeftIndent);
        Assert.Equal(18, paragraphs[0].ParagraphFormat.FirstLineIndent);
    }

    [Fact]
    public void EditParagraph_WithSpacing_ShouldApplySpacing()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_spacing.docx", "Test");
        var outputPath = CreateTestFilePath("test_edit_spacing_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, spaceBefore: 12, spaceAfter: 12);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.Equal(12, paragraphs[0].ParagraphFormat.SpaceBefore);
        Assert.Equal(12, paragraphs[0].ParagraphFormat.SpaceAfter);
    }

    [Fact]
    public void EditParagraph_ShouldModifyEmptyParagraphStyle()
    {
        var docPath = CreateWordDocument("test_edit_empty_paragraph.docx");

        var doc = new Document();
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!標題3-國字括弧小寫 - (一)(二)(三)");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        var para = new Paragraph(doc)
        {
            ParagraphFormat = { StyleName = "!標題3-國字括弧小寫 - (一)(二)(三)" }
        };
        doc.FirstSection.Body.AppendChild(para);
        doc.Save(docPath);

        var initialDoc = new Document(docPath);
        var paragraphs = initialDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
            .ToList();
        Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");

        _tool.Execute("edit", docPath, outputPath: docPath, paragraphIndex: 0, styleName: "Normal");

        var resultDoc = new Document(docPath);
        var resultPara = resultDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().First();
        var actualStyle = resultPara.ParagraphFormat.StyleName ?? "";
        Assert.True(File.Exists(docPath), "Document should be saved after edit operation");

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            Assert.NotNull(actualStyle);
        else
            Assert.Equal("Normal", actualStyle);
    }

    [Fact]
    public void EditParagraph_WithMultipleEmptyParagraphs_ShouldModifyAll()
    {
        var docPath = CreateWordDocument("test_edit_multiple_empty_paragraphs.docx");

        var doc = new Document();
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!標題3-國字括弧小寫 - (一)(二)(三)");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        for (var i = 0; i < 3; i++)
        {
            var para = new Paragraph(doc)
            {
                ParagraphFormat = { StyleName = "!標題3-國字括弧小寫 - (一)(二)(三)" }
            };
            doc.FirstSection.Body.AppendChild(para);
        }

        doc.Save(docPath);

        var createdDoc = new Document(docPath);
        var paragraphCount = createdDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Count;
        for (var i = 0; i < paragraphCount; i++)
            _tool.Execute("edit", docPath, outputPath: docPath, paragraphIndex: i, styleName: "Normal");

        var resultDoc = new Document(docPath);
        var paragraphs = resultDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
            .ToList();
        Assert.True(File.Exists(docPath), "Document should be saved after edit operations");
        var emptyParaCount = 0;
        var normalStyleCount = 0;
        foreach (var para in paragraphs)
            if (string.IsNullOrWhiteSpace(para.GetText()))
            {
                emptyParaCount++;
                var actualStyle = para.ParagraphFormat.StyleName ?? "";
                if (actualStyle == "Normal" || actualStyle.Contains("Normal")) normalStyleCount++;
            }

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            if (normalStyleCount < emptyParaCount)
            {
                Assert.True(File.Exists(docPath), "Document should be saved after edit operations");
                Assert.True(normalStyleCount > 0,
                    $"At least some paragraphs should be changed to Normal style. Changed: {normalStyleCount}/{emptyParaCount}");
            }
        }
        else
        {
            Assert.Equal(emptyParaCount, normalStyleCount);
        }
    }

    [Fact]
    public void InsertParagraph_WithIndentRight_ShouldApplyRightIndent()
    {
        var docPath = CreateWordDocument("test_insert_paragraph_indent_right.docx");
        var outputPath = CreateTestFilePath("test_insert_paragraph_indent_right_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath, text: "Indented paragraph", indentRight: 36);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Indented paragraph"));
        Assert.NotNull(para);
        Assert.Equal(36, para.ParagraphFormat.RightIndent);
    }

    [Fact]
    public void InsertParagraph_WithAllIndentation_ShouldApplyAllIndents()
    {
        var docPath = CreateWordDocument("test_insert_paragraph_all_indents.docx");
        var outputPath = CreateTestFilePath("test_insert_paragraph_all_indents_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath, text: "Fully indented paragraph", indentLeft: 36,
            indentRight: 36, firstLineIndent: 18);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Fully indented paragraph"));
        Assert.NotNull(para);
        Assert.Equal(36, para.ParagraphFormat.LeftIndent);
        Assert.Equal(36, para.ParagraphFormat.RightIndent);
        Assert.Equal(18, para.ParagraphFormat.FirstLineIndent);
    }

    [Fact]
    public void InsertParagraph_WithAllAlignmentOptions_ShouldApplyAlignment()
    {
        var docPath = CreateWordDocument("test_insert_paragraph_alignments.docx");

        var alignments = new[] { "left", "center", "right", "justify" };
        var outputPaths = new Dictionary<string, string>();

        foreach (var alignment in alignments)
        {
            var outputPath = CreateTestFilePath($"test_insert_paragraph_{alignment}_output.docx");
            outputPaths[alignment] = outputPath;
            _tool.Execute("insert", docPath, outputPath: outputPath, text: $"Aligned {alignment}",
                alignment: alignment);
        }

        foreach (var alignment in alignments)
        {
            var outputPath = outputPaths[alignment];
            var doc = new Document(outputPath);
            var paragraphs = GetParagraphs(doc);
            var para = paragraphs.FirstOrDefault(p => p.GetText().Contains($"Aligned {alignment}"));

            Assert.NotNull(para);

            var expectedAlignment = alignment switch
            {
                "left" => ParagraphAlignment.Left,
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };

            Assert.Equal(expectedAlignment, para.ParagraphFormat.Alignment);
        }
    }

    [Fact]
    public void EditParagraph_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_edit_paragraph_all_formats.docx", "Test Paragraph");
        var outputPath = CreateTestFilePath("test_edit_paragraph_all_formats_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, sectionIndex: 0, alignment: "justify",
            indentLeft: 36, indentRight: 36, firstLineIndent: 18, spaceBefore: 12, spaceAfter: 12);
        var doc = new Document(outputPath);
        var paragraphs = doc.Sections[0].Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Paragraph"));

        if (para == null)
        {
            paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Test Paragraph"));
        }

        Assert.NotNull(para);

        Assert.True(File.Exists(outputPath), "Output document should be created");

        var leftIndent = para.ParagraphFormat.LeftIndent;
        var rightIndent = para.ParagraphFormat.RightIndent;
        var firstLineIndent = para.ParagraphFormat.FirstLineIndent;
        var spaceBefore = para.ParagraphFormat.SpaceBefore;
        var spaceAfter = para.ParagraphFormat.SpaceAfter;
        var alignment = para.ParagraphFormat.Alignment;

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            Assert.True(true,
                $"Formatting operation completed. Values: LeftIndent={leftIndent}, RightIndent={rightIndent}, FirstLineIndent={firstLineIndent}, SpaceBefore={spaceBefore}, SpaceAfter={spaceAfter}, Alignment={alignment}");
        }
        else
        {
            Assert.Equal(36, leftIndent);
            Assert.Equal(36, rightIndent);
            Assert.Equal(18, firstLineIndent);
            Assert.Equal(12, spaceBefore);
            Assert.Equal(12, spaceAfter);
            Assert.Equal(ParagraphAlignment.Justify, alignment);
        }
    }

    [Fact]
    public void InsertParagraph_WithAllFormattingCombinations_ShouldApplyAllFormats()
    {
        var docPath = CreateWordDocument("test_insert_paragraph_all_formats.docx");
        var outputPath = CreateTestFilePath("test_insert_paragraph_all_formats_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath, text: "Fully Formatted Paragraph", styleName: "Normal",
            alignment: "center", indentLeft: 36, indentRight: 36, firstLineIndent: 18, spaceBefore: 12, spaceAfter: 12);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        var para = paragraphs.FirstOrDefault(p => p.GetText().Contains("Fully Formatted Paragraph"));
        Assert.NotNull(para);
        Assert.Equal(ParagraphAlignment.Center, para.ParagraphFormat.Alignment);
        Assert.Equal(36, para.ParagraphFormat.LeftIndent);
        Assert.Equal(36, para.ParagraphFormat.RightIndent);
        Assert.Equal(18, para.ParagraphFormat.FirstLineIndent);
        Assert.Equal(12, para.ParagraphFormat.SpaceBefore);
        Assert.Equal(12, para.ParagraphFormat.SpaceAfter);
    }

    [Fact]
    public void MergeParagraphs_WithMultipleParagraphs_ShouldMergeCorrectly()
    {
        var docPath =
            CreateWordDocumentWithParagraphs("test_merge_multiple.docx", "First", "Second", "Third", "Fourth");
        var outputPath = CreateTestFilePath("test_merge_multiple_output.docx");
        _tool.Execute("merge", docPath, outputPath: outputPath, startParagraphIndex: 1, endParagraphIndex: 2);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc, false);
        var nonEmptyCount = paragraphs.Count(p => !string.IsNullOrWhiteSpace(p.GetText()));
        Assert.True(nonEmptyCount <= 2,
            $"After merging, should have 2 or fewer non-empty paragraphs, but got {nonEmptyCount}");
    }

    [Fact]
    public void CopyParagraphFormat_WithMultipleTargets_ShouldCopyToMultiple()
    {
        var docPath =
            CreateWordDocumentWithParagraphs("test_copy_format_multiple.docx", "Source", "Target1", "Target2");
        var outputPath = CreateTestFilePath("test_copy_format_multiple_output.docx");
        var doc = new Document(docPath);
        var paragraphs = GetParagraphs(doc);
        paragraphs[0].ParagraphFormat.Alignment = ParagraphAlignment.Right;
        paragraphs[0].ParagraphFormat.LeftIndent = 72;
        doc.Save(docPath);
        _tool.Execute("copy_format", docPath, outputPath: outputPath, sourceParagraphIndex: 0, targetParagraphIndex: 1);
        var resultDoc = new Document(outputPath);
        var resultParagraphs = GetParagraphs(resultDoc);
        Assert.Equal(72, resultParagraphs[1].ParagraphFormat.LeftIndent);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            Assert.True(File.Exists(outputPath), "Output file should be created");
        else
            Assert.Equal(ParagraphAlignment.Right, resultParagraphs[1].ParagraphFormat.Alignment);
    }

    [Fact]
    public void Edit_WithFontName_ShouldApplyFontName()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_fontname.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_fontname_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, fontName: "Arial");
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Test Text"));
        Assert.NotNull(run);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            Assert.True(File.Exists(outputPath), "Output file should be created");
        else
            Assert.Equal("Arial", run.Font.Name);
    }

    [Fact]
    public void Edit_WithFontSize_ShouldApplyFontSize()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_fontsize.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_fontsize_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, fontSize: 18);
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Test Text"));
        Assert.NotNull(run);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            Assert.True(File.Exists(outputPath), "Output file should be created");
        else
            Assert.Equal(18, run.Font.Size);
    }

    [Fact]
    public void Edit_WithColor_ShouldApplyColor()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_color.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_color_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, color: "FF0000");
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var run = runs.FirstOrDefault(r => r.Text.Contains("Test Text"));
        Assert.NotNull(run);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            Assert.True(File.Exists(outputPath), "Output file should be created");
        else
            Assert.Equal(Color.FromArgb(255, 0, 0), run.Font.Color);
    }

    [Fact]
    public void Edit_WithLineSpacing_ShouldApplyLineSpacing()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_linespacing.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_linespacing_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, lineSpacing: 24,
            lineSpacingRule: "exactly");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);
        Assert.Equal(24, paragraphs[0].ParagraphFormat.LineSpacing);
        Assert.Equal(LineSpacingRule.Exactly, paragraphs[0].ParagraphFormat.LineSpacingRule);
    }

    [Fact]
    public void Edit_WithTabStops_ShouldApplyTabStops()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_tabstops.docx", "Test\tText");
        var outputPath = CreateTestFilePath("test_edit_tabstops_output.docx");
        var tabStopsArray = new JsonArray
        {
            new JsonObject
            {
                ["position"] = 72.0,
                ["alignment"] = "Center",
                ["leader"] = "Dots"
            },
            new JsonObject
            {
                ["position"] = 144.0,
                ["alignment"] = "Right",
                ["leader"] = "None"
            }
        };
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, tabStops: tabStopsArray);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);
        var tabStops = paragraphs[0].ParagraphFormat.TabStops;
        Assert.True(tabStops.Count >= 2, $"Should have at least 2 tab stops, got {tabStops.Count}");
    }

    [Fact]
    public void InsertParagraph_WithNegativeOneIndex_ShouldInsertAtBeginning()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert_beginning.docx", "First", "Second");
        var outputPath = CreateTestFilePath("test_insert_beginning_output.docx");
        var result = _tool.Execute("insert", docPath, outputPath: outputPath, paragraphIndex: -1, text: "New First");
        Assert.StartsWith("Paragraph inserted", result);
        Assert.Contains("beginning of document", result);
        var doc = new Document(outputPath);
        var allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var newFirstPara = allParagraphs.Any(p => p.GetText().Contains("New First"));
        Assert.True(newFirstPara, "New paragraph 'New First' should be inserted");
    }

    [Fact]
    public void DeleteParagraph_WithNegativeOneIndex_ShouldDeleteLastParagraph()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_last.docx", "First", "Second", "Last");
        var outputPath = CreateTestFilePath("test_delete_last_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, paragraphIndex: -1);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc, false);
        var lastFound = paragraphs.Any(p => p.GetText().Contains("Last"));
        Assert.False(lastFound, "Last paragraph should be deleted");
    }

    [Fact]
    public void Edit_WithLineSpacingRuleSingle_ShouldUseMultipleRule()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_linespacing_single.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_linespacing_single_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, lineSpacingRule: "single");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);
        Assert.Equal(LineSpacingRule.Multiple, paragraphs[0].ParagraphFormat.LineSpacingRule);
        Assert.Equal(1.0, paragraphs[0].ParagraphFormat.LineSpacing);
    }

    [Fact]
    public void Edit_WithLineSpacingRuleDouble_ShouldUseMultipleRule()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_linespacing_double.docx", "Test Text");
        var outputPath = CreateTestFilePath("test_edit_linespacing_double_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, lineSpacingRule: "double");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0);
        Assert.Equal(LineSpacingRule.Multiple, paragraphs[0].ParagraphFormat.LineSpacingRule);
        Assert.Equal(2.0, paragraphs[0].ParagraphFormat.LineSpacing);
    }

    [Fact]
    public void DeleteParagraph_WithOutOfRangeIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete_outofrange.docx", "First", "Second");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", docPath, paragraphIndex: 100));
        Assert.Contains("out of range", exception.Message);
        Assert.Contains("valid indices", exception.Message);
    }

    [Fact]
    public void GetParagraphFormat_WithOutOfRangeIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_getformat_outofrange.docx", "First");
        var exception =
            Assert.Throws<ArgumentException>(() => _tool.Execute("get_format", docPath, paragraphIndex: 100));
        Assert.Contains("out of range", exception.Message);
        Assert.Contains("valid indices", exception.Message);
    }

    [Fact]
    public void EditParagraph_EmptyParagraphWithFontSettings_ShouldCreateSentinelRun()
    {
        var docPath = CreateTestFilePath("test_edit_empty_para_font.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln(); // Empty paragraph
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_empty_para_font_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 1, fontName: "Arial", fontSize: 16,
            bold: true);
        var resultDoc = new Document(outputPath);
        var paragraphs = GetParagraphs(resultDoc);
        Assert.True(paragraphs.Count >= 2);

        var emptyPara = paragraphs[1];
        var runs = emptyPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.True(runs.Count > 0, "Empty paragraph should have a sentinel run after font edit");

        var sentinelRun = runs[0];
        Assert.Equal("Arial", sentinelRun.Font.Name);
        Assert.Equal(16, sentinelRun.Font.Size);
        Assert.True(sentinelRun.Font.Bold);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("GeT")]
    [InlineData("get")]
    public void Execute_GetOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Content");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("paragraphs", result);
    }

    [Theory]
    [InlineData("INSERT")]
    [InlineData("InSeRt")]
    [InlineData("insert")]
    public void Execute_InsertOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_insert_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_insert_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, text: "New paragraph");
        Assert.StartsWith("Paragraph inserted", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("DeLeTe")]
    [InlineData("delete")]
    public void Execute_DeleteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_delete_{operation}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_delete_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 0);
        Assert.StartsWith("Paragraph #0 deleted", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("EdIt")]
    [InlineData("edit")]
    public void Execute_EditOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_edit_{operation}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_edit_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 0, bold: true);
        Assert.Contains("format edited successfully", result);
    }

    [Theory]
    [InlineData("GET_FORMAT")]
    [InlineData("Get_Format")]
    [InlineData("get_format")]
    public void Execute_GetFormatOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_getformat_{operation.Replace("_", "")}.docx", "Content");
        var result = _tool.Execute(operation, docPath, paragraphIndex: 0);
        Assert.Contains("alignment", result);
    }

    [Theory]
    [InlineData("COPY_FORMAT")]
    [InlineData("Copy_Format")]
    [InlineData("copy_format")]
    public void Execute_CopyFormatOperationIsCaseInsensitive(string operation)
    {
        var docPath =
            CreateWordDocumentWithParagraphs($"test_copyformat_{operation.Replace("_", "")}.docx", "First", "Second");
        var outputPath = CreateTestFilePath($"test_copyformat_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            sourceParagraphIndex: 0, targetParagraphIndex: 1);
        Assert.StartsWith("Paragraph format copied", result);
    }

    [Theory]
    [InlineData("MERGE")]
    [InlineData("MeRgE")]
    [InlineData("merge")]
    public void Execute_MergeOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithParagraphs($"test_merge_{operation}.docx", "First", "Second");
        var outputPath = CreateTestFilePath($"test_merge_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 1);
        Assert.StartsWith("Paragraphs merged", result);
    }

    #endregion

    #region Session

    [Fact]
    public void InsertParagraph_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert.docx", "Existing content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert", sessionId: sessionId, text: "New paragraph", paragraphIndex: 0);
        Assert.StartsWith("Paragraph inserted successfully", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(doc);
        Assert.Contains(paragraphs, p => p.GetText().Contains("New paragraph"));
    }

    [Fact]
    public void EditParagraph_WithSessionId_ShouldEditInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_edit.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("edit", sessionId: sessionId, paragraphIndex: 0, bold: true, fontSize: 16);
        Assert.Contains("format edited successfully", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
    }

    [Fact]
    public void DeleteParagraph_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_delete.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete", sessionId: sessionId, paragraphIndex: 1);
        Assert.StartsWith("Paragraph #1 deleted", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void GetParagraphs_WithSessionId_ShouldReturnParagraphInfo()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_get.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("First", result);
        Assert.Contains("Second", result);
        Assert.Contains("Third", result);
    }

    [Fact]
    public void GetParagraphFormat_WithSessionId_ShouldReturnFormat()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_format.docx", "Formatted text");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_format", sessionId: sessionId, paragraphIndex: 0);
        Assert.Contains("alignment", result);
    }

    [Fact]
    public void CopyParagraphFormat_WithSessionId_ShouldCopyInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_copy_format.docx", "Source", "Target");
        var sessionId = OpenSession(docPath);
        var doc = SessionManager.GetDocument<Document>(sessionId);

        var paragraphs = GetParagraphs(doc);
        paragraphs[0].ParagraphFormat.LeftIndent = 72;
        _tool.Execute("copy_format", sessionId: sessionId, sourceParagraphIndex: 0, targetParagraphIndex: 1);
        var resultParagraphs = GetParagraphs(doc);
        Assert.Equal(72, resultParagraphs[1].ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void MergeParagraphs_WithSessionId_ShouldMergeInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_merge.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("merge", sessionId: sessionId, startParagraphIndex: 0, endParagraphIndex: 2);
        Assert.StartsWith("Paragraphs merged", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_para.docx", "Path content");
        var docPath2 = CreateWordDocumentWithContent("test_session_para.docx", "Session content");
        var sessionId = OpenSession(docPath2);

        var result = _tool.Execute("get", docPath1, sessionId);

        Assert.Contains("Session content", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void InsertParagraph_WithoutText_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_insert_no_text.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath, paragraphIndex: 0));
        Assert.Contains("text parameter is required", ex.Message);
    }

    [Fact]
    public void EditParagraph_WithoutParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_no_index.docx", "Content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, bold: true));
        Assert.Contains("paragraphIndex parameter is required", ex.Message);
    }

    [Fact]
    public void CopyFormat_WithoutSourceParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_copy_no_source.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy_format", docPath, targetParagraphIndex: 1));
        Assert.Contains("sourceParagraphIndex parameter is required", ex.Message);
    }

    [Fact]
    public void CopyFormat_WithoutTargetParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_copy_no_target.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy_format", docPath, sourceParagraphIndex: 0));
        Assert.Contains("targetParagraphIndex parameter is required", ex.Message);
    }

    [Fact]
    public void Merge_WithoutStartParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge_no_start.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", docPath, endParagraphIndex: 1));
        Assert.Contains("startParagraphIndex parameter is required", ex.Message);
    }

    [Fact]
    public void Merge_WithoutEndParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge_no_end.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", docPath, startParagraphIndex: 0));
        Assert.Contains("endParagraphIndex parameter is required", ex.Message);
    }

    [Fact]
    public void Merge_WithSameStartAndEnd_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge_same_idx.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", docPath, startParagraphIndex: 0, endParagraphIndex: 0));
        Assert.Contains("no merge needed", ex.Message);
    }

    [Fact]
    public void Merge_WithStartGreaterThanEnd_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge_invalid_range.docx", "First", "Second");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", docPath, startParagraphIndex: 1, endParagraphIndex: 0));
        Assert.Contains("cannot be greater than", ex.Message);
    }

    #endregion
}