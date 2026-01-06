using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordStyleToolTests : WordTestBase
{
    private readonly WordStyleTool _tool;

    public WordStyleToolTests()
    {
        _tool = new WordStyleTool(SessionManager);
    }

    #region General

    [Fact]
    public void GetStyles_ShouldReturnAllStyles()
    {
        var docPath = CreateWordDocument("test_get_styles.docx");
        var result = _tool.Execute("get_styles", docPath);
        Assert.Contains("\"count\":", result);
        Assert.Contains("\"paragraphStyles\":", result);
        Assert.Contains("Normal", result);
    }

    [Fact]
    public void GetStyles_WithIncludeBuiltIn_ShouldIncludeBuiltInStyles()
    {
        var docPath = CreateWordDocument("test_get_styles_builtin.docx");
        var result = _tool.Execute("get_styles", docPath, includeBuiltIn: true);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("\"count\":", result);
        Assert.Contains("\"paragraphStyles\":", result);
        Assert.Contains("Normal", result);
    }

    [Fact]
    public void GetStyles_ShouldReturnJsonFormat()
    {
        var docPath = CreateWordDocument("test_get_styles_json.docx");
        var result = _tool.Execute("get_styles", docPath);
        Assert.Contains("\"count\"", result);
        Assert.Contains("\"paragraphStyles\"", result);
    }

    [Fact]
    public void CreateStyle_ShouldCreateNewStyle()
    {
        var docPath = CreateWordDocument("test_create_style.docx");
        var outputPath = CreateTestFilePath("test_create_style_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: "CustomStyle", styleType: "paragraph", fontSize: 14, bold: true);
        var doc = new Document(outputPath);
        var style = doc.Styles["CustomStyle"];
        Assert.NotNull(style);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.Bold);
    }

    [Fact]
    public void CreateStyle_WithBaseStyle_ShouldInheritFromBase()
    {
        var docPath = CreateWordDocument("test_create_style_base.docx");
        var outputPath = CreateTestFilePath("test_create_style_base_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: "CustomHeading", styleType: "paragraph", baseStyle: "Heading 1", fontSize: 18);
        var doc = new Document(outputPath);
        var style = doc.Styles["CustomHeading"];
        Assert.NotNull(style);
        Assert.Equal("Heading 1", style.BaseStyleName);
    }

    [Fact]
    public void CreateStyle_WithAllFontOptions_ShouldCreateStyleWithAllFonts()
    {
        var docPath = CreateWordDocument("test_create_style_all_fonts.docx");
        var outputPath = CreateTestFilePath("test_create_style_all_fonts_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: "MultiFontStyle", styleType: "paragraph",
            fontName: "Arial", fontNameAscii: "Times New Roman", fontNameFarEast: "Microsoft YaHei", fontSize: 12);
        var doc = new Document(outputPath);
        var style = doc.Styles["MultiFontStyle"];
        Assert.NotNull(style);
        Assert.Equal("Times New Roman", style.Font.NameAscii);
        Assert.Equal("Microsoft YaHei", style.Font.NameFarEast);
    }

    [Fact]
    public void CreateStyle_WithAllFormattingOptions_ShouldCreateCompleteStyle()
    {
        var docPath = CreateWordDocument("test_create_style_complete.docx");
        var outputPath = CreateTestFilePath("test_create_style_complete_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: "CompleteStyle", styleType: "paragraph",
            fontName: "Arial", fontSize: 14, bold: true, italic: true, underline: true,
            color: "FF0000", alignment: "center", spaceBefore: 12, spaceAfter: 12, lineSpacing: 1.5);
        var doc = new Document(outputPath);
        var style = doc.Styles["CompleteStyle"];
        Assert.NotNull(style);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.Bold);
        Assert.True(style.Font.Italic);
        Assert.Equal(ParagraphAlignment.Center, style.ParagraphFormat.Alignment);
        Assert.Equal(12, style.ParagraphFormat.SpaceBefore);
        Assert.Equal(12, style.ParagraphFormat.SpaceAfter);
    }

    [Theory]
    [InlineData("character", StyleType.Character)]
    [InlineData("table", StyleType.Table)]
    [InlineData("list", StyleType.List)]
    [InlineData("paragraph", StyleType.Paragraph)]
    public void CreateStyle_WithStyleTypes_ShouldCreateCorrectType(string styleType, StyleType expectedType)
    {
        var docPath = CreateWordDocument($"test_create_style_{styleType}.docx");
        var outputPath = CreateTestFilePath($"test_create_style_{styleType}_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: $"{styleType}Style", styleType: styleType, fontSize: 12);
        var doc = new Document(outputPath);
        var style = doc.Styles[$"{styleType}Style"];
        Assert.NotNull(style);
        Assert.Equal(expectedType, style.Type);
    }

    [Fact]
    public void CreateStyle_WithoutStyleType_ShouldDefaultToParagraph()
    {
        var docPath = CreateWordDocument("test_create_no_type.docx");
        var outputPath = CreateTestFilePath("test_create_no_type_output.docx");

        var result = _tool.Execute("create_style", docPath, outputPath: outputPath, styleName: "TestStyle");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Style 'TestStyle' created", result);

        var doc = new Document(outputPath);
        var style = doc.Styles["TestStyle"];
        Assert.NotNull(style);
        Assert.Equal(StyleType.Paragraph, style.Type);
    }

    [Fact]
    public void ApplyStyle_ToSingleParagraph_ShouldApplyStyle()
    {
        var docPath = CreateWordDocumentWithContent("test_apply_style_single.docx", "Test");
        var outputPath = CreateTestFilePath("test_apply_style_single_output.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "TestStyle");
        customStyle.Font.Size = 16;
        doc.Save(docPath);
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "TestStyle", paragraphIndex: 0);
        var resultDoc = new Document(outputPath);
        var paragraphs = GetParagraphs(resultDoc);

        if (IsEvaluationMode())
        {
            Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");
            Assert.True(File.Exists(outputPath), "Output file should be created");
            Assert.NotNull(resultDoc.Styles["TestStyle"]);
        }
        else
        {
            Assert.Equal("TestStyle", paragraphs[0].ParagraphFormat.StyleName);
        }
    }

    [Fact]
    public void ApplyStyle_ToMultipleParagraphs_ShouldApplyToAll()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_apply_style_multiple.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_apply_style_multiple_output.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "TestStyle");
        customStyle.Font.Size = 16;
        doc.Save(docPath);
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "TestStyle", paragraphIndices: [0, 1, 2]);
        var resultDoc = new Document(outputPath);
        var paragraphs = GetParagraphs(resultDoc);
        foreach (var para in paragraphs.Take(3)) Assert.Equal("TestStyle", para.ParagraphFormat.StyleName);
    }

    [Fact]
    public void ApplyStyle_ToTable_ShouldApplyTableStyle()
    {
        var docPath = CreateWordDocument("test_apply_style_table.docx");
        var outputPath = CreateTestFilePath("test_apply_style_table_output.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        var tableStyle = doc.Styles.Add(StyleType.Table, "TestTableStyle");
        tableStyle.Font.Size = 12;
        doc.Save(docPath);
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "TestTableStyle", tableIndex: 0);
        var resultDoc = new Document(outputPath);
        var tables = resultDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal("TestTableStyle", tables[0].StyleName);
    }

    [Fact]
    public void ApplyStyle_ShouldModifyEmptyParagraphStyle()
    {
        var docPath = CreateWordDocument("test_apply_style_empty_paragraph.docx");

        var doc = new Document();
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestStyle");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        var para = new Paragraph(doc)
        {
            ParagraphFormat = { StyleName = "!TestStyle" }
        };
        doc.FirstSection.Body.AppendChild(para);
        doc.Save(docPath);

        var initialDoc = new Document(docPath);
        var paragraphs = initialDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
            .ToList();
        Assert.True(paragraphs.Count > 0, "Document should have at least one paragraph");

        _tool.Execute("apply_style", docPath, outputPath: docPath,
            styleName: "Normal", paragraphIndex: 0);

        var resultDoc = new Document(docPath);
        var resultPara = resultDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().First();
        var actualStyle = resultPara.ParagraphFormat.StyleName ?? "";
        Assert.True(File.Exists(docPath), "Document should be saved after apply style operation");

        if (IsEvaluationMode())
        {
            Assert.NotNull(actualStyle);
            Assert.True(actualStyle.Length > 0, "Style name should not be empty");
            Assert.NotNull(resultDoc.Styles["Normal"]);
        }
        else
        {
            Assert.Equal("Normal", actualStyle);
        }
    }

    [Fact]
    public void ApplyStyle_WithMultipleEmptyParagraphs_ShouldModifyAll()
    {
        var docPath = CreateWordDocument("test_apply_style_multiple_empty_paragraphs.docx");

        var doc = new Document();
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "!TestStyle2");
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        for (var i = 0; i < 3; i++)
        {
            var para = new Paragraph(doc)
            {
                ParagraphFormat = { StyleName = "!TestStyle2" }
            };
            doc.FirstSection.Body.AppendChild(para);
        }

        doc.Save(docPath);

        _tool.Execute("apply_style", docPath, outputPath: docPath,
            styleName: "Normal", applyToAllParagraphs: true);

        var resultDoc = new Document(docPath);
        var paragraphs = resultDoc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
            .ToList();
        foreach (var para in paragraphs)
            if (string.IsNullOrWhiteSpace(para.GetText()))
            {
                var actualStyle = para.ParagraphFormat.StyleName ?? "";
                Assert.True(actualStyle == "Normal",
                    $"Empty paragraph should be changed to Normal style, but got: {actualStyle}");
            }
    }

    [Fact]
    public void ApplyStyle_WithSectionIndex_ShouldApplyToSpecificSection()
    {
        var docPath = CreateWordDocument("test_apply_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Section 0 Para");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 1 Para");
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "SectionStyle");
        customStyle.Font.Size = 18;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_apply_section_output.docx");
        _tool.Execute("apply_style", docPath, outputPath: outputPath,
            styleName: "SectionStyle", paragraphIndex: 0, sectionIndex: 1);
        var resultDoc = new Document(outputPath);
        var section1Paras = resultDoc.Sections[1].Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>()
            .ToList();
        Assert.Equal("SectionStyle", section1Paras[0].ParagraphFormat.StyleName);
    }

    [Fact]
    public void CopyStyles_ShouldCopyStylesFromSource()
    {
        var sourcePath = CreateWordDocument("test_copy_styles_source.docx");
        var doc = new Document(sourcePath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "SourceStyle");
        customStyle.Font.Size = 16;
        doc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_styles_target.docx");
        var outputPath = CreateTestFilePath("test_copy_styles_output.docx");
        _tool.Execute("copy_styles", targetPath, outputPath: outputPath, sourceDocument: sourcePath);
        var resultDoc = new Document(outputPath);
        var copiedStyle = resultDoc.Styles["SourceStyle"];
        Assert.NotNull(copiedStyle);
    }

    [Fact]
    public void CopyStyles_WithStyleNames_ShouldCopyOnlySpecifiedStyles()
    {
        var sourcePath = CreateWordDocument("test_copy_specific_source.docx");
        var doc = new Document(sourcePath);
        doc.Styles.Add(StyleType.Paragraph, "StyleA");
        doc.Styles.Add(StyleType.Paragraph, "StyleB");
        doc.Styles.Add(StyleType.Paragraph, "StyleC");
        doc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_specific_target.docx");
        var outputPath = CreateTestFilePath("test_copy_specific_output.docx");
        var result = _tool.Execute("copy_styles", targetPath, outputPath: outputPath,
            sourceDocument: sourcePath, styleNames: ["StyleA", "StyleC"]);
        Assert.StartsWith("Copied 2 style(s)", result);
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc.Styles["StyleA"]);
        Assert.Null(resultDoc.Styles["StyleB"]);
        Assert.NotNull(resultDoc.Styles["StyleC"]);
    }

    [Fact]
    public void CopyStyles_WithOverwrite_ShouldOverwriteExistingStyles()
    {
        var sourcePath = CreateWordDocument("test_copy_overwrite_source.docx");
        var sourceDoc = new Document(sourcePath);
        var sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "SharedStyle");
        sourceStyle.Font.Size = 20;
        sourceDoc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_overwrite_target.docx");
        var targetDoc = new Document(targetPath);
        var targetStyle = targetDoc.Styles.Add(StyleType.Paragraph, "SharedStyle");
        targetStyle.Font.Size = 12;
        targetDoc.Save(targetPath);

        var outputPath = CreateTestFilePath("test_copy_overwrite_output.docx");
        _tool.Execute("copy_styles", targetPath, outputPath: outputPath,
            sourceDocument: sourcePath, styleNames: ["SharedStyle"], overwriteExisting: true);
        var resultDoc = new Document(outputPath);
        var resultStyle = resultDoc.Styles["SharedStyle"];
        Assert.NotNull(resultStyle);
        Assert.Equal(20, resultStyle.Font.Size);
    }

    [Fact]
    public void CopyStyles_WithoutOverwrite_ShouldSkipExistingStyles()
    {
        var sourcePath = CreateWordDocument("test_copy_no_overwrite_source.docx");
        var sourceDoc = new Document(sourcePath);
        var sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "SharedStyle");
        sourceStyle.Font.Size = 20;
        sourceDoc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_no_overwrite_target.docx");
        var targetDoc = new Document(targetPath);
        var targetStyle = targetDoc.Styles.Add(StyleType.Paragraph, "SharedStyle");
        targetStyle.Font.Size = 12;
        targetDoc.Save(targetPath);

        var outputPath = CreateTestFilePath("test_copy_no_overwrite_output.docx");
        var result = _tool.Execute("copy_styles", targetPath, outputPath: outputPath,
            sourceDocument: sourcePath, styleNames: ["SharedStyle"], overwriteExisting: false);
        Assert.Contains("Skipped: 1", result); // Meaningful check for skipped styles
        var resultDoc = new Document(outputPath);
        var resultStyle = resultDoc.Styles["SharedStyle"];
        Assert.NotNull(resultStyle);
        Assert.Equal(12, resultStyle.Font.Size);
    }

    [Theory]
    [InlineData("GET_STYLES")]
    [InlineData("GeT_sTyLeS")]
    [InlineData("get_styles")]
    public void Execute_OperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation.GetHashCode()}_case.docx");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("Styles", result);
    }

    [Theory]
    [InlineData("CREATE_STYLE")]
    [InlineData("Create_Style")]
    [InlineData("create_style")]
    public void Execute_CreateStyleOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation.GetHashCode()}_create_case.docx");
        var outputPath = CreateTestFilePath($"test_{operation.GetHashCode()}_create_case_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            styleName: $"Style{operation.GetHashCode()}", styleType: "paragraph");
        Assert.StartsWith("Style '", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("PARAGRAPH", StyleType.Paragraph)]
    [InlineData("ChArAcTeR", StyleType.Character)]
    [InlineData("table", StyleType.Table)]
    public void CreateStyle_StyleTypeIsCaseInsensitive(string styleType, StyleType expectedType)
    {
        var docPath = CreateWordDocument($"test_create_{styleType.GetHashCode()}_type.docx");
        var outputPath = CreateTestFilePath($"test_create_{styleType.GetHashCode()}_type_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: $"TypeStyle{styleType.GetHashCode()}", styleType: styleType);
        var doc = new Document(outputPath);
        var style = doc.Styles[$"TypeStyle{styleType.GetHashCode()}"];
        Assert.NotNull(style);
        Assert.Equal(expectedType, style.Type);
    }

    [Theory]
    [InlineData("CENTER", ParagraphAlignment.Center)]
    [InlineData("Right", ParagraphAlignment.Right)]
    [InlineData("left", ParagraphAlignment.Left)]
    public void CreateStyle_AlignmentIsCaseInsensitive(string alignment, ParagraphAlignment expectedAlignment)
    {
        var docPath = CreateWordDocument($"test_create_{alignment.GetHashCode()}_align.docx");
        var outputPath = CreateTestFilePath($"test_create_{alignment.GetHashCode()}_align_output.docx");
        _tool.Execute("create_style", docPath, outputPath: outputPath,
            styleName: $"AlignStyle{alignment.GetHashCode()}", styleType: "paragraph", alignment: alignment);
        var doc = new Document(outputPath);
        var style = doc.Styles[$"AlignStyle{alignment.GetHashCode()}"];
        Assert.NotNull(style);
        Assert.Equal(expectedAlignment, style.ParagraphFormat.Alignment);
    }

    #endregion

    #region Exception

    [Theory]
    [InlineData("unknown_operation")]
    [InlineData("invalid")]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, docPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void CreateStyle_WithInvalidStyleName_ShouldThrowArgumentException(string? styleName)
    {
        var docPath = CreateWordDocument($"test_create_invalid_name_{styleName?.GetHashCode() ?? 0}.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create_style", docPath, styleName: styleName, styleType: "paragraph"));

        Assert.Contains("styleName is required", ex.Message);
    }

    [Fact]
    public void CreateStyle_DuplicateName_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_create_style_duplicate.docx");
        var doc = new Document(docPath);
        doc.Styles.Add(StyleType.Paragraph, "ExistingStyle");
        doc.Save(docPath);
        Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("create_style", docPath,
                styleName: "ExistingStyle", styleType: "paragraph"));
    }

    [Fact]
    public void CreateStyle_WithInvalidColor_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_create_style_invalid_color.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("create_style", docPath,
                styleName: "InvalidColorStyle", styleType: "paragraph", color: "not_a_color"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ApplyStyle_WithInvalidStyleName_ShouldThrowArgumentException(string? styleName)
    {
        var docPath =
            CreateWordDocumentWithContent($"test_apply_invalid_style_{styleName?.GetHashCode() ?? 0}.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath, styleName: styleName, paragraphIndex: 0));

        Assert.Contains("styleName is required", ex.Message);
    }

    [Fact]
    public void ApplyStyle_NonExistentStyleName_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_apply_nonexistent_style.docx", "Test");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath,
                styleName: "NonExistentStyle", paragraphIndex: 0));
    }

    [Fact]
    public void ApplyStyle_InvalidParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_apply_invalid_index.docx", "Test");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath,
                styleName: "Normal", paragraphIndex: 999));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-10)]
    public void ApplyStyle_WithNegativeParagraphIndex_ShouldThrowException(int index)
    {
        var docPath = CreateWordDocumentWithContent($"test_apply_neg_index_{index}.docx", "Test");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath, styleName: "Normal", paragraphIndex: index));
    }

    [Fact]
    public void ApplyStyle_InvalidTableIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_apply_invalid_table.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath,
                styleName: "Normal", tableIndex: 999));
    }

    [Fact]
    public void ApplyStyle_WithNegativeTableIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_apply_neg_table.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();
        doc.Save(docPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath, styleName: "Normal", tableIndex: -1));
    }

    [Fact]
    public void ApplyStyle_WithNegativeSectionIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_apply_neg_section.docx", "Test");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath, styleName: "Normal", paragraphIndex: 0, sectionIndex: -1));
    }

    [Fact]
    public void ApplyStyle_NoTargetSpecified_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_apply_no_target.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_style", docPath, styleName: "Normal"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void CopyStyles_WithInvalidSourceDocument_ShouldThrowArgumentException(string? sourceDoc)
    {
        var docPath = CreateWordDocument($"test_copy_invalid_source_{sourceDoc?.GetHashCode() ?? 0}.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy_styles", docPath, sourceDocument: sourceDoc));

        Assert.Contains("sourceDocument is required", ex.Message);
    }

    [Fact]
    public void CopyStyles_SourceNotFound_ShouldThrowException()
    {
        var targetPath = CreateWordDocument("test_copy_source_not_found.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("copy_styles", targetPath, sourceDocument: "nonexistent_file.docx"));
    }

    #endregion

    #region Session

    [Fact]
    public void GetStyles_WithSessionId_ShouldReturnStyles()
    {
        var docPath = CreateWordDocument("test_session_get_styles.docx");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "SessionStyle");
        customStyle.Font.Size = 16;
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_styles", sessionId: sessionId);
        Assert.Contains("SessionStyle", result);
        Assert.Contains("\"count\"", result);
    }

    [Fact]
    public void CreateStyle_WithSessionId_ShouldCreateStyleInMemory()
    {
        var docPath = CreateWordDocument("test_session_create_style.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("create_style", sessionId: sessionId,
            styleName: "SessionCreatedStyle", styleType: "paragraph", fontSize: 20, bold: true);
        Assert.Contains("SessionCreatedStyle", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var style = doc.Styles["SessionCreatedStyle"];
        Assert.NotNull(style);
        Assert.Equal(20, style.Font.Size);
        Assert.True(style.Font.Bold);
    }

    [Fact]
    public void ApplyStyle_WithSessionId_ShouldApplyStyleInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_apply_style.docx", "Test content");
        var doc = new Document(docPath);
        var customStyle = doc.Styles.Add(StyleType.Paragraph, "ApplySessionStyle");
        customStyle.Font.Size = 18;
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("apply_style", sessionId: sessionId,
            styleName: "ApplySessionStyle", paragraphIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = sessionDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.Equal("ApplySessionStyle", paragraphs[0].ParagraphFormat.StyleName);
    }

    [Fact]
    public void CopyStyles_WithSessionId_ShouldCopyToMemory()
    {
        var sourcePath = CreateWordDocument("test_copy_session_source.docx");
        var sourceDoc = new Document(sourcePath);
        sourceDoc.Styles.Add(StyleType.Paragraph, "SourceSessionStyle");
        sourceDoc.Save(sourcePath);

        var targetPath = CreateWordDocument("test_copy_session_target.docx");
        var sessionId = OpenSession(targetPath);

        var result = _tool.Execute("copy_styles", sessionId: sessionId, sourceDocument: sourcePath);
        Assert.StartsWith("Copied", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc.Styles["SourceSessionStyle"]);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_styles", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_style_path.docx");
        var doc1 = new Document(docPath1);
        doc1.Styles.Add(StyleType.Paragraph, "PathStyle");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_style_session.docx");
        var doc2 = new Document(docPath2);
        doc2.Styles.Add(StyleType.Paragraph, "SessionStyle");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        var result = _tool.Execute("get_styles", docPath1, sessionId);

        Assert.Contains("SessionStyle", result);
        Assert.DoesNotContain("PathStyle", result);
    }

    #endregion
}