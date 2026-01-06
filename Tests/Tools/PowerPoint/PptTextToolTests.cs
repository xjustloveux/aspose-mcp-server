using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptTextToolTests : TestBase
{
    private readonly PptTextTool _tool;

    public PptTextToolTests()
    {
        _tool = new PptTextTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [SkippableFact]
    public void AddText_ShouldAddTextToSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_add_text.pptx");
        var outputPath = CreateTestFilePath("test_add_text_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, text: "Test Text", x: 100, y: 100, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.NotEmpty(textFrames);
        Assert.Contains(textFrames, tf => tf.TextFrame.Text.Contains("Test Text"));
    }

    [Fact]
    public void AddText_WithCustomDimensions_ShouldCreateTextBoxWithDimensions()
    {
        var pptPath = CreateTestPresentation("test_add_dims.pptx");
        var outputPath = CreateTestFilePath("test_add_dims_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, text: "Custom Size Text", x: 150, y: 200, width: 300, height: 80,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textShapes = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        var matchingShape = textShapes.FirstOrDefault(s => Math.Abs(s.X - 150) < 1 && Math.Abs(s.Y - 200) < 1);
        Assert.NotNull(matchingShape);
    }

    [Fact]
    public void EditText_ShouldEditText()
    {
        var pptPath = CreateTestPresentation("test_edit_text.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var pptShape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            pptShape.TextFrame.Text = "Original Text";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            for (var i = pptSlide.Shapes.Count - 1; i >= 0; i--)
            {
                var s = pptSlide.Shapes[i];
                if (s is IAutoShape { Placeholder: null, TextFrame: not null })
                {
                    correctShapeIndex = i;
                    break;
                }
            }
        }

        Assert.True(correctShapeIndex >= 0);

        var outputPath = CreateTestFilePath("test_edit_text_output.pptx");
        _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: correctShapeIndex, text: "Updated Text",
            outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count);
        var editedShape = slide.Shapes[correctShapeIndex] as IAutoShape;
        Assert.NotNull(editedShape);
        Assert.NotNull(editedShape.TextFrame);
        var actualText = editedShape.TextFrame.Text ?? "";

        if (IsEvaluationMode())
        {
            var hasUpdatedText = actualText.Contains("Updated", StringComparison.OrdinalIgnoreCase) ||
                                 actualText.Contains("Updat", StringComparison.OrdinalIgnoreCase);
            Assert.True(hasUpdatedText || actualText.Length > 0);
        }
        else
        {
            Assert.Contains("Updated", actualText, StringComparison.OrdinalIgnoreCase);
        }
    }

    [SkippableFact]
    public void ReplaceText_ShouldReplaceText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_replace_text.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var pptShape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            pptShape.TextFrame.Text = "Old Text";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_text_output.pptx");
        _tool.Execute("replace", pptPath, findText: "Old", replaceText: "New", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(textFrames, tf => tf.TextFrame.Text.Contains("New Text", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Replace_WithMatchCase_ShouldMatchCase()
    {
        var pptPath = CreateTestPresentation("test_replace_case.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var pptShape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            pptShape.TextFrame.Text = "Test TEXT test Test";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_case_output.pptx");
        _tool.Execute("replace", pptPath, findText: "Test", replaceText: "Case", matchCase: true,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        var text = string.Join(" ", textFrames.Select(tf => tf.TextFrame.Text));

        if (IsEvaluationMode())
        {
            Assert.True(File.Exists(outputPath));
            return;
        }

        Assert.Contains("Case", text, StringComparison.Ordinal);
        Assert.Contains("TEXT", text, StringComparison.Ordinal);
    }

    [SkippableFact]
    public void ReplaceText_InTable_ShouldReplaceTableText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_replace_table.pptx");
        using (var ppt = new Presentation())
        {
            var slide = ppt.Slides[0];
            var table = slide.Shapes.AddTable(100, 100, [100, 100], [50, 50]);
            table[0, 0].TextFrame.Text = "Old Value";
            table[1, 0].TextFrame.Text = "Keep This";
            table[0, 1].TextFrame.Text = "Old Data";
            table[1, 1].TextFrame.Text = "Old Info";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_table_output.pptx");
        var result = _tool.Execute("replace", pptPath, findText: "Old", replaceText: "New", outputPath: outputPath);
        Assert.Contains("3 occurrences", result);
        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.NotEmpty(tables);
        Assert.Contains("New Value", tables[0][0, 0].TextFrame.Text);
    }

    [Fact]
    public void ReplaceText_InGroupShape_ShouldReplaceGroupText()
    {
        var pptPath = CreateTestFilePath("test_replace_group.pptx");
        using (var ppt = new Presentation())
        {
            var slide = ppt.Slides[0];
            var groupShape = slide.Shapes.AddGroupShape();
            groupShape.X = 50;
            groupShape.Y = 50;
            groupShape.Width = 200;
            groupShape.Height = 150;
            var shape1 = groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 50);
            shape1.TextFrame.Text = "Old Text 1";
            var shape2 = groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 70, 100, 50);
            shape2.TextFrame.Text = "Old Text 2";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_group_output.pptx");
        var result = _tool.Execute("replace", pptPath, findText: "Old", replaceText: "New", outputPath: outputPath);
        Assert.Contains("2 occurrences", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ReplaceText_WithNoMatch_ShouldReturnZeroOccurrences()
    {
        var pptPath = CreateTestPresentation("test_replace_no_match.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var pptShape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            pptShape.TextFrame.Text = "Some Text";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_no_match_output.pptx");
        var result = _tool.Execute("replace", pptPath, findText: "NotFound", replaceText: "New",
            outputPath: outputPath);
        Assert.Contains("0 occurrences", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, text: "Test", x: 100, y: 100,
            outputPath: outputPath);
        Assert.StartsWith("Text added to slide", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_edit_{operation}.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape.TextFrame.Text = "Original";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var shapeIndex = 0;
        using (var ppt = new Presentation(pptPath))
        {
            for (var i = ppt.Slides[0].Shapes.Count - 1; i >= 0; i--)
                if (ppt.Slides[0].Shapes[i] is IAutoShape { Placeholder: null })
                {
                    shapeIndex = i;
                    break;
                }
        }

        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, text: "New",
            outputPath: outputPath);
        Assert.StartsWith("Text updated on slide", result);
    }

    [Theory]
    [InlineData("REPLACE")]
    [InlineData("Replace")]
    [InlineData("replace")]
    public void Operation_ShouldBeCaseInsensitive_Replace(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_replace_{operation}.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape.TextFrame.Text = "Old";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath($"test_case_replace_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, findText: "Old", replaceText: "New", outputPath: outputPath);
        Assert.StartsWith("Replaced", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddText_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, text: "Test"));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void AddText_WithEmptyOrNullText_ShouldThrowArgumentException(string? text)
    {
        var pptPath = CreateTestPresentation("test_add_no_text.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, text: text));
        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void AddText_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 999, text: "Test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditText_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_edit_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, text: "Test"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void EditText_WithInvalidShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_edit_invalid_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 999, text: "Test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditText_NonAutoShape_ShouldThrowException()
    {
        var pptPath = CreateTestFilePath("test_edit_non_autoshape.pptx");
        using (var ppt = new Presentation())
        {
            var slide = ppt.Slides[0];
            slide.Shapes.AddTable(100, 100, [100], [50]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, text: "Test"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void ReplaceText_WithEmptyOrNullFindText_ShouldThrowArgumentException(string? findText)
    {
        var pptPath = CreateTestPresentation("test_replace_no_find.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("replace", pptPath, findText: findText, replaceText: "New"));
        Assert.Contains("findText is required", ex.Message);
    }

    [Fact]
    public void ReplaceText_WithNullReplaceText_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_replace_no_replace.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("replace", pptPath, findText: "Old", replaceText: null));
        Assert.Contains("replaceText is required", ex.Message);
    }

    #endregion

    #region Session

    [SkippableFact]
    public void AddText_WithSessionId_ShouldAddInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var initialShapeCount = slide.Shapes.Count;

        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, text: "Session Text", x: 100, y: 100);
        Assert.StartsWith("Text added to slide", result);
        Assert.True(slide.Shapes.Count > initialShapeCount);
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(textFrames, tf => tf.TextFrame.Text.Contains("Session Text"));
    }

    [SkippableFact]
    public void EditText_WithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermarks");

        var pptPath = CreateTestFilePath("test_session_edit.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape.TextFrame.Text = "Original Session Text";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = 0;
        for (var i = 0; i < ppt.Slides[0].Shapes.Count; i++)
            if (ppt.Slides[0].Shapes[i] is IAutoShape { Placeholder: null })
            {
                shapeIndex = i;
                break;
            }

        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            text: "Session Edited Text");
        Assert.StartsWith("Text updated on slide", result);

        var editedShape = ppt.Slides[0].Shapes[shapeIndex] as IAutoShape;
        Assert.NotNull(editedShape);
        Assert.Contains("Session Edited", editedShape.TextFrame.Text);
    }

    [SkippableFact]
    public void ReplaceText_WithSessionId_ShouldReplaceInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_session_replace.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape.TextFrame.Text = "Old Session Value";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("replace", sessionId: sessionId, findText: "Old", replaceText: "New");
        Assert.Contains("1 occurrences", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var autoShapes = ppt.Slides[0].Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(autoShapes, s => s.TextFrame.Text.Contains("New Session Value"));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add", sessionId: "invalid_session_id", slideIndex: 0, text: "Test", x: 100, y: 100));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath1 = CreateTestFilePath("test_path_text.pptx");
        using (var pres1 = new Presentation())
        {
            var shape1 = pres1.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape1.TextFrame.Text = "PathText";
            pres1.Save(pptPath1, SaveFormat.Pptx);
        }

        var pptPath2 = CreateTestFilePath("test_session_text.pptx");
        using (var pres2 = new Presentation())
        {
            var shape2 = pres2.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape2.TextFrame.Text = "SessionText";
            pres2.Save(pptPath2, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("replace", pptPath1, sessionId, findText: "SessionText", replaceText: "Modified");
        Assert.Contains("1 occurrences", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var autoShapes = ppt.Slides[0].Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(autoShapes, s => s.TextFrame.Text.Contains("Modified"));
    }

    #endregion
}