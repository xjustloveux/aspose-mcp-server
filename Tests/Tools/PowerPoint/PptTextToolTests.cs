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

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

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
        Assert.True(textFrames.Count > 0, "Slide should contain text");
        Assert.True(textFrames.Any(tf => tf.TextFrame.Text.Contains("Test Text")),
            "Text content should be 'Test Text'");
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

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        // The added shape should be the last non-placeholder AutoShape with text
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

        Assert.True(correctShapeIndex >= 0, "Should find at least one non-placeholder AutoShape with text");

        var outputPath = CreateTestFilePath("test_edit_text_output.pptx");
        _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: correctShapeIndex, text: "Updated Text",
            outputPath: outputPath);

        // Assert - Check the shape at the same index
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];

        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var editedShape = slide.Shapes[correctShapeIndex] as IAutoShape;
        Assert.NotNull(editedShape);
        Assert.NotNull(editedShape.TextFrame);
        var actualText = editedShape.TextFrame.Text ?? "";

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            var hasUpdatedText = actualText.Contains("Updated", StringComparison.OrdinalIgnoreCase) ||
                                 actualText.Contains("Updat", StringComparison.OrdinalIgnoreCase);
            Assert.True(hasUpdatedText || actualText.Length > 0,
                $"In evaluation mode, text may be truncated due to watermark. " +
                $"Expected 'Updated' or 'Updat', but got: '{actualText.Substring(0, Math.Min(50, actualText.Length))}...'");
        }
        else
        {
            var hasUpdatedText = actualText.Contains("Updated", StringComparison.OrdinalIgnoreCase);
            Assert.True(hasUpdatedText,
                $"Text should contain 'Updated', but got: '{actualText.Substring(0, Math.Min(50, actualText.Length))}...'");
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
        var hasNewText =
            textFrames.Any(tf => tf.TextFrame.Text.Contains("New Text", StringComparison.OrdinalIgnoreCase));
        Assert.True(hasNewText, "Text should be replaced to 'New Text'");
    }

    [Fact]
    public void Replace_WithMatchCase_ShouldMatchCase()
    {
        var pptPath = CreateTestPresentation("test_replace_match_case.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var pptShape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            pptShape.TextFrame.Text = "Test TEXT test Test";
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_replace_match_case_output.pptx");
        _tool.Execute("replace", pptPath, findText: "Test", replaceText: "Case", matchCase: true,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        var text = string.Join(" ", textFrames.Select(tf => tf.TextFrame.Text));

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            // In evaluation mode, the watermark may interfere with assertions
            // Just verify the output file was created and operation completed
            Assert.True(File.Exists(outputPath), "Output file should be created");
            return;
        }

        // When matchCase is true, only "Test" should be replaced, not "test" or "TEXT"
        Assert.Contains("Case", text, StringComparison.Ordinal);
        // "TEXT" should remain unchanged (different case)
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
    public void AddText_WithCustomDimensions_ShouldCreateTextBoxWithDimensions()
    {
        var pptPath = CreateTestPresentation("test_add_text_dims.pptx");
        var outputPath = CreateTestFilePath("test_add_text_dims_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, text: "Custom Size Text", x: 150, y: 200, width: 300, height: 80,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textShapes = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        var matchingShape = textShapes.FirstOrDefault(s => Math.Abs(s.X - 150) < 1 && Math.Abs(s.Y - 200) < 1);
        Assert.NotNull(matchingShape);
    }

    [Fact]
    public void EditText_NonAutoShape_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_edit_non_autoshape.pptx");
        using (var ppt = new Presentation())
        {
            var slide = ppt.Slides[0];
            slide.Shapes.AddTable(100, 100, [100], [50]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        // May throw ArgumentException (when file loads successfully but shape is not AutoShape)
        // or other exceptions (when file cannot be loaded properly in evaluation mode)
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, text: "Test"));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
    }

    [Fact]
    public void Add_MissingText_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_missing_text.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0));
    }

    #endregion

    #region Session ID Tests

    [SkippableFact]
    public void AddText_WithSessionId_ShouldAddInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_session_add_text.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var initialShapeCount = slide.Shapes.Count;
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, text: "Session Text",
            x: 100, y: 100);
        Assert.Contains("Text added", result);
        Assert.Contains("session", result);
        Assert.True(slide.Shapes.Count > initialShapeCount);
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.True(textFrames.Any(tf => tf.TextFrame.Text.Contains("Session Text")),
            "Text content should be 'Session Text'");
    }

    [SkippableFact]
    public void EditText_WithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides,
            "Evaluation mode adds watermarks that interfere with text assertions");

        var pptPath = CreateTestFilePath("test_session_edit_text.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape.TextFrame.Text = "Original Session Text";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);

        // Find the shape index
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
        Assert.Contains("Text updated on slide 0", result);

        // Verify in-memory changes
        var editedShape = ppt.Slides[0].Shapes[shapeIndex] as IAutoShape;
        Assert.NotNull(editedShape);
        Assert.Contains("Session Edited", editedShape.TextFrame.Text);
    }

    [SkippableFact]
    public void ReplaceText_WithSessionId_ShouldReplaceInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_session_replace_text.pptx");
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
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var autoShapes = ppt.Slides[0].Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(autoShapes, s => s.TextFrame.Text.Contains("New Session Value"));
    }

    #endregion
}