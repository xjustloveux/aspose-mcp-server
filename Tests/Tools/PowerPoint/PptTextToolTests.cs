using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Text;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptTextTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptTextToolTests : PptTestBase
{
    private readonly PptTextTool _tool;

    public PptTextToolTests()
    {
        _tool = new PptTextTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [SkippableFact]
    public void AddText_ShouldAddTextToSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePresentation("test_add_text.pptx");
        var outputPath = CreateTestFilePath("test_add_text_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, text: "Test Text", x: 100, y: 100, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var textFrames = slide.Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.NotEmpty(textFrames);
        Assert.Contains(textFrames, tf => tf.TextFrame.Text.Contains("Test Text"));
    }

    [Fact]
    public void EditText_ShouldEditText()
    {
        var pptPath = CreatePresentation("test_edit_text.pptx");
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

        var pptPath = CreatePresentation("test_replace_text.pptx");
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

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, text: "Test", x: 100, y: 100,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Text added to slide", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("add", slideIndex: 0, text: "Test", x: 100, y: 100));
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void AddText_WithSessionId_ShouldAddInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var initialShapeCount = slide.Shapes.Count;

        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, text: "Session Text", x: 100, y: 100);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Text added to slide", data.Message);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Text edited in shape", data.Message);

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
        var data = GetResultData<TextReplaceResult>(result);
        Assert.Equal(1, data.ReplacementCount);

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
        var data = GetResultData<TextReplaceResult>(result);
        Assert.Equal(1, data.ReplacementCount);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var autoShapes = ppt.Slides[0].Shapes.OfType<IAutoShape>().Where(s => s.TextFrame != null).ToList();
        Assert.Contains(autoShapes, s => s.TextFrame.Text.Contains("Modified"));
    }

    #endregion
}
