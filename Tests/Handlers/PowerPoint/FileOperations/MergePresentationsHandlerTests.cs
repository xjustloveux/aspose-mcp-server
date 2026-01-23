using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

public class MergePresentationsHandlerTests : PptHandlerTestBase
{
    private readonly MergePresentationsHandler _handler = new();
    private readonly string _input1Path;
    private readonly string _input2Path;

    public MergePresentationsHandlerTests()
    {
        _input1Path = Path.Combine(TestDir, "input1.pptx");
        using (var pres1 = new Presentation())
        {
            pres1.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
            pres1.Save(_input1Path, SaveFormat.Pptx);
        }

        _input2Path = Path.Combine(TestDir, "input2.pptx");
        using (var pres2 = new Presentation())
        {
            pres2.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 50, 50, 100, 100);
            pres2.Save(_input2Path, SaveFormat.Pptx);
        }
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Merge()
    {
        Assert.Equal("merge", _handler.Operation);
    }

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesPresentations()
    {
        var outputPath = Path.Combine(TestDir, "merged.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath },
            { "inputPaths", new[] { _input1Path, _input2Path } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        using var mergedPres = new Presentation(outputPath);
        Assert.Equal(2, mergedPres.Slides.Count);
    }

    [Fact]
    public void Execute_WithOutputPath_MergesPresentations()
    {
        var outputPath = Path.Combine(TestDir, "merged_output.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "inputPaths", new[] { _input1Path, _input2Path } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");
    }

    [Fact]
    public void Execute_WithoutKeepSourceFormatting_MergesPresentations()
    {
        var outputPath = Path.Combine(TestDir, "merged_no_keep.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath },
            { "inputPaths", new[] { _input1Path, _input2Path } },
            { "keepSourceFormatting", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrOutputPath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { _input1Path, _input2Path } }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutInputPaths_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "merged.pptx") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyInputPaths_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "merged.pptx") },
            { "inputPaths", Array.Empty<string>() }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
