using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

[SupportedOSPlatform("windows")]
public class EditPptTableHandlerTests : PptHandlerTestBase
{
    private readonly EditPptTableHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Edit()
    {
        SkipIfNotWindows();
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [SkippableFact]
    public void Execute_WithSlideIndex_UpdatesOnCorrectSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex },
            { "x", 200.0f }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        Assert.NotNull(table);
        Assert.Equal(200.0f, table.X, 0.1);
    }

    #endregion

    #region Basic Edit Operations

    [SkippableFact]
    public void Execute_EditsTableProperties()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 200.0f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal(200.0f, table.X, 0.1);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 200.0f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal(200.0f, table.X, 0.1);
    }

    #endregion

    #region Position Updates

    [SkippableFact]
    public void Execute_WithX_UpdatesXPosition()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 300.0f }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal(300.0f, table.X, 0.1);
    }

    [SkippableFact]
    public void Execute_WithY_UpdatesYPosition()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "y", 250.0f }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal(250.0f, table.Y, 0.1);
    }

    [SkippableFact]
    public void Execute_WithXAndY_UpdatesBothPositions()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 150.0f },
            { "y", 180.0f }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal(150.0f, table.X, 0.1);
        Assert.Equal(180.0f, table.Y, 0.1);
    }

    #endregion

    #region Size Updates (Rejected)

    [SkippableFact]
    public void Execute_WithWidth_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "width", 400.0f }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("column widths", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithHeight_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "height", 200.0f }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("row heights", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithWidthAndHeight_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "width", 350.0f },
            { "height", 180.0f }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "x", 200.0f }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 },
            { "x", 200.0f }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNonTableShape_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Sample");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 200.0f }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a table", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 },
            { "x", 200.0f }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithTable(int rows, int cols)
    {
        var pres = new Presentation();
        AddTableToSlide(pres, 0, rows, cols);
        return pres;
    }

    private static int AddTableToSlide(Presentation pres, int slideIndex, int rows, int cols)
    {
        var slide = pres.Slides[slideIndex];
        var colWidths = Enumerable.Repeat(100.0, cols).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        return slide.Shapes.Count - 1;
    }

    #endregion
}
