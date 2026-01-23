using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class AddPptTableHandlerTests : PptHandlerTestBase
{
    private readonly AddPptTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTableToSlide()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 3 },
            { "columns", 4 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Table added", result.Message);
        Assert.Contains("3 rows", result.Message);
        Assert.Contains("4 columns", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("shapeIndex", result.Message);
    }

    [Fact]
    public void Execute_AddsTableShape()
    {
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 3 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
    }

    [Fact]
    public void Execute_CreatesTableWithCorrectDimensions()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 4 },
            { "columns", 5 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.Equal(4, table.Rows.Count);
        Assert.Equal(5, table.Columns.Count);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_AddsToCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "rows", 2 },
            { "columns", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 1", result.Message);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_AddsToFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 0", result.Message);
    }

    #endregion

    #region Position Parameters

    [Fact]
    public void Execute_WithCustomPosition_SetsPosition()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 },
            { "x", 200.0 },
            { "y", 150.0 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.Equal(200.0f, table.X, 0.1);
        Assert.Equal(150.0f, table.Y, 0.1);
    }

    [Fact]
    public void Execute_WithDefaultPosition_UsesDefault()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.Equal(100.0f, table.X, 0.1);
        Assert.Equal(100.0f, table.Y, 0.1);
    }

    #endregion

    #region Size Parameters

    [Fact]
    public void Execute_WithCustomColumnWidth_SetsWidth()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 3 },
            { "columnWidth", 150.0 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.True(table.Width > 400);
    }

    [Fact]
    public void Execute_WithCustomRowHeight_SetsHeight()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 3 },
            { "columns", 2 },
            { "rowHeight", 50.0 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.True(table.Height > 100);
    }

    #endregion

    #region Data Parameter

    [Fact]
    public void Execute_WithData_PopulatesCells()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 },
            { "data", "[[\"A1\",\"B1\"],[\"A2\",\"B2\"]]" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.Equal("A1", table[0, 0].TextFrame.Text);
        Assert.Equal("B1", table[1, 0].TextFrame.Text);
        Assert.Equal("A2", table[0, 1].TextFrame.Text);
        Assert.Equal("B2", table[1, 1].TextFrame.Text);
    }

    [Fact]
    public void Execute_WithoutData_CreatesEmptyCells()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 2 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[^1] as ITable;
        Assert.NotNull(table);
        Assert.Equal(string.Empty, table[0, 0].TextFrame.Text);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRows_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columns", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rows", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumns_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columns", ex.Message);
    }

    [Fact]
    public void Execute_WithZeroRows_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 0 },
            { "columns", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rows", ex.Message);
    }

    [Fact]
    public void Execute_WithZeroColumns_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", 2 },
            { "columns", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columns", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeRows_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", -1 },
            { "columns", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rows", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "rows", 2 },
            { "columns", 2 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
