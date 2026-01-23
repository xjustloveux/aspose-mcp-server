using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class InsertPptTableColumnHandlerTests : PptHandlerTestBase
{
    private readonly InsertPptTableColumnHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertColumn()
    {
        Assert.Equal("insert_column", _handler.Operation);
    }

    #endregion

    #region CopyFromColumn Parameter

    [Fact]
    public void Execute_WithCopyFromColumn_CopiesFromSpecifiedColumn()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "copyFromColumn", 1 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.Equal(4, table!.Columns.Count);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_InsertsOnCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        Assert.Equal(3, table!.Columns.Count);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsColumn()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Column inserted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_IncreasesColumnCount()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount + 1, table.Columns.Count);
    }

    [Fact]
    public void Execute_ReturnsInsertIndex()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 1", result.Message);
    }

    #endregion

    #region Column Index Parameter

    [Fact]
    public void Execute_DefaultColumnIndex_InsertsAtEnd()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 3", result.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithColumnIndex_InsertsAtCorrectPosition(int colIndex)
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", colIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"index {colIndex}", result.Message);
    }

    [Fact]
    public void Execute_WithColumnIndexAtEnd_InsertsAtEnd()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 3 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.Equal(4, table!.Columns.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonTableShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Sample");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a table", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidColumnIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeColumnIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", -1 }
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
