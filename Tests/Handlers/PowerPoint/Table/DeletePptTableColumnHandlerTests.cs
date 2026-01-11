using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class DeletePptTableColumnHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptTableColumnHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteColumn()
    {
        Assert.Equal("delete_column", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_DeletesFromCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 3);
        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex },
            { "columnIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesColumn()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DecreasesColumnCount()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    [Fact]
    public void Execute_ReturnsDeletedColumnIndex()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Column 1", result);
    }

    #endregion

    #region Various Column Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesColumnAtVariousIndices(int colIndex)
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", colIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    [Fact]
    public void Execute_DeletesFirstColumn()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Column 0", result);
    }

    [Fact]
    public void Execute_DeletesLastColumn()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Column 2", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 },
            { "columnIndex", 0 }
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
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a table", ex.Message);
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
