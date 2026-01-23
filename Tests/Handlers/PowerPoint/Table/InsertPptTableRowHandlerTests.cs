using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class InsertPptTableRowHandlerTests : PptHandlerTestBase
{
    private readonly InsertPptTableRowHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertRow()
    {
        Assert.Equal("insert_row", _handler.Operation);
    }

    #endregion

    #region CopyFromRow Parameter

    [Fact]
    public void Execute_WithCopyFromRow_CopiesFromSpecifiedRow()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "copyFromRow", 1 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.Equal(4, table!.Rows.Count);
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
        Assert.Equal(3, table!.Rows.Count);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsRow()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Row inserted", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_IncreasesRowCount()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialRowCount = table!.Rows.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialRowCount + 1, table.Rows.Count);
    }

    [Fact]
    public void Execute_ReturnsInsertIndex()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 1", result.Message);
    }

    #endregion

    #region Row Index Parameter

    [Fact]
    public void Execute_DefaultRowIndex_InsertsAtEnd()
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
    public void Execute_WithRowIndex_InsertsAtCorrectPosition(int rowIndex)
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", rowIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"index {rowIndex}", result.Message);
    }

    [Fact]
    public void Execute_WithRowIndexAtEnd_InsertsAtEnd()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 3 }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.Equal(4, table!.Rows.Count);
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
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeRowIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", -1 }
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
