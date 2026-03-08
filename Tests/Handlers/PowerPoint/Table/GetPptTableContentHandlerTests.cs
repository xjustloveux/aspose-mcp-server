using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.PowerPoint.Table;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

[SupportedOSPlatform("windows")]
public class GetPptTableContentHandlerTests : PptHandlerTestBase
{
    private readonly GetPptTableContentHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetContent()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [SkippableFact]
    public void Execute_WithSlideIndex_GetsFromCorrectSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(1, result.SlideIndex);
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_GetsTableContent()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.True(result.RowCount > 0);
        Assert.True(result.ColumnCount > 0);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectRowCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(4, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(4, result.RowCount);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectColumnCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(5, result.ColumnCount);
    }

    [SkippableFact]
    public void Execute_ReturnsDataArray()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.NotNull(result.Data);
        Assert.Equal(2, result.Data.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(0, result.SlideIndex);
    }

    [SkippableFact]
    public void Execute_ReturnsShapeIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(0, result.ShapeIndex);
    }

    #endregion

    #region Cell Data

    [SkippableFact]
    public void Execute_ReturnsCorrectCellData()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTableData([["A1", "B1"], ["A2", "B2"]]);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal("A1", result.Data[0][0]);
        Assert.Equal("B1", result.Data[0][1]);
        Assert.Equal("A2", result.Data[1][0]);
        Assert.Equal("B2", result.Data[1][1]);
    }

    [SkippableFact]
    public void Execute_WithEmptyCells_ReturnsEmptyStrings()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTableContentResult>(res);

        Assert.Equal(string.Empty, result.Data[0][0]);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

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
            { "shapeIndex", 99 }
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
            { "shapeIndex", 0 }
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
            { "shapeIndex", 0 }
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

    private static Presentation CreatePresentationWithTableData(string[][] data)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var rows = data.Length;
        var cols = data[0].Length;
        var colWidths = Enumerable.Repeat(100.0, cols).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);

        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
            table[col, row].TextFrame.Text = data[row][col];

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
