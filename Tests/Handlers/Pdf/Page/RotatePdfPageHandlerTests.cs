using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class RotatePdfPageHandlerTests : PdfHandlerTestBase
{
    private static readonly int[] MultiplePageIndices = [1, 3, 5];

    private readonly RotatePdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Rotate()
    {
        Assert.Equal("rotate", _handler.Operation);
    }

    #endregion

    #region Multiple Page Rotation

    [SkippableFact]
    public void Execute_WithPageIndices_RotatesMultiplePages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 180 },
            { "pageIndices", MultiplePageIndices }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
        Assert.Equal(Rotation.on180, doc.Pages[1].Rotate);
        Assert.Equal(Rotation.None, doc.Pages[2].Rotate);
        Assert.Equal(Rotation.on180, doc.Pages[3].Rotate);
        Assert.Equal(Rotation.None, doc.Pages[4].Rotate);
        Assert.Equal(Rotation.on180, doc.Pages[5].Rotate);
        AssertModified(context);
    }

    #endregion

    #region Out of Range Page Indices

    [Fact]
    public void Execute_WithOutOfRangePageIndex_IgnoresInvalidPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 90 },
            { "pageIndex", 100 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Rotated", result.Message);
    }

    #endregion

    #region Basic Rotation

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void Execute_WithValidRotation_RotatesPage(int rotation)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", rotation }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Rotated", result.Message);
        Assert.Contains(rotation.ToString(), result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_RotatesAllPagesDefault()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 90 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
        Assert.Equal(Rotation.on90, doc.Pages[1].Rotate);
        Assert.Equal(Rotation.on90, doc.Pages[2].Rotate);
        Assert.Equal(Rotation.on90, doc.Pages[3].Rotate);
        AssertModified(context);
    }

    #endregion

    #region Single Page Rotation

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithPageIndex_RotatesSinglePage(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 90 },
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("1", result.Message);
        Assert.Equal(Rotation.on90, doc.Pages[pageIndex].Rotate);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithPageIndex_PreservesOtherPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 90 },
            { "pageIndex", 2 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(Rotation.None, doc.Pages[1].Rotate);
        Assert.Equal(Rotation.on90, doc.Pages[2].Rotate);
        Assert.Equal(Rotation.None, doc.Pages[3].Rotate);
    }

    #endregion

    #region Rotation Values

    [Fact]
    public void Execute_With90Degrees_SetsOn90()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 90 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(Rotation.on90, doc.Pages[1].Rotate);
    }

    [Fact]
    public void Execute_With180Degrees_SetsOn180()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 180 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(Rotation.on180, doc.Pages[1].Rotate);
    }

    [Fact]
    public void Execute_With270Degrees_SetsOn270()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 270 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(Rotation.on270, doc.Pages[1].Rotate);
    }

    [Fact]
    public void Execute_With0Degrees_SetsNone()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Pages[1].Rotate = Rotation.on90;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(Rotation.None, doc.Pages[1].Rotate);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRotation_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rotation", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(45)]
    [InlineData(100)]
    [InlineData(-90)]
    [InlineData(360)]
    public void Execute_WithInvalidRotation_ThrowsArgumentException(int invalidRotation)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rotation", invalidRotation }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rotation", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
