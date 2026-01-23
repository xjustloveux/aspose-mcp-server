using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class ReplaceImageWordHandlerTests : WordHandlerTestBase
{
    private readonly ReplaceImageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Replace()
    {
        Assert.Equal("replace", _handler.Operation);
    }

    #endregion

    #region Basic Replace Operations

    [Fact]
    public void Execute_ReplacesImage()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 255, 0);
        var doc = CreateDocumentWithImage(tempFile1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "newImagePath", tempFile2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("replaced", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithImagePath_ReplacesImage()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 0, 255);
        var doc = CreateDocumentWithImage(tempFile1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "imagePath", tempFile2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("replaced", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithPreserveSize_PreservesOriginalSize()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 255, 0);
        var doc = CreateDocumentWithImage(tempFile1);
        var originalShape = GetFirstImage(doc);
        var originalWidth = originalShape.Width;
        var originalHeight = originalShape.Height;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "newImagePath", tempFile2 },
            { "preserveSize", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("preserved size", result.Message, StringComparison.OrdinalIgnoreCase);
        var newShape = GetFirstImage(doc);
        Assert.Equal(originalWidth, newShape.Width, 1);
        Assert.Equal(originalHeight, newShape.Height, 1);
    }

    [Fact]
    public void Execute_WithPreservePosition_PreservesPosition()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 255, 0);
        var doc = CreateDocumentWithImage(tempFile1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "newImagePath", tempFile2 },
            { "preservePosition", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("preserved position", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithSmartFit_CalculatesProportionalHeight()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 255, 0);
        var doc = CreateDocumentWithImage(tempFile1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "newImagePath", tempFile2 },
            { "preserveSize", true },
            { "smartFit", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("smart fit", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImagePath_ThrowsArgumentException()
    {
        var tempFile = CreateTempImageFileWithColor(255, 0, 0);
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentImage_ThrowsFileNotFoundException()
    {
        var tempFile = CreateTempImageFileWithColor(255, 0, 0);
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 },
            { "newImagePath", "/nonexistent/path/image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidImageIndex_ThrowsArgumentException()
    {
        var tempFile1 = CreateTempImageFileWithColor(255, 0, 0);
        var tempFile2 = CreateTempImageFileWithColor(0, 255, 0);
        var doc = CreateDocumentWithImage(tempFile1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 99 },
            { "newImagePath", tempFile2 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private string CreateTempImageFileWithColor(byte red, byte green, byte blue)
    {
        var width = 10;
        var height = 10;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = blue;
            bmp[i + 1] = green;
            bmp[i + 2] = red;
        }

        return CreateTempFile(".bmp", bmp);
    }

    private static Document CreateDocumentWithImage(string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        return doc;
    }

    private static WordShape GetFirstImage(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .First(s => s.HasImage);
    }

    #endregion
}
