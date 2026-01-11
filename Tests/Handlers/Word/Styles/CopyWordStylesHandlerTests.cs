using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Styles;

public class CopyWordStylesHandlerTests : WordHandlerTestBase
{
    private readonly CopyWordStylesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopyStyles()
    {
        Assert.Equal("copy_styles", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSourceDocument_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentSourceDocument_ThrowsFileNotFoundException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceDocument", @"C:\nonexistent\document.docx" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
