using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Protection;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Protection;

public class ProtectWordHandlerTests : WordHandlerTestBase
{
    private readonly ProtectWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Protect()
    {
        Assert.Equal("protect", _handler.Operation);
    }

    #endregion

    #region Basic Protect Operations

    [Fact]
    public void Execute_ProtectsDocument()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("protected", result.ToLower());
        Assert.NotEqual(ProtectionType.NoProtection, doc.ProtectionType);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultsToReadOnly()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(ProtectionType.ReadOnly, doc.ProtectionType);
    }

    [Theory]
    [InlineData("ReadOnly", ProtectionType.ReadOnly)]
    [InlineData("AllowOnlyComments", ProtectionType.AllowOnlyComments)]
    [InlineData("AllowOnlyFormFields", ProtectionType.AllowOnlyFormFields)]
    [InlineData("AllowOnlyRevisions", ProtectionType.AllowOnlyRevisions)]
    public void Execute_WithProtectionType_SetsCorrectType(string typeStr, ProtectionType expectedType)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "protectionType", typeStr }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(expectedType, doc.ProtectionType);
    }

    [Fact]
    public void Execute_ReturnsProtectionTypeInMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "protectionType", "AllowOnlyComments" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AllowOnlyComments", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPassword_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("password", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithEmptyPassword_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("password", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithWhitespacePassword_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "   " }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("password", ex.Message.ToLower());
    }

    #endregion
}
