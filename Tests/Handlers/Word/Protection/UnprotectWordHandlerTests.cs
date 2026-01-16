using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Protection;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Protection;

public class UnprotectWordHandlerTests : WordHandlerTestBase
{
    private readonly UnprotectWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Unprotect()
    {
        Assert.Equal("unprotect", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithWrongPassword_ThrowsInvalidOperationException()
    {
        var doc = CreateProtectedDocument("correctpassword");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "wrongpassword" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateProtectedDocument(string password, ProtectionType type = ProtectionType.ReadOnly)
    {
        var doc = new Document();
        doc.Protect(type, password);
        return doc;
    }

    #endregion

    #region Basic Unprotect Operations

    [Fact]
    public void Execute_UnprotectsDocument()
    {
        var doc = CreateProtectedDocument("test123");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.Equal(ProtectionType.NoProtection, doc.ProtectionType);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsPreviousProtectionType()
    {
        var doc = CreateProtectedDocument("test123", ProtectionType.AllowOnlyComments);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AllowOnlyComments", result);
    }

    [Fact]
    public void Execute_WithUnprotectedDocument_ReturnsNoNeedMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("not protected", result.ToLower());
    }

    #endregion

    #region Various Protection Types

    [Theory]
    [InlineData(ProtectionType.AllowOnlyComments)]
    [InlineData(ProtectionType.AllowOnlyFormFields)]
    [InlineData(ProtectionType.AllowOnlyRevisions)]
    [InlineData(ProtectionType.ReadOnly)]
    public void Execute_WithVariousProtectionTypes_UnprotectsSuccessfully(ProtectionType protectionType)
    {
        var doc = CreateProtectedDocument("test123", protectionType);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.Equal(ProtectionType.NoProtection, doc.ProtectionType);
    }

    [Fact]
    public void Execute_WithFormFieldsProtection_ReturnsPreviousType()
    {
        var doc = CreateProtectedDocument("pass", ProtectionType.AllowOnlyFormFields);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "pass" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AllowOnlyFormFields", result);
    }

    [Fact]
    public void Execute_WithRevisionsProtection_ReturnsPreviousType()
    {
        var doc = CreateProtectedDocument("pass", ProtectionType.AllowOnlyRevisions);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "pass" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AllowOnlyRevisions", result);
    }

    #endregion

    #region Password Edge Cases

    [Fact]
    public void Execute_WithNullPassword_ThrowsOnProtectedDocument()
    {
        var doc = CreateProtectedDocument("test123");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyPassword_ThrowsOnProtectedDocument()
    {
        var doc = CreateProtectedDocument("test123");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
