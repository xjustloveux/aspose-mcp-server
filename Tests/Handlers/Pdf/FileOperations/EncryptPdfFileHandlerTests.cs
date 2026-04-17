using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class EncryptPdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly EncryptPdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Encrypt()
    {
        Assert.Equal("encrypt", _handler.Operation);
    }

    #endregion

    #region Basic Encrypt Operations

    [Fact]
    public void Execute_EncryptsPdf()
    {
        var doc = CreateDocumentWithText("Confidential content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "user123" },
            { "ownerPassword", "owner456" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("encrypted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
        Assert.True(doc.IsEncrypted, "Document should be encrypted after operation");
    }

    #endregion

    #region Default Parameters Back-Compat

    [Fact]
    public void Execute_DefaultParams_BackCompatBehavior()
    {
        var doc = CreateDocumentWithText("Confidential");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "user123" },
            { "ownerPassword", "owner456" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted, "Document should be encrypted");
        AssertModified(context);
    }

    #endregion

    #region Combined Parameters

    [Fact]
    public void Execute_AllNewParams_AESx256_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "AESx256" },
            { "permissions", new[] { "AssembleDocument" } },
            { "usePdf20", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
        AssertModified(context);
    }

    #endregion

    #region Error Message Safety

    [Fact]
    public void Execute_ErrorMessage_NoPathLeak()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "INVALID_ALGO_XYZ" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.DoesNotContain("/", ex.Message);
        Assert.DoesNotContain("\\", ex.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutUserPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "ownerPassword", "owner456" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOwnerPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "user123" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutAnyPassword_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Algorithm Variants

    [Fact]
    public void Execute_WithAlgorithm_AESx128_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "AESx128" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithAlgorithm_RC4x128_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "RC4x128" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithAlgorithm_CaseInsensitive_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "aesx256" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithAlgorithm_EmptyString_UsesDefault()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithInvalidAlgorithm_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "RSA" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("RSA", ex.Message);
        Assert.Contains("AESx256", ex.Message);
        Assert.DoesNotContain("/", ex.Message);
        Assert.DoesNotContain("\\", ex.Message);
    }

    [Fact]
    public void Execute_WithAlgorithm_Custom_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "Custom" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Custom", ex.Message);
        Assert.Contains("AESx256", ex.Message);
    }

    #endregion

    #region Permissions Variants

    [Fact]
    public void Execute_WithPermissions_SingleFlag_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", new[] { "PrintDocument" } }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithPermissions_MultipleFlags_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", new[] { "PrintDocument", "FillForm" } }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithPermissions_EmptyArray_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", Array.Empty<string>() }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithPermissions_CaseInsensitive_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", new[] { "printdocument" } }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithPermissions_None_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", new[] { "None" } }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("None", ex.Message);
        Assert.Contains("PrintDocument", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidPermission_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "permissions", new[] { "FullControl" } }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("FullControl", ex.Message);
        Assert.Contains("PrintDocument", ex.Message);
        Assert.DoesNotContain("/", ex.Message);
        Assert.DoesNotContain("\\", ex.Message);
    }

    #endregion

    #region UsePdf20 Variants

    [Fact]
    public void Execute_WithUsePdf20True_AESx256_Succeeds()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "usePdf20", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(doc.IsEncrypted);
    }

    [Fact]
    public void Execute_WithUsePdf20True_RC4x128_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "RC4x128" },
            { "usePdf20", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("usePdf20=true", ex.Message);
        Assert.Contains("AESx256", ex.Message);
    }

    [Fact]
    public void Execute_WithUsePdf20True_AESx128_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "userPassword", "u" },
            { "ownerPassword", "o" },
            { "algorithm", "AESx128" },
            { "usePdf20", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("usePdf20=true", ex.Message);
        Assert.Contains("AESx256", ex.Message);
    }

    #endregion
}
