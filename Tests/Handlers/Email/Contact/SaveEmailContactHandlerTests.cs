using Aspose.Email.Mapi;
using AsposeMcpServer.Handlers.Email.Contact;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Contact;

/// <summary>
///     Tests for <see cref="SaveEmailContactHandler" />.
/// </summary>
public class SaveEmailContactHandlerTests : HandlerTestBase<object>
{
    private readonly SaveEmailContactHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Save()
    {
        Assert.Equal("save", _handler.Operation);
    }

    #endregion

    #region Basic Save Operations

    [Fact]
    public void Execute_SaveAsVcf_CreatesVcfFile()
    {
        var inputPath = CreateTestVcfFile("save_input.vcf", "Save Test User", "save@example.com");
        var outputPath = Path.Combine(TestDir, "save_output.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "vcf" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Save Test User", result.Message);
        Assert.Contains(outputPath, result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Save Test User", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_SaveAsMsg_CreatesMsgFile()
    {
        var inputPath = CreateTestVcfFile("save_msg_input.vcf", "MSG Save User", "msg@example.com");
        var outputPath = Path.Combine(TestDir, "save_output.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "msg" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(outputPath, result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithoutFormat_AutoDetectsFromVcfExtension()
    {
        var inputPath = CreateTestVcfFile("auto_vcf_input.vcf", "Auto VCF User");
        var outputPath = Path.Combine(TestDir, "auto_output.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Auto VCF User", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_WithoutFormat_AutoDetectsFromMsgExtension()
    {
        var inputPath = CreateTestVcfFile("auto_msg_input.vcf", "Auto MSG User");
        var outputPath = Path.Combine(TestDir, "auto_output.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_FromMsgToVcf_ConvertsCorrectly()
    {
        var inputPath = CreateTestMsgContactFile("msg_to_vcf_input.msg", "Convert User", "convert@example.com");
        var outputPath = Path.Combine(TestDir, "msg_to_vcf_output.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "vcf" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Convert User", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_PreservesContactData()
    {
        var inputPath = CreateTestVcfFile("preserve_save_input.vcf", "Preserve User",
            "preserve@example.com", "+1-555-0100", "Test Corp", "Manager");
        var outputPath = Path.Combine(TestDir, "preserve_save_output.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath }
        });

        _handler.Execute(context, parameters);

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Preserve User", contact.NameInfo.DisplayName);
        Assert.Equal("preserve@example.com", contact.ElectronicAddresses.Email1.EmailAddress);
        Assert.Equal("+1-555-0100", contact.Telephones.PrimaryTelephoneNumber);
        Assert.Equal("Test Corp", contact.ProfessionalInfo.CompanyName);
        Assert.Equal("Manager", contact.ProfessionalInfo.Title);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.vcf") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestVcfFile("no_output_save.vcf", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.vcf") },
            { "outputPath", Path.Combine(TestDir, "output.vcf") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test VCF contact file.
    /// </summary>
    private string CreateTestVcfFile(string fileName, string displayName, string? email = null,
        string? phone = null, string? company = null, string? jobTitle = null)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var contact = new MapiContact();
        contact.NameInfo = new MapiContactNamePropertySet
        {
            DisplayName = displayName,
            GivenName = displayName.Split(' ').FirstOrDefault() ?? ""
        };

        if (email != null)
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress
            {
                EmailAddress = email
            };

        if (phone != null)
            contact.Telephones.PrimaryTelephoneNumber = phone;

        if (company != null)
            contact.ProfessionalInfo.CompanyName = company;

        if (jobTitle != null)
            contact.ProfessionalInfo.Title = jobTitle;

        contact.Save(filePath, ContactSaveFormat.VCard);
        return filePath;
    }

    /// <summary>
    ///     Creates a test MSG contact file.
    /// </summary>
    private string CreateTestMsgContactFile(string fileName, string displayName, string? email = null)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var contact = new MapiContact();
        contact.NameInfo = new MapiContactNamePropertySet { DisplayName = displayName };

        if (email != null)
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress
            {
                EmailAddress = email
            };

        contact.Save(filePath, ContactSaveFormat.Msg);
        return filePath;
    }

    #endregion
}
