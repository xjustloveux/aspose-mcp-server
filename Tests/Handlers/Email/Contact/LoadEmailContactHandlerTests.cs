using Aspose.Email.Mapi;
using AsposeMcpServer.Handlers.Email.Contact;
using AsposeMcpServer.Results.Email.Contact;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Contact;

/// <summary>
///     Tests for <see cref="LoadEmailContactHandler" />.
/// </summary>
public class LoadEmailContactHandlerTests : HandlerTestBase<object>
{
    private readonly LoadEmailContactHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
    }

    #endregion

    #region Basic Load Operations

    [Fact]
    public void Execute_WithValidVcfFile_ReturnsContactInfo()
    {
        var vcfPath = CreateTestVcfFile("load_test.vcf", "John Doe", "john@example.com",
            "+1-555-0100", "Acme Corp", "Engineer");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", vcfPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.Equal("John Doe", result.DisplayName);
        Assert.Equal("john@example.com", result.Email);
        Assert.Equal("+1-555-0100", result.Phone);
        Assert.Equal("Acme Corp", result.Company);
        Assert.Equal("Engineer", result.JobTitle);
        Assert.False(result.HasPhoto);
        Assert.Contains("John Doe", result.Message);
    }

    [Fact]
    public void Execute_WithDisplayNameOnly_ReturnsNameOnly()
    {
        var vcfPath = CreateTestVcfFile("name_only.vcf", "Jane Smith");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", vcfPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.Equal("Jane Smith", result.DisplayName);
        Assert.Null(result.Email);
        Assert.Null(result.Phone);
        Assert.False(result.HasPhoto);
    }

    [Fact]
    public void Execute_WithEmailOnly_ReturnsEmailOnly()
    {
        var vcfPath = CreateTestVcfFile("email_only.vcf", email: "test@example.com");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", vcfPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.Equal("test@example.com", result.Email);
    }

    [Fact]
    public void Execute_WithMsgFile_ReturnsContactInfo()
    {
        var msgPath = CreateTestMsgContactFile("load_msg.msg", "MSG Contact", "msg@example.com");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", msgPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.Equal("MSG Contact", result.DisplayName);
        Assert.Equal("msg@example.com", result.Email);
    }

    [Fact]
    public void Execute_WithNoPhoto_ReturnsFalseForHasPhoto()
    {
        var vcfPath = CreateTestVcfFile("no_photo.vcf", "No Photo User");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", vcfPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.False(result.HasPhoto);
    }

    [Fact]
    public void Execute_WithMinimalContact_ReturnsUnknownInMessage()
    {
        var vcfPath = Path.Combine(TestDir, "minimal.vcf");
        var contact = new MapiContact();
        contact.Save(vcfPath, ContactSaveFormat.VCard);

        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", vcfPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ContactEmailInfo>(res);
        Assert.Contains("loaded", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.vcf") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test VCF contact file.
    /// </summary>
    private string CreateTestVcfFile(string fileName, string? displayName = null, string? email = null,
        string? phone = null, string? company = null, string? jobTitle = null)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var contact = new MapiContact();

        if (displayName != null)
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
