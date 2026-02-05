using Aspose.Email.Mapi;
using AsposeMcpServer.Handlers.Email.Contact;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Contact;

/// <summary>
///     Tests for <see cref="CreateEmailContactHandler" />.
/// </summary>
public class CreateEmailContactHandlerTests : HandlerTestBase<object>
{
    private readonly CreateEmailContactHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_WithOutputPathOnly_CreatesDefaultContact()
    {
        var outputPath = Path.Combine(TestDir, "default.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(outputPath, result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithAllParameters_CreatesFullContact()
    {
        var outputPath = Path.Combine(TestDir, "full_contact.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "displayName", "John Doe" },
            { "email", "john@example.com" },
            { "phone", "+1-555-0100" },
            { "company", "Acme Corp" },
            { "jobTitle", "Senior Engineer" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("John Doe", result.Message);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("John Doe", contact.NameInfo.DisplayName);
        Assert.Equal("john@example.com", contact.ElectronicAddresses.Email1.EmailAddress);
        Assert.Equal("+1-555-0100", contact.Telephones.PrimaryTelephoneNumber);
        Assert.Equal("Acme Corp", contact.ProfessionalInfo.CompanyName);
        Assert.Equal("Senior Engineer", contact.ProfessionalInfo.Title);
    }

    [Fact]
    public void Execute_WithDisplayNameOnly_SetsDisplayNameAndGivenName()
    {
        var outputPath = Path.Combine(TestDir, "name_only.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "displayName", "Jane Smith" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Jane Smith", result.Message);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Jane Smith", contact.NameInfo.DisplayName);
        Assert.Equal("Jane", contact.NameInfo.GivenName);
    }

    [Fact]
    public void Execute_WithEmailOnly_SetsEmail()
    {
        var outputPath = Path.Combine(TestDir, "email_only.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "email", "test@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("test@example.com", contact.ElectronicAddresses.Email1.EmailAddress);
    }

    [Fact]
    public void Execute_WithPhoneOnly_SetsPhone()
    {
        var outputPath = Path.Combine(TestDir, "phone_only.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "phone", "+1-800-555-0199" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("+1-800-555-0199", contact.Telephones.PrimaryTelephoneNumber);
    }

    [Fact]
    public void Execute_WithCompanyAndJobTitle_SetsProfessionalInfo()
    {
        var outputPath = Path.Combine(TestDir, "professional.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "company", "Tech Corp" },
            { "jobTitle", "CTO" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Tech Corp", contact.ProfessionalInfo.CompanyName);
        Assert.Equal("CTO", contact.ProfessionalInfo.Title);
    }

    [Fact]
    public void Execute_WithMsgExtension_SavesAsMsgFormat()
    {
        var outputPath = Path.Combine(TestDir, "contact.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "displayName", "MSG Contact" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("MSG Contact", result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithVcfExtension_SavesAsVCardFormat()
    {
        var outputPath = Path.Combine(TestDir, "contact.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "displayName", "VCF Contact" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("VCF Contact", result.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithNoDisplayName_UsesUnknownInMessage()
    {
        var outputPath = Path.Combine(TestDir, "no_name.vcf");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Unknown", result.Message);
    }

    #endregion
}
