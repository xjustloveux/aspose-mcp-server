using Aspose.Email.Mapi;
using AsposeMcpServer.Handlers.Email.Contact;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Contact;

/// <summary>
///     Tests for <see cref="SetPhotoEmailContactHandler" />.
/// </summary>
public class SetPhotoEmailContactHandlerTests : HandlerTestBase<object>
{
    private readonly SetPhotoEmailContactHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPhoto()
    {
        Assert.Equal("set_photo", _handler.Operation);
    }

    #endregion

    #region Basic Set Photo Operations

    [Fact]
    public void Execute_WithValidPhoto_SetsPhotoOnVcfContact()
    {
        var inputPath = CreateTestVcfFile("photo_input.vcf", "Photo User");
        var outputPath = Path.Combine(TestDir, "photo_output.vcf");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "photoPath", photoPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Photo User", result.Message);
        Assert.Contains(outputPath, result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithValidPhoto_SetsPhotoOnMsgContact()
    {
        var inputPath = CreateTestMsgContactFile("photo_msg_input.msg", "MSG Photo User");
        var outputPath = Path.Combine(TestDir, "photo_msg_output.msg");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "photoPath", photoPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("MSG Photo User", result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_PreservesContactData()
    {
        var inputPath = CreateTestVcfFile("preserve_photo_input.vcf", "Preserve Photo User",
            "preserve@example.com", "+1-555-0100", "Corp", "Dev");
        var outputPath = Path.Combine(TestDir, "preserve_photo_output.vcf");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "photoPath", photoPath }
        });

        _handler.Execute(context, parameters);

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Preserve Photo User", contact.NameInfo.DisplayName);
        Assert.Equal("preserve@example.com", contact.ElectronicAddresses.Email1.EmailAddress);
    }

    [Fact]
    public void Execute_WithVcfOutput_SavesAsVCard()
    {
        var inputPath = CreateTestVcfFile("vcf_photo_input.vcf", "VCF Photo Test");
        var outputPath = Path.Combine(TestDir, "vcf_photo_output.vcf");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "photoPath", photoPath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithMsgOutput_SavesAsMsg()
    {
        var inputPath = CreateTestVcfFile("msg_photo_input.vcf", "MSG Output Photo");
        var outputPath = Path.Combine(TestDir, "msg_photo_output.msg");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "photoPath", photoPath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.vcf") },
            { "photoPath", photoPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestVcfFile("no_output_photo.vcf", "Test");
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "photoPath", photoPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutPhotoPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestVcfFile("no_photo_path.vcf", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.vcf") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentInputFile_ThrowsFileNotFoundException()
    {
        var photoPath = CreateTempImageFile();
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.vcf") },
            { "outputPath", Path.Combine(TestDir, "output.vcf") },
            { "photoPath", photoPath }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentPhotoFile_ThrowsFileNotFoundException()
    {
        var inputPath = CreateTestVcfFile("no_photo_file.vcf", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.vcf") },
            { "photoPath", Path.Combine(TestDir, "nonexistent.jpg") }
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
    private string CreateTestMsgContactFile(string fileName, string displayName)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var contact = new MapiContact();
        contact.NameInfo = new MapiContactNamePropertySet { DisplayName = displayName };
        contact.Save(filePath, ContactSaveFormat.Msg);
        return filePath;
    }

    #endregion
}
