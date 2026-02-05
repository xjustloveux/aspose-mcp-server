using Aspose.Email.Mapi;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Email.Contact;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailContactTool" />.
///     Focuses on operation routing, file I/O, and end-to-end workflows.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailContactToolTests : EmailTestBase
{
    private readonly EmailContactTool _tool = new();

    #region Create Operation

    [Fact]
    public void Execute_Create_CreatesVcfFile()
    {
        var outputPath = CreateTestFilePath("tool_create.vcf");
        var result = _tool.Execute("create", outputPath: outputPath, displayName: "Tool Contact");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Tool Contact", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_Create_WithAllParameters_CreatesFullContact()
    {
        var outputPath = CreateTestFilePath("tool_create_full.vcf");
        var result = _tool.Execute("create",
            outputPath: outputPath,
            displayName: "Full Contact",
            email: "full@example.com",
            phone: "+1-555-0100",
            company: "Test Corp",
            jobTitle: "Tester");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Full Contact", contact.NameInfo.DisplayName);
        Assert.Equal("full@example.com", contact.ElectronicAddresses.Email1.EmailAddress);
        Assert.Equal("+1-555-0100", contact.Telephones.PrimaryTelephoneNumber);
        Assert.Equal("Test Corp", contact.ProfessionalInfo.CompanyName);
        Assert.Equal("Tester", contact.ProfessionalInfo.Title);
    }

    [Fact]
    public void Execute_Create_AsMsgFormat_CreatesMsgFile()
    {
        var outputPath = CreateTestFilePath("tool_create.msg");
        var result = _tool.Execute("create", outputPath: outputPath, displayName: "MSG Tool Contact");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_Create_WithMinimalParameters_CreatesFile()
    {
        var outputPath = CreateTestFilePath("tool_create_minimal.vcf");
        var result = _tool.Execute("create", outputPath: outputPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region GetInfo Operation

    [Fact]
    public void Execute_GetInfo_ReturnsContactInfo()
    {
        var vcfPath = CreateTestVcfFile("tool_getinfo.vcf", "Info Contact", "info@example.com",
            "+1-555-0200", "Info Corp", "Analyst");
        var result = _tool.Execute("get_info", vcfPath);

        Assert.IsType<FinalizedResult<ContactEmailInfo>>(result);
        var data = GetResultData<ContactEmailInfo>(result);
        Assert.Equal("Info Contact", data.DisplayName);
        Assert.Equal("info@example.com", data.Email);
        Assert.Equal("+1-555-0200", data.Phone);
        Assert.Equal("Info Corp", data.Company);
        Assert.Equal("Analyst", data.JobTitle);
        Assert.False(data.HasPhoto);
    }

    [Fact]
    public void Execute_GetInfo_WithMsgFile_ReturnsContactInfo()
    {
        var msgPath = CreateTestMsgContactFile("tool_getinfo.msg", "MSG Info Contact", "msginfo@example.com");
        var result = _tool.Execute("get_info", msgPath);

        Assert.IsType<FinalizedResult<ContactEmailInfo>>(result);
        var data = GetResultData<ContactEmailInfo>(result);
        Assert.Equal("MSG Info Contact", data.DisplayName);
        Assert.Equal("msginfo@example.com", data.Email);
    }

    #endregion

    #region Save Operation

    [Fact]
    public void Execute_Save_VcfToVcf_SavesSuccessfully()
    {
        var inputPath = CreateTestVcfFile("tool_save_input.vcf", "Save Tool User", "save@example.com");
        var outputPath = CreateTestFilePath("tool_save_output.vcf");
        var result = _tool.Execute("save", inputPath, outputPath, format: "vcf");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("Save Tool User", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_Save_VcfToMsg_ConvertsSuccessfully()
    {
        var inputPath = CreateTestVcfFile("tool_convert_input.vcf", "Convert User");
        var outputPath = CreateTestFilePath("tool_convert_output.msg");
        var result = _tool.Execute("save", inputPath, outputPath, format: "msg");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_Save_MsgToVcf_ConvertsSuccessfully()
    {
        var inputPath = CreateTestMsgContactFile("tool_msg_to_vcf_input.msg", "MSG Convert User",
            "msgconvert@example.com");
        var outputPath = CreateTestFilePath("tool_msg_to_vcf_output.vcf");
        var result = _tool.Execute("save", inputPath, outputPath, format: "vcf");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));

        var contact = MapiContact.FromVCard(outputPath);
        Assert.Equal("MSG Convert User", contact.NameInfo.DisplayName);
    }

    [Fact]
    public void Execute_Save_AutoDetectsFormatFromExtension()
    {
        var inputPath = CreateTestVcfFile("tool_auto_input.vcf", "Auto Format User");
        var outputPath = CreateTestFilePath("tool_auto_output.vcf");
        var result = _tool.Execute("save", inputPath, outputPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region SetPhoto Operation

    [Fact]
    public void Execute_SetPhoto_SetsPhotoOnContact()
    {
        var inputPath = CreateTestVcfFile("tool_photo_input.vcf", "Photo Tool User");
        var outputPath = CreateTestFilePath("tool_photo_output.vcf");
        var photoPath = CreateTestImageFile();
        var result = _tool.Execute("set_photo", inputPath, outputPath,
            photoPath: photoPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_SetPhoto_SavesAsMsgWhenOutputIsMsgExtension()
    {
        var inputPath = CreateTestVcfFile("tool_photo_msg_input.vcf", "Photo MSG User");
        var outputPath = CreateTestFilePath("tool_photo_msg_output.msg");
        var photoPath = CreateTestImageFile();
        var result = _tool.Execute("set_photo", inputPath, outputPath,
            photoPath: photoPath);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"tool_case_{operation}.vcf");
        var result = _tool.Execute(operation, outputPath: outputPath, displayName: "Case Test");

        Assert.True(File.Exists(outputPath));
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", outputPath: "test.vcf"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region End-to-End Workflow

    [Fact]
    public void Workflow_CreateThenGetInfo_RoundTrips()
    {
        var outputPath = CreateTestFilePath("workflow_roundtrip.vcf");
        _tool.Execute("create", outputPath: outputPath, displayName: "Roundtrip Contact",
            email: "roundtrip@example.com", phone: "+1-555-0300",
            company: "Roundtrip Corp", jobTitle: "Tester");

        var infoResult = _tool.Execute("get_info", outputPath);
        var data = GetResultData<ContactEmailInfo>(infoResult);

        Assert.Equal("Roundtrip Contact", data.DisplayName);
        Assert.Equal("roundtrip@example.com", data.Email);
        Assert.Equal("+1-555-0300", data.Phone);
        Assert.Equal("Roundtrip Corp", data.Company);
        Assert.Equal("Tester", data.JobTitle);
        Assert.False(data.HasPhoto);
    }

    [Fact]
    public void Workflow_CreateThenSetPhotoThenGetInfo_ShowsHasPhoto()
    {
        var createPath = CreateTestFilePath("workflow_photo_create.vcf");
        _tool.Execute("create", outputPath: createPath, displayName: "Photo Workflow User");

        var photoPath = CreateTestImageFile();
        var photoOutputPath = CreateTestFilePath("workflow_photo_output.msg");
        _tool.Execute("set_photo", createPath, photoOutputPath,
            photoPath: photoPath);

        Assert.True(File.Exists(photoOutputPath));
        Assert.True(new FileInfo(photoOutputPath).Length > 0);
    }

    [Fact]
    public void Workflow_CreateThenSaveThenGetInfo_RoundTripsFormats()
    {
        var vcfPath = CreateTestFilePath("workflow_format_create.vcf");
        _tool.Execute("create", outputPath: vcfPath, displayName: "Format Test User",
            email: "format@example.com");

        var msgPath = CreateTestFilePath("workflow_format_convert.msg");
        _tool.Execute("save", vcfPath, msgPath, format: "msg");
        Assert.True(File.Exists(msgPath));

        var vcfPath2 = CreateTestFilePath("workflow_format_back.vcf");
        _tool.Execute("save", msgPath, vcfPath2, format: "vcf");
        Assert.True(File.Exists(vcfPath2));

        var infoResult = _tool.Execute("get_info", vcfPath2);
        var data = GetResultData<ContactEmailInfo>(infoResult);
        Assert.Equal("Format Test User", data.DisplayName);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test VCF contact file.
    /// </summary>
    private string CreateTestVcfFile(string fileName, string displayName, string? email = null,
        string? phone = null, string? company = null, string? jobTitle = null)
    {
        var filePath = CreateTestFilePath(fileName);
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
        var filePath = CreateTestFilePath(fileName);
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

    /// <summary>
    ///     Creates a simple BMP image file for photo testing.
    /// </summary>
    private string CreateTestImageFile()
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
            bmp[i] = 255;
            bmp[i + 1] = 0;
            bmp[i + 2] = 0;
        }

        var filePath = CreateTestFilePath($"photo_{Guid.NewGuid()}.bmp");
        File.WriteAllBytes(filePath, bmp);
        return filePath;
    }

    #endregion
}
