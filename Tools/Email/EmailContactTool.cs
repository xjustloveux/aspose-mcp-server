using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for creating and managing email contacts (VCF/MSG).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.Contact")]
[McpServerToolType]
public class EmailContactTool
{
    /// <summary>
    ///     Handler registry for email contact operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailContactTool" /> class.
    /// </summary>
    public EmailContactTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.Contact");
    }

    /// <summary>
    ///     Executes an email contact operation (create, get_info, save, set_photo).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: create, get_info, save, set_photo.
    /// </param>
    /// <param name="path">Input contact file path (VCF or MSG, required for get_info, save, set_photo).</param>
    /// <param name="outputPath">Output file path (required for create, save, set_photo).</param>
    /// <param name="displayName">Display name of the contact (for create).</param>
    /// <param name="email">Email address of the contact (for create).</param>
    /// <param name="phone">Phone number of the contact (for create).</param>
    /// <param name="company">Company name of the contact (for create).</param>
    /// <param name="jobTitle">Job title of the contact (for create).</param>
    /// <param name="photoPath">Photo image file path (required for set_photo).</param>
    /// <param name="format">Output format: "vcf" or "msg" (optional, auto-detected from extension).</param>
    /// <returns>Contact operation result depending on the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_contact",
        Title = "Email Contact Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Create and manage email contacts (VCF/MSG). Supports 4 operations: create, get_info, save, set_photo.

Usage examples:
- Create contact: email_contact(operation='create', outputPath='contact.vcf', displayName='John Doe', email='john@example.com')
- Get info: email_contact(operation='get_info', path='contact.vcf')
- Convert format: email_contact(operation='save', path='contact.vcf', outputPath='contact.msg')
- Set photo: email_contact(operation='set_photo', path='contact.vcf', outputPath='contact_photo.vcf', photoPath='photo.jpg')

Supported formats: VCF (vCard), MSG (Outlook)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a new contact (required params: outputPath; optional: displayName, email, phone, company, jobTitle, format)
- 'get_info': Get contact information (required params: path)
- 'save': Save/convert contact to a different format (required params: path, outputPath; optional: format)
- 'set_photo': Set a photo on a contact (required params: path, outputPath, photoPath)")]
        string operation,
        [Description("Input contact file path (VCF or MSG)")]
        string? path = null,
        [Description("Output file path for the contact")]
        string? outputPath = null,
        [Description("Display name of the contact")]
        string? displayName = null,
        [Description("Email address of the contact")]
        string? email = null,
        [Description("Phone number of the contact")]
        string? phone = null,
        [Description("Company name of the contact")]
        string? company = null,
        [Description("Job title of the contact")]
        string? jobTitle = null,
        [Description("Photo image file path (for set_photo)")]
        string? photoPath = null,
        [Description("Output format: 'vcf' or 'msg' (auto-detected from extension if not specified)")]
        string? format = null)
    {
        var parameters = BuildParameters(path, outputPath, displayName, email, phone, company, jobTitle, photoPath,
            format);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        return ResultHelper.FinalizeResult((dynamic)result, outputPath ?? path, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="displayName">The contact display name.</param>
    /// <param name="email">The contact email address.</param>
    /// <param name="phone">The contact phone number.</param>
    /// <param name="company">The contact company name.</param>
    /// <param name="jobTitle">The contact job title.</param>
    /// <param name="photoPath">The photo image file path.</param>
    /// <param name="format">The output format.</param>
    /// <returns>OperationParameters configured for the contact operation.</returns>
    private static OperationParameters BuildParameters(
        string? path,
        string? outputPath,
        string? displayName,
        string? email,
        string? phone,
        string? company,
        string? jobTitle,
        string? photoPath,
        string? format)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("displayName", displayName);
        parameters.SetIfNotNull("email", email);
        parameters.SetIfNotNull("phone", phone);
        parameters.SetIfNotNull("company", company);
        parameters.SetIfNotNull("jobTitle", jobTitle);
        parameters.SetIfNotNull("photoPath", photoPath);
        parameters.SetIfNotNull("format", format);
        return parameters;
    }
}
