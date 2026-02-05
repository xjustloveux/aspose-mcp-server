using Aspose.Email.Mapi;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Contact;

/// <summary>
///     Handler for creating a new email contact (VCF/MSG).
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CreateEmailContactHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new email contact and saves it to the specified output path.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: outputPath.
    ///     Optional: displayName, email, phone, company, jobTitle, format (default: "vcf").
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the contact was created.</returns>
    /// <exception cref="ArgumentException">Thrown when outputPath is missing or invalid.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetRequired<string>("outputPath");
        var displayName = parameters.GetOptional<string?>("displayName");
        var email = parameters.GetOptional<string?>("email");
        var phone = parameters.GetOptional<string?>("phone");
        var company = parameters.GetOptional<string?>("company");
        var jobTitle = parameters.GetOptional<string?>("jobTitle");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var contact = new MapiContact();
        contact.NameInfo = new MapiContactNamePropertySet
        {
            DisplayName = displayName ?? "",
            GivenName = displayName?.Split(' ').FirstOrDefault() ?? ""
        };

        if (!string.IsNullOrEmpty(email))
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress
            {
                EmailAddress = email
            };

        if (!string.IsNullOrEmpty(phone))
            contact.Telephones.PrimaryTelephoneNumber = phone;

        if (!string.IsNullOrEmpty(company))
            contact.ProfessionalInfo.CompanyName = company;

        if (!string.IsNullOrEmpty(jobTitle))
            contact.ProfessionalInfo.Title = jobTitle;

        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        contact.Save(outputPath, ext == ".msg" ? ContactSaveFormat.Msg : ContactSaveFormat.VCard);

        return new SuccessResult
        {
            Message = $"Contact created: {displayName ?? "Unknown"} -> {outputPath}"
        };
    }
}
