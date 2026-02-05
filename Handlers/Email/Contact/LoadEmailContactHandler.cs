using Aspose.Email.Mapi;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Contact;

namespace AsposeMcpServer.Handlers.Email.Contact;

/// <summary>
///     Handler for loading and retrieving email contact information.
/// </summary>
[ResultType(typeof(ContactEmailInfo))]
public class LoadEmailContactHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_info";

    /// <summary>
    ///     Loads an email contact from a VCF or MSG file and returns its information.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: path (VCF or MSG file path).
    /// </param>
    /// <returns>A <see cref="ContactEmailInfo" /> containing the contact details.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");

        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        var contact = LoadContact(path);

        return new ContactEmailInfo
        {
            DisplayName = contact.NameInfo?.DisplayName,
            Email = contact.ElectronicAddresses?.Email1?.EmailAddress,
            Phone = contact.Telephones?.PrimaryTelephoneNumber,
            Company = contact.ProfessionalInfo?.CompanyName,
            JobTitle = contact.ProfessionalInfo?.Title,
            HasPhoto = contact.Photo?.Data is { Length: > 0 },
            Message = $"Contact loaded: {contact.NameInfo?.DisplayName ?? "Unknown"}"
        };
    }

    /// <summary>
    ///     Loads a MapiContact from a file, detecting format by extension.
    /// </summary>
    /// <param name="path">The file path to load from.</param>
    /// <returns>The loaded <see cref="MapiContact" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the contact cannot be loaded.</exception>
    internal static MapiContact LoadContact(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is ".vcf")
            return MapiContact.FromVCard(path);

        var msg = MapiMessage.Load(path);
        return (MapiContact)msg.ToMapiMessageItem()
               ?? throw new ArgumentException($"Failed to load contact from: {path}");
    }
}
