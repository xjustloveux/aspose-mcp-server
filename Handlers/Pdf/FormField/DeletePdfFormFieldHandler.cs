using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for deleting form fields from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePdfFormFieldHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a form field from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldName
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteParameters(parameters);

        var document = context.Document;

        if (document.Form.Cast<Field>().All(f => f.PartialName != p.FieldName))
            throw new ArgumentException($"Form field '{p.FieldName}' not found");

        document.Form.Delete(p.FieldName);
        MarkModified(context);

        return new SuccessResult { Message = $"Deleted form field '{p.FieldName}'." };
    }

    /// <summary>
    ///     Extracts delete parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<string>("fieldName"));
    }

    /// <summary>
    ///     Parameters for deleting a form field.
    /// </summary>
    /// <param name="FieldName">The name of the form field to delete.</param>
    private sealed record DeleteParameters(string FieldName);
}
