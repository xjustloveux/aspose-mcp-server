using Microsoft.IdentityModel.Tokens;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Represents the result of building JWT validation parameters.
/// </summary>
/// <param name="Parameters">The token validation parameters, or null if configuration is invalid.</param>
/// <param name="CryptoKey">The disposable cryptographic key (RSA or ECDsa) that needs cleanup.</param>
public record ValidationParametersResult(TokenValidationParameters? Parameters, IDisposable? CryptoKey);
