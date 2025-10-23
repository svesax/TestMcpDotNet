using Azure.Identity;
using Microsoft.Graph;

/// <summary>
/// Helper class for configuring Microsoft Graph authentication.
/// This demonstrates how to properly initialize the GraphServiceClient with authentication.
/// </summary>
public static class GraphAuthenticationHelper
{
    /// <summary>
    /// Creates a GraphServiceClient using Client Credentials flow (app-only authentication).
    /// This is suitable for scenarios where the application needs to access data without a signed-in user.
    /// </summary>
    /// <param name="tenantId">The Azure AD tenant ID</param>
    /// <param name="clientId">The Azure AD application client ID</param>
    /// <param name="clientSecret">The Azure AD application client secret</param>
    /// <returns>Configured GraphServiceClient</returns>
    public static GraphServiceClient CreateWithClientCredentials(string tenantId, string clientId, string clientSecret)
    {
        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        return new GraphServiceClient(credential);
    }

    /// <summary>
    /// Creates a GraphServiceClient using Client Certificate authentication (app-only authentication).
    /// This is more secure than client secret and recommended for production scenarios.
    /// </summary>
    /// <param name="tenantId">The Azure AD tenant ID</param>
    /// <param name="clientId">The Azure AD application client ID</param>
    /// <param name="certificatePath">Path to the certificate file (.pfx)</param>
    /// <param name="certificatePassword">Password for the certificate (if any)</param>
    /// <returns>Configured GraphServiceClient</returns>
    public static GraphServiceClient CreateWithClientCertificate(string tenantId, string clientId, string certificatePath, string? certificatePassword = null)
    {
        var credential = new ClientCertificateCredential(tenantId, clientId, certificatePath, 
            new ClientCertificateCredentialOptions { SendCertificateChain = true });
        return new GraphServiceClient(credential);
    }

    /// <summary>
    /// Creates a GraphServiceClient using Device Code flow (delegated authentication).
    /// This is suitable for scenarios where user interaction is possible.
    /// </summary>
    /// <param name="tenantId">The Azure AD tenant ID</param>
    /// <param name="clientId">The Azure AD application client ID</param>
    /// <returns>Configured GraphServiceClient</returns>
    public static GraphServiceClient CreateWithDeviceCode(string tenantId, string clientId)
    {
        var credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            TenantId = tenantId,
            ClientId = clientId,
            DeviceCodeCallback = (code, cancellation) =>
            {
                Console.WriteLine($"Go to {code.VerificationUri} and enter code: {code.UserCode}");
                return Task.CompletedTask;
            }
        });
        return new GraphServiceClient(credential);
    }

    /// <summary>
    /// Creates a GraphServiceClient using environment variables for configuration.
    /// Expected environment variables:
    /// - AZURE_TENANT_ID
    /// - AZURE_CLIENT_ID  
    /// - AZURE_CLIENT_SECRET
    /// </summary>
    /// <returns>Configured GraphServiceClient or null if environment variables are not set</returns>
    public static GraphServiceClient? CreateFromEnvironment()
    {
 
var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
var clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET");

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
        {
            return null;
        }

        return CreateWithClientCredentials(tenantId, clientId, clientSecret);
    }
}