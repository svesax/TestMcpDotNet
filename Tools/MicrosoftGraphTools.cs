using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using System.Text.Json;

/// <summary>
/// Microsoft Graph tools for accessing Microsoft 365 data.
/// These tools can be invoked by MCP clients to interact with Microsoft Graph API.
/// </summary>
internal class MicrosoftGraphTools
{
    private readonly GraphServiceClient? _graphServiceClient;
    public MicrosoftGraphTools(GraphServiceClient graphClient)
    {
        // Try to initialize Graph client from environment variables
        // In a production scenario, you would configure this properly with your authentication method

        //_graphServiceClient = GraphAuthenticationHelper.CreateFromEnvironment();
        _graphServiceClient = graphClient;

        // Alternative initialization examples (uncomment and configure as needed):
        // _graphServiceClient = GraphAuthenticationHelper.CreateWithClientCredentials("tenant-id", "client-id", "client-secret");
        // _graphServiceClient = GraphAuthenticationHelper.CreateWithClientCertificate("tenant-id", "client-id", "cert-path");
    }

    [McpServerTool]
    [Description("Gets the current user's profile information")]
    public async Task<string> GetUserProfileAsync()
    {
        try
        {
            var user = await _graphServiceClient.Me.GetAsync();
            return $"User: {user?.DisplayName} ({user?.UserPrincipalName})";
        }
        catch (Exception ex)
        {           
            return $"Error: {ex.Message}";
        }
    }

    [McpServerTool]
    [Description("Gets a list of users from Microsoft Graph with optional filtering. Requires proper authentication to be configured.")]
    public async Task<string> GetUsers(
        [Description("Optional filter to apply to the user query (e.g., \"startswith(displayName,'John')\" or \"department eq 'Sales'\")")] string? filter = null,
        [Description("Optional search query to find users (e.g., \"John Smith\" or \"john@contoso.com\")")] string? search = null,
        [Description("Maximum number of users to return (default: 10, max: 100)")] int top = 10,
        [Description("Comma-separated list of properties to select (e.g., \"displayName,mail,department\")")] string? select = null)
    {
        // Validate top parameter
        if (top <= 0 || top > 100)
        {
            top = 10;
        }

        try
        {
            // Check if Graph client is properly initialized
            if (_graphServiceClient == null)
            {
                return JsonSerializer.Serialize(new
                {
                    error = "Microsoft Graph client not configured",
                    message = "Authentication needs to be properly configured to access Microsoft Graph API",
                    configurationRequired = new
                    {
                        clientId = "Your Azure AD application client ID",
                        tenantId = "Your Azure AD tenant ID",
                        clientSecret = "Your Azure AD application client secret (for app-only access)",
                        scopes = new[] { "https://graph.microsoft.com/.default" }
                    }
                });
            }

            // Execute the Graph API request using the modern syntax
            var users = await _graphServiceClient.Users.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = top;
                
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                }
                
                if (!string.IsNullOrWhiteSpace(search))
                {
                    requestConfiguration.QueryParameters.Search = $"\"{search}\"";
                    // When using search, we need to add the ConsistencyLevel header
                    requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                }
                
                if (!string.IsNullOrWhiteSpace(select))
                {
                    requestConfiguration.QueryParameters.Select = select.Split(',', StringSplitOptions.RemoveEmptyEntries);
                }
            });

            // Format the response
            var userList = users?.Value?.Select(user => new
            {
                id = user.Id,
                displayName = user.DisplayName,
                mail = user.Mail,
                userPrincipalName = user.UserPrincipalName,
                department = user.Department,
                jobTitle = user.JobTitle,
                officeLocation = user.OfficeLocation,
                mobilePhone = user.MobilePhone,
                businessPhones = user.BusinessPhones,
                accountEnabled = user.AccountEnabled
            }).ToList() ?? [];

            var result = new
            {
                totalCount = users?.Value?.Count ?? 0,
                users = userList,
                appliedFilter = filter,
                appliedSearch = search,
                selectedProperties = select
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (ServiceException ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Microsoft Graph API error",
                code = ex.ResponseStatusCode,
                message = ex.Message,
                details = ex.ToString()
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Unexpected error",
                message = ex.Message,
                type = ex.GetType().Name
            });
        }
    }

    [McpServerTool]
    [Description("Gets detailed information about a specific user from Microsoft Graph by their ID or User Principal Name.")]
    public async Task<string> GetUserById(
        [Description("User ID (GUID) or User Principal Name (email) of the user to retrieve")] string userId,
        [Description("Comma-separated list of properties to select (e.g., \"displayName,mail,department,manager\")")] string? select = null)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            return JsonSerializer.Serialize(new
            {
                error = "Invalid input",
                message = "User ID or User Principal Name is required"
            });
        }

        try
        {
            // Check if Graph client is properly initialized
            if (_graphServiceClient == null)
            {
                return JsonSerializer.Serialize(new
                {
                    error = "Microsoft Graph client not configured",
                    message = "Authentication needs to be properly configured to access Microsoft Graph API"
                });
            }

            // Execute the Graph API request using modern syntax
            var user = await _graphServiceClient.Users[userId].GetAsync(requestConfiguration =>
            {
                if (!string.IsNullOrWhiteSpace(select))
                {
                    requestConfiguration.QueryParameters.Select = select.Split(',', StringSplitOptions.RemoveEmptyEntries);
                }
            });

            if (user == null)
            {
                return JsonSerializer.Serialize(new
                {
                    error = "User not found",
                    message = $"No user found with ID or UPN: {userId}"
                });
            }

            // Format the response
            var result = new
            {
                id = user.Id,
                displayName = user.DisplayName,
                mail = user.Mail,
                userPrincipalName = user.UserPrincipalName,
                department = user.Department,
                jobTitle = user.JobTitle,
                officeLocation = user.OfficeLocation,
                mobilePhone = user.MobilePhone,
                businessPhones = user.BusinessPhones,
                accountEnabled = user.AccountEnabled,
                createdDateTime = user.CreatedDateTime,
                lastSignInDateTime = user.SignInActivity?.LastSignInDateTime,
                city = user.City,
                country = user.Country,
                companyName = user.CompanyName,
                selectedProperties = select
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (ServiceException ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Microsoft Graph API error",
                code = ex.ResponseStatusCode,
                message = ex.Message,
                details = ex.ToString()
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Unexpected error",
                message = ex.Message,
                type = ex.GetType().Name
            });
        }
    }

    [McpServerTool]
    [Description("Gets configuration information needed to authenticate with Microsoft Graph API.")]
    public string GetGraphAuthenticationInfo()
    {
        var authInfo = new
        {
            message = "To use Microsoft Graph tools, you need to configure authentication",
            steps = new[]
            {
                "1. Register an application in Azure Active Directory",
                "2. Grant appropriate permissions (User.Read.All, User.ReadBasic.All, etc.)",
                "3. Create a client secret or certificate",
                "4. Configure the GraphServiceClient with proper authentication provider"
            },
            requiredPermissions = new[]
            {
                "User.Read.All - Read all users' full profiles",
                "User.ReadBasic.All - Read all users' basic profiles",
                "User.Read - Read signed-in user's profile",
                "Directory.Read.All - Read directory data"
            },
            codeExample = @"
// Example of how to initialize GraphServiceClient with ClientSecretCredential
var credential = new ClientSecretCredential(
    tenantId: ""your-tenant-id"",
    clientId: ""your-client-id"",
    clientSecret: ""your-client-secret""
);

var graphServiceClient = new GraphServiceClient(credential);
"
        };

        return JsonSerializer.Serialize(authInfo, new JsonSerializerOptions { WriteIndented = true });
    }
}