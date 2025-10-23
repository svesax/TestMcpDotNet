# Microsoft Graph Tools Configuration

This MCP server includes Microsoft Graph tools that allow you to interact with Microsoft 365 user data. To use these tools, you need to configure authentication with Azure Active Directory.

## Tools Available

### 1. GetUsers
Gets a list of users from Microsoft Graph with optional filtering.

**Parameters:**
- `filter` (optional): OData filter expression (e.g., "startswith(displayName,'John')" or "department eq 'Sales'")
- `search` (optional): Search query (e.g., "John Smith" or "john@contoso.com")
- `top` (optional): Maximum number of users to return (default: 10, max: 100)
- `select` (optional): Comma-separated list of properties to select (e.g., "displayName,mail,department")

**Example filters:**
```
startswith(displayName,'John')
department eq 'Sales'
accountEnabled eq true
startswith(mail,'admin')
```

### 2. GetUserById
Gets detailed information about a specific user by their ID or User Principal Name.

**Parameters:**
- `userId`: User ID (GUID) or User Principal Name (email)
- `select` (optional): Comma-separated list of properties to select

### 3. GetGraphAuthenticationInfo
Returns information about how to configure authentication for the Microsoft Graph API.

## Authentication Setup

To enable these tools, you need to:

1. **Register an Azure AD Application:**
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Click "New registration"
   - Provide a name and select account types
   - Note the Application (client) ID and Directory (tenant) ID

2. **Configure Permissions:**
   Grant the following API permissions:
   - `User.Read.All` - Read all users' full profiles
   - `User.ReadBasic.All` - Read all users' basic profiles
   - `Directory.Read.All` - Read directory data

3. **Create Authentication Credentials:**
   - For app-only access: Create a client secret
   - For delegated access: Configure redirect URIs

4. **Update the Code:**
   Modify the `MicrosoftGraphTools` constructor to initialize the `GraphServiceClient` with proper authentication:

```csharp
public MicrosoftGraphTools()
{
    var credential = new ClientSecretCredential(
        tenantId: "your-tenant-id",
        clientId: "your-client-id", 
        clientSecret: "your-client-secret"
    );
    
    _graphServiceClient = new GraphServiceClient(credential);
}
```

## Security Considerations

- Store credentials securely (use Azure Key Vault, environment variables, or secure configuration)
- Follow the principle of least privilege when granting permissions
- Regularly rotate client secrets
- Consider using certificates instead of client secrets for production scenarios
- Implement proper error handling and logging

## Example Usage

Once configured, you can use the tools like:

```json
{
  "method": "tools/call",
  "params": {
    "name": "GetUsers",
    "arguments": {
      "filter": "department eq 'Engineering'",
      "top": 20,
      "select": "displayName,mail,jobTitle"
    }
  }
}
```

## Troubleshooting

- **"Microsoft Graph client not configured"**: Authentication setup is incomplete
- **403 Forbidden**: Insufficient permissions granted to the application
- **401 Unauthorized**: Invalid credentials or expired tokens
- **400 Bad Request**: Invalid filter syntax or unsupported properties in select

For more information, visit the [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/).