using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Azure.Identity;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Server.Kestrel.Core;


var builder = WebApplication.CreateBuilder(args);

// Configure all logs to go to stderr (stdout is used for the MCP protocol messages).
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

// Configure the server to listen on port 5000
/*
builder.Services.Configure<KestrelServerOptions>(options =>
{
    options.ListenLocalhost(5000);
});
*/

// Configure Azure AD app settings
var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
var redirectUri = Environment.GetEnvironmentVariable("redirectUri");

// Configure Interactive Browser Credential for delegated permissions
var options = new InteractiveBrowserCredentialOptions
{
    ClientId = clientId,
    TenantId = tenantId,
    RedirectUri = new Uri(redirectUri)
};

var credential = new InteractiveBrowserCredential(options);

// Register GraphServiceClient with delegated permissions
builder.Services.AddSingleton<GraphServiceClient>(serviceProvider =>
{
    return new GraphServiceClient(credential);
});

// Add the MCP services: the transport to use (HTTP) and the tools to register.
builder.Services
    .AddMcpServer()
    .WithHttpTransport()
    .WithTools<RandomNumberTools>()
    .WithTools<MicrosoftGraphTools>();

var app = builder.Build();
app.MapMcp();
//app.MapMcp("/mcp");

await app.RunAsync();
