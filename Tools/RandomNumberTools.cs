using System.ComponentModel;
using ModelContextProtocol.Server;

/// <summary>
/// Sample MCP tools for demonstration purposes.
/// These tools can be invoked by MCP clients to perform various operations.
/// </summary>
internal class RandomNumberTools
{
    [McpServerTool]
    [Description("Generates a random number between the specified minimum and maximum values.")]
    public int GetRandomNumber(
        [Description("Minimum value (inclusive)")] int min = 0,
        [Description("Maximum value (exclusive)")] int max = 100)
    {
        return Random.Shared.Next(min, max);
    }

    [McpServerTool]
    [Description("Get representing an Italian city.")]
    public string GetItalyCities(
        [Description("Indicate if want a city in north or south")] string northOrSouth = "north")
    {
        if (northOrSouth == "north")
        {
            return "Milan"; // North cities
        }
        else
        {
            return "Naples"; // South cities
        }
    }
}
