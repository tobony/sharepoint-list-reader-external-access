using Microsoft.Extensions.Configuration;
using SpListReader;

var config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json")
    .Build();

var method = args.Length > 0 ? args[0].ToLower() : "";

try
{
    switch (method)
    {
        case "restapi":
            await RestApiReader.RunAsync(config);
            break;
        case "pnp":
            await PnpCsomReader.RunAsync(config);
            break;
        default:
            Console.WriteLine("사용법: dotnet run -- restapi   (방법 A: REST API 직접 호출)");
            Console.WriteLine("        dotnet run -- pnp       (방법 B: PnP Framework CSOM)");
            break;
    }
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"\n오류: {ex.Message}");
    if (ex.InnerException != null)
        Console.WriteLine($"내부: {ex.InnerException.Message}");
    Console.ResetColor();
}
