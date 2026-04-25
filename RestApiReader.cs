using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;

namespace SpListReader;

/// <summary>
/// 방법 A: MSAL + HttpClient로 SharePoint REST API 직접 호출
/// </summary>
public static class RestApiReader
{
    public static async Task RunAsync(IConfiguration config)
    {
        var tenantId = config["TenantId"]!;
        var clientId = config["ClientId"]!;
        var siteUrl = config["SiteUrl"]!;
        var listTitle = config["ListTitle"]!;
        var certPath = Path.Combine(AppContext.BaseDirectory, config["CertPath"]!);
        var certPassword = config["CertPassword"]!;

        // 1. 인증서 로드
        var cert = new X509Certificate2(certPath, certPassword);

        // 2. MSAL로 토큰 발급
        var app = ConfidentialClientApplicationBuilder.Create(clientId)
            .WithCertificate(cert)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .Build();

        var host = new Uri(siteUrl).GetLeftPart(UriPartial.Authority);
        var result = await app.AcquireTokenForClient(new[] { $"{host}/.default" }).ExecuteAsync();

        Console.WriteLine("[REST API] 토큰 발급 성공");

        // 3. SharePoint REST API 호출
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
        http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var apiUrl = $"{siteUrl}/_api/web/lists/getByTitle('{listTitle}')/items?$top=50";
        var response = await http.GetAsync(apiUrl);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);
        var items = doc.RootElement.GetProperty("value");

        Console.WriteLine($"[REST API] {items.GetArrayLength()}건 조회됨\n");
        Console.WriteLine($"{"ID",-6} {"Title"}");
        Console.WriteLine(new string('-', 40));

        foreach (var item in items.EnumerateArray())
        {
            var id = item.GetProperty("Id").GetInt32();
            var title = item.TryGetProperty("Title", out var t) ? t.GetString() ?? "(null)" : "(no Title)";
            Console.WriteLine($"{id,-6} {title}");
        }
    }
}
