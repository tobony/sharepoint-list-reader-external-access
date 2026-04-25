using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using PnP.Framework;

namespace SpListReader;

/// <summary>
/// 방법 B: PnP Framework (CSOM 래퍼) 사용
/// </summary>
public static class PnpCsomReader
{
    public static async Task RunAsync(IConfiguration config)
    {
        var tenantId = config["TenantId"]!;
        var clientId = config["ClientId"]!;
        var siteUrl = config["SiteUrl"]!;
        var listTitle = config["ListTitle"]!;
        var certPath = Path.Combine(AppContext.BaseDirectory, config["CertPath"]!);
        var certPassword = config["CertPassword"]!;

        // 1. PnP AuthenticationManager로 인증 (인증서 기반)
        var authManager = new AuthenticationManager(clientId, certPath, certPassword, $"{tenantId}");

        using var ctx = await authManager.GetContextAsync(siteUrl);

        Console.WriteLine("[PnP CSOM] 인증 성공");

        // 2. List 아이템 조회
        var list = ctx.Web.Lists.GetByTitle(listTitle);
        var query = CamlQuery.CreateAllItemsQuery(50);
        var items = list.GetItems(query);

        ctx.Load(items);
        await ctx.ExecuteQueryAsync();

        Console.WriteLine($"[PnP CSOM] {items.Count}건 조회됨\n");
        Console.WriteLine($"{"ID",-6} {"Title"}");
        Console.WriteLine(new string('-', 40));

        foreach (var item in items)
        {
            Console.WriteLine($"{item.Id,-6} {item["Title"] ?? "(null)"}");
        }
    }
}
