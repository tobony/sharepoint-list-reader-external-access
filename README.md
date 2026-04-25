# SharePoint Add-in 종료 후, 외부에서 SharePoint List 데이터 가져오기

> **작성일**: 2026-04-25  
> **환경**: Windows, .NET 8.0, SharePoint Online  
> **키워드**: SharePoint Add-in 종료, SPFx, Entra ID, 인증서 인증, REST API, PnP Framework, CSOM

---

## 목차

1. [배경: SharePoint Add-in 모델 종료](#1-배경-sharepoint-add-in-모델-종료)
2. [대안 분석: 외부에서 SharePoint List에 접근하는 방법](#2-대안-분석-외부에서-sharepoint-list에-접근하는-방법)
3. [사전 준비](#3-사전-준비)
4. [Step 1: 자체 서명 인증서 생성](#4-step-1-자체-서명-인증서-생성)
5. [Step 2: Entra ID 앱 등록](#5-step-2-entra-id-앱-등록)
6. [Step 3: C# 프로젝트 구성](#6-step-3-c-프로젝트-구성)
7. [Step 4: 방법 A - MSAL + HttpClient로 REST API 직접 호출](#7-step-4-방법-a---msal--httpclient로-rest-api-직접-호출)
8. [Step 5: 방법 B - PnP Framework (CSOM) 사용](#8-step-5-방법-b---pnp-framework-csom-사용)
9. [Step 6: 통합 진입점 (Program.cs)](#9-step-6-통합-진입점-programcs)
10. [실행 결과](#10-실행-결과)
11. [두 방법 비교](#11-두-방법-비교)
12. [주의사항 및 트러블슈팅](#12-주의사항-및-트러블슈팅)
13. [참고 링크](#13-참고-링크)

---

## 1. 배경: SharePoint Add-in 모델 종료

Microsoft는 SharePoint Add-in 확장 모델을 단계적으로 종료했습니다.

| 일자 | 내용 |
|---|---|
| 2023-11-27 | SharePoint Add-in 모델 및 Azure ACS 공식 deprecated |
| 2024-11-01 | **신규 테넌트**에서 Add-in 작동 중지 |
| **2026-04-02** | **모든 테넌트에서 완전 종료 (End-of-life)** |

함께 종료된 것:
- **Azure ACS (Access Control Service)**: Add-in이 사용하던 인증 방식도 동시에 종료
- 기존에 `client_id` + `client_secret`으로 토큰을 발급받던 방식이 더 이상 작동하지 않음

공식 대체 기술은 **SharePoint Framework (SPFx)** 이지만, SPFx는 SharePoint 페이지 내부에서 동작하는 웹파트/확장 모델입니다. **외부 앱에서 SharePoint 데이터에 접근**하려면 별도의 방법이 필요합니다.

> 📌 공식 문서: [SharePoint Add-In retirement in Microsoft 365](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/retirement-announcement-for-add-ins)

---

## 2. 대안 분석: 외부에서 SharePoint List에 접근하는 방법

| 방법 | 설명 | 제약 |
|---|---|---|
| **Microsoft Graph API** | Microsoft 365 통합 API | 회사 정책으로 차단될 수 있음 |
| **Entra ID + 인증서 + SharePoint REST API** | SharePoint 자체 REST 엔드포인트 직접 호출 | ✅ **본 글에서 사용하는 방법** |
| **Entra ID + 인증서 + PnP Framework (CSOM)** | CSOM 래퍼 라이브러리 사용 | ✅ **본 글에서 사용하는 방법** |
| Power Automate | 노코드 자동화 | 프로그래밍 방식이 아님 |
| PnP PowerShell | PowerShell 기반 | 스크립트 전용 |

**Graph API를 사용할 수 없는 환경**에서는 **Entra ID 앱 등록 + 인증서 인증 + SharePoint REST API**가 가장 현실적인 대안입니다. Graph API 엔드포인트(`graph.microsoft.com`)를 사용하지 않고, SharePoint 자체 REST 엔드포인트(`{tenant}.sharepoint.com/_api/...`)를 직접 호출합니다.

### 핵심 포인트

- Azure ACS가 종료되었으므로, 인증은 반드시 **Entra ID (구 Azure AD)** 를 통해야 합니다
- App-Only(무인 실행) 방식에서는 **반드시 인증서**를 사용해야 합니다 (client secret 불가)
- 토큰 scope는 `https://{tenant}.sharepoint.com/.default` 형식입니다

---

## 3. 사전 준비

- **OS**: Windows 10/11 (PowerShell 5.1 이상)
- **.NET SDK**: 8.0 이상 ([다운로드](https://dotnet.microsoft.com/download))
- **Entra ID 관리자 권한**: 앱 등록 및 admin consent 가능해야 함
- **SharePoint Online 사이트**: 접근할 List가 있는 사이트

본 글에서 사용하는 SharePoint List 정보:
- 사이트: `https://contoso.sharepoint.com/sites/my_site`
- 리스트: `site_list_issue_tracker`

### 최종 프로젝트 구조

```
SpListReader/
├── Create-Certificate.ps1      ← 인증서 생성 스크립트
├── SPListReaderCert.pfx        ← 인증서 (개인키 포함 - 앱에서 사용)
├── SPListReaderCert.cer        ← 인증서 (공개키만 - Entra ID에 업로드)
├── appsettings.json            ← 설정 파일
├── SpListReader.csproj         ← 프로젝트 파일
├── Program.cs                  ← 진입점
├── RestApiReader.cs            ← 방법 A: REST API 직접 호출
└── PnpCsomReader.cs            ← 방법 B: PnP Framework CSOM
```


---

## 4. Step 1: 자체 서명 인증서 생성

Entra ID App-Only 인증에는 **반드시 인증서**가 필요합니다. 테스트 환경에서는 자체 서명 인증서를 사용합니다.

### PowerShell 스크립트 (`Create-Certificate.ps1`)

```powershell
# 자체 서명 인증서 생성 스크립트
# SharePoint REST API App-Only 인증용

$certName = "SPListReaderCert"
$certDir  = $PSScriptRoot
$pfxPath  = Join-Path $certDir "$certName.pfx"
$cerPath  = Join-Path $certDir "$certName.cer"
$passwordInput = Read-Host -Prompt "PFX 비밀번호를 입력하세요" -AsSecureString
$password = $passwordInput

# 1. 자체 서명 인증서 생성 (유효기간 2년)
$cert = New-SelfSignedCertificate `
    -Subject "CN=$certName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

Write-Host "인증서 생성 완료 - Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

# 2. PFX 내보내기 (개인키 포함 - 앱에서 사용)
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null
Write-Host "PFX 내보내기 완료: $pfxPath" -ForegroundColor Green

# 3. CER 내보내기 (공개키만 - Entra ID에 업로드)
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Write-Host "CER 내보내기 완료: $cerPath" -ForegroundColor Green

# 4. 결과 요약
Write-Host "`n========== 인증서 정보 ==========" -ForegroundColor Cyan
Write-Host "Thumbprint : $($cert.Thumbprint)"
Write-Host "Subject    : $($cert.Subject)"
Write-Host "NotAfter   : $($cert.NotAfter)"
Write-Host "PFX 파일   : $pfxPath"
Write-Host "CER 파일   : $cerPath"
Write-Host "PFX 비밀번호: (실행 시 입력한 값)"
Write-Host "=================================" -ForegroundColor Cyan
```

### 실행

```bash
powershell -ExecutionPolicy Bypass -File Create-Certificate.ps1
```

### 실행 결과

```
인증서 생성 완료 - Thumbprint: A1B2C3D4E5F6...
PFX 내보내기 완료: C:\...\SpListReader\SPListReaderCert.pfx
CER 내보내기 완료: C:\...\SpListReader\SPListReaderCert.cer

========== 인증서 정보 ==========
Thumbprint : A1B2C3D4E5F6...
Subject    : CN=SPListReaderCert
NotAfter   : 04/25/2028 12:54:42
================================
```

### 생성되는 파일

| 파일 | 용도 | 포함 내용 |
|---|---|---|
| `SPListReaderCert.pfx` | C# 앱에서 인증에 사용 | 개인키 + 공개키 |
| `SPListReaderCert.cer` | Entra ID 앱에 업로드 | 공개키만 |

> ⚠️ **주의**: `.pfx` 파일은 개인키를 포함하므로 절대 외부에 노출하지 마세요. Git에 커밋하지 않도록 `.gitignore`에 추가하세요.


---

## 5. Step 2: Entra ID 앱 등록

이 단계는 Azure Portal에서 수동으로 진행합니다.

### 5-1. 앱 등록 생성

1. [https://entra.microsoft.com](https://entra.microsoft.com) 접속
2. 좌측 메뉴 → **Applications** → **App registrations**
3. **+ New registration** 클릭
4. 설정:
   - **Name**: `SPListReader`
   - **Supported account types**: `Accounts in this organizational directory only` (단일 테넌트)
   - **Redirect URI**: 비워두기 (콘솔 앱이므로 불필요)
5. **Register** 클릭

### 5-2. Client ID / Tenant ID 기록

앱 등록 후 **Overview** 페이지에서 두 값을 복사합니다:

- **Application (client) ID** → 예: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- **Directory (tenant) ID** → 예: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

> 📌 이 두 값은 나중에 `appsettings.json`에 입력합니다.

### 5-3. 인증서 업로드

1. 좌측 메뉴 → **Certificates & secrets**
2. **Certificates** 탭 클릭
3. **Upload certificate** 클릭
4. Step 1에서 생성한 `SPListReaderCert.cer` 파일 선택 → **Add**
5. 업로드된 인증서의 Thumbprint가 Step 1의 결과와 일치하는지 확인

> ⚠️ **주의**: `.cer` 파일(공개키만)을 업로드해야 합니다. `.pfx` 파일이 아닙니다!

### 5-4. SharePoint API 권한 추가

1. 좌측 메뉴 → **API permissions**
2. **+ Add a permission** 클릭
3. **SharePoint** 선택 (**Microsoft Graph가 아닙니다!**)
4. **Application permissions** 선택
5. `Sites.Read.All` 체크 → **Add permissions**
6. 상단의 **✓ Grant admin consent for [테넌트명]** 클릭 → **Yes**
7. Status가 모두 **Granted for [테넌트명]** 으로 표시되는지 확인

> ⚠️ **주의**: 반드시 **Application permissions**를 선택해야 합니다. Delegated permissions가 아닙니다. App-Only(무인 실행) 방식에서는 사용자 로그인 없이 앱 자체 권한으로 동작하기 때문입니다.

> ⚠️ **주의**: **Grant admin consent** 버튼을 반드시 클릭해야 합니다. 클릭하지 않으면 권한이 부여되지 않아 401 Unauthorized 오류가 발생합니다.

### 5-5. 권한 전파 대기

Admin consent 후 실제로 권한이 적용되기까지 **최대 수 분**이 소요될 수 있습니다. 즉시 실행 시 401 오류가 발생하면 1~2분 후 재시도하세요.


---

## 6. Step 3: C# 프로젝트 구성

### 프로젝트 생성

```bash
dotnet new console -n SpListReader
cd SpListReader
```

### NuGet 패키지 설치

```bash
dotnet add package Microsoft.Identity.Client --version 4.70.2
dotnet add package PnP.Framework --version 1.18.0
dotnet add package Microsoft.Extensions.Configuration.Json --version 8.0.1
```

| 패키지 | 용도 |
|---|---|
| `Microsoft.Identity.Client` (MSAL) | Entra ID 토큰 발급 |
| `PnP.Framework` | 방법 B에서 CSOM 래퍼로 사용 |
| `Microsoft.Extensions.Configuration.Json` | appsettings.json 읽기 |

> ⚠️ **주의**: `PnP.Framework`가 의존하는 MSAL 버전과 직접 설치하는 MSAL 버전이 일치해야 합니다. PnP.Framework 1.18.0은 MSAL 4.70.2를 요구하므로, 반드시 `4.70.2`를 지정하세요. 버전이 다르면 `NU1605` 다운그레이드 오류가 발생합니다.

### 프로젝트 파일 (`SpListReader.csproj`)

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="8.0.1" />
    <PackageReference Include="Microsoft.Identity.Client" Version="4.70.2" />
    <PackageReference Include="PnP.Framework" Version="1.18.0" />
  </ItemGroup>

  <ItemGroup>
    <None Update="appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="SPListReaderCert.pfx">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>

</Project>
```

> 📌 `appsettings.json`과 `SPListReaderCert.pfx`가 빌드 출력 디렉토리에 복사되도록 `CopyToOutputDirectory`를 설정합니다.

### 설정 파일 (`appsettings.json`)

```json
{
  "TenantId": "YOUR_TENANT_ID",
  "ClientId": "YOUR_CLIENT_ID",
  "CertPath": "SPListReaderCert.pfx",
  "CertPassword": "YOUR_CERT_PASSWORD",
  "SiteUrl": "https://contoso.sharepoint.com/sites/my_site",
  "ListTitle": "site_list_issue_tracker"
}
```

Step 2에서 기록한 Tenant ID와 Client ID를 입력합니다.

> ⚠️ **주의**: `appsettings.json`에 인증서 비밀번호가 포함되어 있으므로, 프로덕션 환경에서는 Azure Key Vault 등 안전한 저장소를 사용하세요.


---

## 7. Step 4: 방법 A - MSAL + HttpClient로 REST API 직접 호출

SPFx 내부의 `SPHttpClient`에 대응하는 외부 버전입니다. MSAL로 토큰을 발급받고, `HttpClient`로 SharePoint REST API를 직접 호출합니다.

### 인증 흐름

```
C# App → MSAL (인증서로 서명) → Entra ID → Access Token 발급
C# App → HttpClient + Bearer Token → SharePoint REST API (/_api/web/lists/...)
```

### 코드 (`RestApiReader.cs`)

```csharp
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
        var result = await app.AcquireTokenForClient(
            new[] { $"{host}/.default" }).ExecuteAsync();

        Console.WriteLine("[REST API] 토큰 발급 성공");

        // 3. SharePoint REST API 호출
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", result.AccessToken);
        http.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

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
            var title = item.TryGetProperty("Title", out var t)
                ? t.GetString() ?? "(null)" : "(no Title)";
            Console.WriteLine($"{id,-6} {title}");
        }
    }
}
```

### 핵심 포인트

- **토큰 scope**: `https://{tenant}.sharepoint.com/.default` — Graph API가 아닌 SharePoint 자체 리소스에 대한 토큰
- **REST API URL**: `/_api/web/lists/getByTitle('리스트명')/items` — SharePoint 자체 REST 엔드포인트
- **`$top=50`**: 한 번에 가져올 아이템 수 제한. 기본값은 100이며, 최대 5000까지 가능
- 응답은 JSON 형식이며, `value` 배열에 아이템이 담겨 있음


---

## 8. Step 5: 방법 B - PnP Framework (CSOM) 사용

SPFx 내부의 `PnPJS`에 대응하는 외부 버전입니다. PnP Framework 라이브러리가 인증과 CSOM 호출을 래핑해주므로 코드가 더 간결합니다.

### 인증 흐름

```
C# App → PnP AuthenticationManager (내부적으로 MSAL 사용) → Entra ID → Access Token
PnP AuthenticationManager → ClientContext 생성 → CSOM으로 SharePoint 호출
```

### 코드 (`PnpCsomReader.cs`)

```csharp
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
        var authManager = new AuthenticationManager(
            clientId, certPath, certPassword, $"{tenantId}");

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
```

### 핵심 포인트

- `AuthenticationManager`가 MSAL 토큰 발급을 내부적으로 처리해줌
- `GetContextAsync()`로 인증된 `ClientContext`를 바로 획득
- `CamlQuery.CreateAllItemsQuery(50)`으로 최대 50건 조회
- CSOM은 `Load()` → `ExecuteQueryAsync()` 패턴으로 서버에 요청을 보냄 (배치 처리)


---

## 9. Step 6: 통합 진입점 (Program.cs)

커맨드라인 인자로 두 방법을 선택 실행할 수 있도록 통합합니다.

### 코드 (`Program.cs`)

```csharp
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
```

---

## 10. 실행 결과

### 빌드

```bash
dotnet build
```

```
빌드했습니다.
    경고 0개
    오류 0개
```

### 방법 A 실행

```bash
dotnet run -- restapi
```

```
[REST API] 토큰 발급 성공
[REST API] 3건 조회됨

ID     Title
----------------------------------------
1      서피스프로2 출시일 지연
2      Power Automate 자동화 10개만들기
3      Wipe백 조사하기
```

### 방법 B 실행

```bash
dotnet run -- pnp
```

```
[PnP CSOM] 인증 성공
[PnP CSOM] 3건 조회됨

ID     Title
----------------------------------------
1      서피스프로2 출시일 지연
2      Power Automate 자동화 10개만들기
3      Wipe백 조사하기
```

두 방법 모두 동일한 결과를 반환합니다.

---

## 11. 두 방법 비교

| | 방법 A: REST API 직접 호출 | 방법 B: PnP Framework CSOM |
|---|---|---|
| **SPFx 대응** | `SPHttpClient` | `PnPJS` |
| **인증 처리** | MSAL 직접 사용 | PnP AuthenticationManager (내부 MSAL) |
| **데이터 호출** | `HttpClient` + REST URL | CSOM `ClientContext` |
| **응답 형식** | JSON (직접 파싱) | 강타입 객체 (`ListItem`) |
| **의존성** | `Microsoft.Identity.Client` | `PnP.Framework` (+ 다수 하위 의존성) |
| **코드량** | 상대적으로 많음 | 간결함 |
| **유연성** | OData 쿼리 자유롭게 사용 가능 | CAML 쿼리 사용 |
| **적합한 경우** | REST API에 익숙하거나, 세밀한 제어가 필요할 때 | 빠르게 구현하고 싶을 때 |

### 어떤 방법을 선택해야 할까?

- **REST API 직접 호출 (방법 A)**: 의존성을 최소화하고 싶거나, OData 쿼리(`$select`, `$filter`, `$expand` 등)를 자유롭게 사용하고 싶을 때
- **PnP Framework (방법 B)**: 빠르게 구현하고 싶거나, CSOM에 익숙한 기존 SharePoint 개발자일 때


---

## 12. 주의사항 및 트러블슈팅

### 보안 관련

| 항목 | 주의사항 |
|---|---|
| **PFX 파일** | 개인키를 포함하므로 절대 Git에 커밋하지 마세요. `.gitignore`에 `*.pfx` 추가 |
| **appsettings.json** | 인증서 비밀번호가 포함되어 있으므로 프로덕션에서는 Azure Key Vault 사용 권장 |
| **인증서 유효기간** | 자체 서명 인증서는 만료일이 있습니다. 만료 전에 갱신 필요 |
| **Sites.Read.All 권한** | 테넌트 내 **모든 사이트**를 읽을 수 있는 강력한 권한입니다. 최소 권한 원칙에 따라 `Sites.Selected`를 고려하세요 |
| **Client Secret** | SharePoint REST API App-Only 인증에서는 **client secret을 사용할 수 없습니다**. 반드시 인증서를 사용해야 합니다 |

### 자주 발생하는 오류

#### 1. `AADSTS700027: The certificate with identifier used to sign the client assertion is not registered`

**원인**: Entra ID 앱에 인증서가 업로드되지 않았거나, 다른 인증서를 업로드함

**해결**:
- Entra Portal → App registrations → Certificates & secrets에서 `.cer` 파일이 업로드되었는지 확인
- 업로드된 인증서의 Thumbprint가 `.pfx` 파일의 Thumbprint와 일치하는지 확인

#### 2. `401 Unauthorized`

**원인**: API 권한이 부여되지 않았거나, admin consent가 완료되지 않았거나, 권한 전파 대기 중

**해결**:
- API permissions에서 `SharePoint > Sites.Read.All` (Application)이 있는지 확인
- Status가 "Granted"인지 확인. "Not granted"이면 **Grant admin consent** 클릭
- 권한 부여 직후라면 1~2분 대기 후 재시도

#### 3. `NU1605: 다운그레이드된 패키지 Microsoft.Identity.Client 발견`

**원인**: MSAL 버전과 PnP.Framework가 요구하는 MSAL 버전 불일치

**해결**:
- `PnP.Framework`가 요구하는 MSAL 버전을 확인하고 동일한 버전을 설치
- 현재 PnP.Framework 1.18.0은 MSAL 4.70.2를 요구함

#### 4. `AADSTS900023: Specified tenant identifier is neither a valid DNS name`

**원인**: `appsettings.json`의 `TenantId`가 플레이스홀더(`YOUR_TENANT_ID`) 상태

**해결**: Entra Portal에서 실제 Directory (tenant) ID를 복사하여 입력

#### 5. `403 Forbidden`

**원인**: 해당 사이트에 대한 접근 권한이 없음

**해결**:
- `Sites.Read.All`은 모든 사이트에 대한 읽기 권한이므로 일반적으로 발생하지 않음
- `Sites.Selected`를 사용하는 경우, 특정 사이트에 대한 권한을 별도로 부여해야 함

### 프로덕션 환경 권장사항

1. **Azure Key Vault**에 인증서와 비밀번호를 저장하고, Managed Identity로 접근
2. **Sites.Selected** 권한을 사용하여 특정 사이트에만 접근 허용 (최소 권한 원칙)
3. 인증서 만료 알림 설정 (Entra ID에서 자동 알림 가능)
4. 토큰 캐싱 구현 (MSAL은 기본적으로 인메모리 캐시를 제공하지만, 분산 환경에서는 Redis 등 사용)

---

## 13. 참고 링크

- [SharePoint Add-In retirement in Microsoft 365](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/retirement-announcement-for-add-ins)
- [SharePoint Framework (SPFx) Overview](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Granting access via Entra ID App-Only (인증서 기반)](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread)
- [SharePoint Add-In and Azure ACS retirement FAQ](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/add-ins-and-azure-acs-retirements-faq)
- [Modernization guidance for SharePoint Add-In model](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/sp-add-in-modernize)
- [PnP Framework GitHub](https://github.com/pnp/pnpframework)
- [SharePoint REST API 가이드](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [SPFx Roadmap](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/roadmap)
