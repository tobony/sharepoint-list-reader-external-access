# Entra ID 앱 등록 절차 (SharePoint REST API App-Only 인증)

## 사전 준비
- 인증서 파일: `SPListReaderCert.cer` (`Create-Certificate.ps1`로 생성)
- Entra ID 관리자 권한 필요

---

## 1단계: 앱 등록

1. https://entra.microsoft.com 접속
2. 좌측 메뉴 → **Applications** → **App registrations**
3. **+ New registration** 클릭
4. 설정:
   - **Name**: `SPListReader`
   - **Supported account types**: `Accounts in this organizational directory only` (단일 테넌트)
   - **Redirect URI**: 비워두기 (콘솔 앱이므로 불필요)
5. **Register** 클릭

## 2단계: Client ID / Tenant ID 기록

앱 등록 후 **Overview** 페이지에서:
- **Application (client) ID** → 복사하여 기록
- **Directory (tenant) ID** → 복사하여 기록

## 3단계: 인증서 업로드

1. 좌측 메뉴 → **Certificates & secrets**
2. **Certificates** 탭 클릭
3. **Upload certificate** 클릭
4. `SPListReaderCert.cer` 파일 선택 → **Add**
5. 업로드된 인증서의 Thumbprint가 `Create-Certificate.ps1` 실행 시 출력된 값과 일치하는지 확인

## 4단계: SharePoint API 권한 추가

1. 좌측 메뉴 → **API permissions**
2. **+ Add a permission** 클릭
3. **SharePoint** 선택 (Microsoft Graph가 아님!)
4. **Application permissions** 선택
5. `Sites.Read.All` 체크 → **Add permissions**
6. 상단의 **✓ Grant admin consent for [테넌트명]** 클릭 → **Yes**
7. Status가 모두 **Granted for [테넌트명]** 으로 표시되는지 확인

---

## 완료 후 기록할 값

`appsettings.template.json`을 `appsettings.json`으로 복사한 뒤 아래 값을 입력합니다:

```
TenantId     = (2단계에서 복사한 Directory (tenant) ID)
ClientId     = (2단계에서 복사한 Application (client) ID)
SiteUrl      = https://{테넌트}.sharepoint.com/sites/{사이트명}
ListTitle    = {리스트 이름}
CertPath     = SPListReaderCert.pfx
CertPassword = (Create-Certificate.ps1 실행 시 입력한 비밀번호)
```
