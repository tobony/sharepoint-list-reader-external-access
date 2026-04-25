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
Write-Host "`n다음 단계: CER 파일을 Entra ID 앱 등록 > Certificates & secrets에 업로드하세요." -ForegroundColor Yellow
