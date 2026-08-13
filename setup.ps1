# 【 배포 사이트 자동 설정 스크립트 】
# 목적: 배포 사이트의 app.js를 깃허브 Raw 링크로 업데이트

$appJsPath = "C:\Users\1\Documents\매핑점검\site_publish\public\mapping\app.js"
$githubRawUrl = "https://raw.githubusercontent.com/rkwlcip-oss/inbound-automation/main/입고확인서/inbound_v3_1%20(2).html"

Write-Host "【 배포 사이트 설정 중... 】" -ForegroundColor Cyan

if (-not (Test-Path $appJsPath)) {
    Write-Error "❌ 파일을 찾을 수 없습니다: $appJsPath"
    exit 1
}

$content = Get-Content -Path $appJsPath -Raw -Encoding UTF8
$originalSize = $content.Length

# 모든 로컬 경로를 깃허브 링크로 변경
$patterns = @(
    'C:/Users/1/Documents/입고확인서/inbound_v3_1 \(2\)\.html',
    'C:\\Users\\1\\Documents\\입고확인서\\inbound_v3_1 \(2\)\.html',
    'inbound_v3_1 \(2\)\.html'
)

foreach ($pattern in $patterns) {
    if ($content -match $pattern) {
        $content = $content -replace $pattern, $githubRawUrl
        Write-Host "✅ '$pattern' → 깃허브 링크로 변경"
    }
}

Set-Content -Path $appJsPath -Value $content -Encoding UTF8
$newSize = (Get-Content -Path $appJsPath -Raw -Encoding UTF8).Length

Write-Host "" 
Write-Host "✅ app.js 업데이트 완료!" -ForegroundColor Green
Write-Host "📍 파일 크기: $originalSize → $newSize bytes"
Write-Host "🔗 사용 중인 링크: $githubRawUrl"
Write-Host ""
Write-Host "【 다음 단계 】" -ForegroundColor Cyan
Write-Host "1. 배포 사이트 새로고침 (Ctrl+Shift+R)"
Write-Host "2. 다음 3가지 확인:"
Write-Host "   ✓ 파일 업로드 기준 일시 → 메타데이터 시각 표시"
Write-Host "   ✓ 출고저조 → '개발중' 표시"
Write-Host "   ✓ 품절 상세 → 비고(확정일-나리주임) 표시"

