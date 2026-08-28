[CmdletBinding()]
param(
    [switch]$Prune,
    [switch]$Help
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

function Invoke-GitChecked {
    param([Parameter(Mandatory)][string[]]$Arguments)

    & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ') 명령이 실패했습니다."
    }
}

if ($Help) {
    Write-Host '사용법: .\pull.bat [-Prune] [-Help]'
    Write-Host '  기본값: 현재 브랜치의 upstream을 fast-forward only 방식으로 pull'
    Write-Host '  -Prune: pull하면서 원격에서 삭제된 추적 브랜치도 정리'
    Write-Host '  -Help: 이 도움말 표시'
    exit 0
}

try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'Git을 찾을 수 없습니다. Git for Windows를 먼저 설치하세요.'
    }

    Set-Location -LiteralPath $repositoryRoot
    if (-not (Test-Path -LiteralPath (Join-Path $repositoryRoot '.git'))) {
        throw '이 폴더는 Git clone 저장소가 아닙니다. ZIP 폴더에서는 pull할 수 없으므로 git clone으로 다시 받으세요.'
    }

    $trackedChanges = @(& git status --porcelain=v1 --untracked-files=no)
    if ($LASTEXITCODE -ne 0) {
        throw 'Git 작업 트리 상태를 확인하지 못했습니다.'
    }

    if ($trackedChanges.Count -gt 0) {
        Write-Host '커밋되지 않은 추적 파일 변경이 있어 pull을 중단합니다:' -ForegroundColor Yellow
        $trackedChanges | ForEach-Object { Write-Host "  $_" }
        throw '변경을 커밋하거나 별도로 보관한 뒤 다시 실행하세요. 자동 stash/reset은 수행하지 않습니다.'
    }

    $untrackedFiles = @(& git ls-files --others --exclude-standard)
    if ($LASTEXITCODE -ne 0) {
        throw 'Git 미추적 파일 상태를 확인하지 못했습니다.'
    }

    if ($untrackedFiles.Count -gt 0) {
        Write-Host "미추적 파일 $($untrackedFiles.Count)개는 보존한 채 pull을 계속합니다. 충돌하면 Git이 안전하게 중단합니다." -ForegroundColor Yellow
    }

    $branch = ((& git branch --show-current) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($branch)) {
        throw '분리된 HEAD 상태에서는 자동 pull하지 않습니다. 먼저 작업할 브랜치로 전환하세요.'
    }

    $upstream = ((& git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}' 2>$null) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($upstream)) {
        throw "현재 브랜치 '$branch'에 upstream이 없습니다. git push -u origin $branch 로 연결한 뒤 다시 실행하세요."
    }

    Write-Host "[$branch <- $upstream] fast-forward pull을 시작합니다." -ForegroundColor Cyan
    $pullArguments = @('pull', '--ff-only')
    if ($Prune) {
        $pullArguments += '--prune'
    }
    Invoke-GitChecked -Arguments $pullArguments

    $headCommit = ((& git rev-parse --short HEAD) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        throw '최종 커밋을 확인하지 못했습니다.'
    }

    Write-Host "pull 완료: $branch @ $headCommit" -ForegroundColor Green
}
catch {
    Write-Host "pull 실패: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

exit 0
