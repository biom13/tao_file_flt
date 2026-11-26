# Day_len_GitHub_khong_gh.ps1
# Phien ban KHONG can GitHub CLI (gh). Chi can git.
# - Push: git tu mo cua so dang nhap GitHub (Git Credential Manager) o lan dau.
# - Bat Pages tu dong: can dan 1 token (PAT). Bo qua token thi script chi huong dan bat tay (3 thao tac).

$Owner    = 'biom13'
$Repo     = 'tao_file_flt'
$Branch   = 'main'
$RepoUrl  = "https://github.com/$Owner/$Repo.git"
$PagesUrl = "https://$Owner.github.io/$Repo/"
$Api      = "https://api.github.com/repos/$Owner/$Repo"

$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-Location -LiteralPath $PSScriptRoot

function Step([string]$Message) {
    Write-Host ''
    Write-Host "== $Message" -ForegroundColor Cyan
}

function Run([string]$Exe, [string[]]$Params) {
    & $Exe @Params
    if ($LASTEXITCODE -ne 0) {
        throw "Lenh that bai: $Exe $($Params -join ' ') (ma loi $LASTEXITCODE)"
    }
}

function Has([string]$Name) {
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Refresh-Path {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machine;$user"
}

function Get-StatusCode($ErrorRecord) {
    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch { return 0 }
}

try {
    Write-Host '============================================================'
    Write-Host "  DAY CODE LEN GITHUB (khong can gh): $Owner/$Repo"
    Write-Host '============================================================'

    # ---------- 1. Kiem tra git ----------
    Step '1. Kiem tra git'
    if (-not (Has 'git')) {
        Write-Host 'Chua co git, dang cai bang winget...'
        & winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        Refresh-Path
    }
    if (-not (Has 'git')) { throw 'Van chua thay git. Hay dong cua so nay, mo lai va chay lai script.' }
    Write-Host (& git --version)

    # ---------- 2. Kiem tra repo public ----------
    Step "2. Kiem tra repo $Owner/$Repo"
    $isPublic = $false
    try {
        $info = Invoke-RestMethod -Uri $Api -Headers @{ 'User-Agent' = 'flt-addin-setup' } -TimeoutSec 20
        $isPublic = (-not $info.private)
        Write-Host "Repo ton tai, che do: $(if ($isPublic) { 'public' } else { 'private' })"
    } catch {
        if ((Get-StatusCode $_) -eq 404) {
            Write-Host 'Khong thay repo public nay (co the dang o che do private hoac sai ten).' -ForegroundColor Yellow
        } else {
            Write-Host "Khong kiem tra duoc che do repo: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    if (-not $isPublic) {
        Write-Host 'GitHub Pages mien phi chi chay voi repo PUBLIC. Neu repo dang private, doi sang public:' -ForegroundColor Yellow
        Write-Host "  https://github.com/$Owner/$Repo/settings  (muc Danger Zone > Change visibility)"
        $go = Read-Host 'Van tiep tuc day code len? (Y/N)'
        if ($go -notmatch '^[Yy]') { throw 'Da dung theo yeu cau.' }
    }

    # ---------- 3. Khoi tao git local va commit ----------
    Step '3. Khoi tao git va commit'
    if (-not (Test-Path -LiteralPath '.git')) {
        Run 'git' @('init')
    }
    Run 'git' @('symbolic-ref', 'HEAD', "refs/heads/$Branch")

    $gitName  = (& git config user.name)
    $gitEmail = (& git config user.email)
    if (-not $gitName)  { Run 'git' @('config', 'user.name', $Owner) }
    if (-not $gitEmail) { Run 'git' @('config', 'user.email', "$Owner@users.noreply.github.com") }

    Run 'git' @('add', '-A')
    $pending = (& git status --porcelain)
    if ($pending) {
        Run 'git' @('commit', '-m', 'Add Excel add-in tao file FLT')
    } else {
        Write-Host 'Khong co thay doi moi de commit.'
    }

    # ---------- 4. Gan remote va push ----------
    Step '4. Day code len GitHub (lan dau se hien cua so dang nhap GitHub)'
    $remotes = @(& git remote)
    if ($remotes -contains 'origin') {
        Run 'git' @('remote', 'set-url', 'origin', $RepoUrl)
    } else {
        Run 'git' @('remote', 'add', 'origin', $RepoUrl)
    }

    & git push -u origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Push chua duoc. Thu gop voi noi dung san co tren GitHub roi push lai...' -ForegroundColor Yellow
        Run 'git' @('pull', 'origin', $Branch, '--allow-unrelated-histories', '--no-edit', '-X', 'ours')
        Run 'git' @('push', '-u', 'origin', $Branch)
    }

    # ---------- 5. Bat GitHub Pages ----------
    Step '5. Bat GitHub Pages'
    Write-Host 'De bat Pages tu dong can token (PAT) cua ban:'
    Write-Host '  Tao tai https://github.com/settings/personal-access-tokens/new'
    Write-Host "  Repository access: chon rieng repo $Repo; Permissions > Pages: Read and write"
    Write-Host '  (hoac token classic co quyen repo). Token KHONG duoc luu lai, nen xoa token sau khi xong.'
    $secure = Read-Host 'Dan token vao day (nhan Enter de bo qua va bat tay) ' -AsSecureString
    $token = ''
    if ($secure.Length -gt 0) {
        $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    $pagesDone = $false
    if ($token) {
        $headers = @{
            'Authorization'        = "Bearer $token"
            'Accept'               = 'application/vnd.github+json'
            'X-GitHub-Api-Version' = '2022-11-28'
            'User-Agent'           = 'flt-addin-setup'
        }
        $body = (@{ source = @{ branch = $Branch; path = '/' } } | ConvertTo-Json -Compress)
        try {
            Invoke-RestMethod -Method Post -Uri "$Api/pages" -Headers $headers -Body $body -ContentType 'application/json' | Out-Null
            $pagesDone = $true
            Write-Host 'Da bat GitHub Pages.'
        } catch {
            $code = Get-StatusCode $_
            if ($code -eq 409 -or $code -eq 422) {
                # Pages da duoc bat truoc do -> cap nhat nguon
                try {
                    Invoke-RestMethod -Method Put -Uri "$Api/pages" -Headers $headers -Body $body -ContentType 'application/json' | Out-Null
                    $pagesDone = $true
                    Write-Host 'GitHub Pages da bat san, da cap nhat nguon (main, thu muc goc).'
                } catch {
                    Write-Host "Khong cap nhat duoc Pages (ma $(Get-StatusCode $_)): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Khong bat duoc Pages bang token (ma $code): $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host 'Kiem tra token co quyen Pages (Read and write) cho dung repo.' -ForegroundColor Yellow
            }
        }
        $token = ''
        $headers = $null
    }

    if (-not $pagesDone) {
        Write-Host ''
        Write-Host 'Bat Pages bang tay (khoang 30 giay):' -ForegroundColor Yellow
        Write-Host "  1. Mo https://github.com/$Owner/$Repo/settings/pages"
        Write-Host '  2. Muc Source chon: Deploy from a branch'
        Write-Host '  3. Branch: main, thu muc: / (root), bam Save'
        Read-Host 'Lam xong thi nhan Enter de script kiem tra link'
    }

    # ---------- 6. Cho trang online ----------
    Step '6. Cho GitHub Pages xuat ban (toi da khoang 5 phut)'
    $httpCode = 0
    for ($i = 1; $i -le 30; $i++) {
        try {
            $resp = Invoke-WebRequest -Uri ($PagesUrl + 'taskpane.html') -UseBasicParsing -TimeoutSec 20
            $httpCode = [int]$resp.StatusCode
        } catch {
            $httpCode = Get-StatusCode $_
        }
        Write-Host "  [$i/30] ma phan hoi: $httpCode"
        if ($httpCode -eq 200) { break }
        Start-Sleep -Seconds 10
    }

    Write-Host ''
    Write-Host '============================================================'
    if ($httpCode -eq 200) {
        Write-Host '  HOAN TAT! Add-in da online:' -ForegroundColor Green
    } else {
        Write-Host '  Code da len GitHub nhung trang chua mo duoc.' -ForegroundColor Yellow
        Write-Host '  Kiem tra da bat Pages chua, doi them 1-2 phut roi mo link ben duoi.'
    }
    Write-Host "  Trang add-in : $($PagesUrl)taskpane.html"
    Write-Host "  Repo         : https://github.com/$Owner/$Repo"
    Write-Host "  File manifest: $(Join-Path $PSScriptRoot 'manifest.xml')"
    Write-Host '  Buoc tiep    : cai manifest.xml vao Excel 2021 (xem README.md, Buoc 3).'
    Write-Host '============================================================'
}
catch {
    Write-Host ''
    Write-Host "LOI: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Read-Host 'Nhan Enter de dong'
