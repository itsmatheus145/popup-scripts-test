<#
.SYNOPSIS
    Baixa a imagem de wallpaper mais recente do repositório central e aplica
    na área de trabalho do usuário atual, imediatamente (sem precisar relogar).

.DESCRIPTION
    - Baixa a imagem via API do GitHub (mesmo repo dos scripts de popup)
    - Salva em cache local
    - Aplica via SystemParametersInfo (P/Invoke) - reflete na hora, ao contrário
      de só alterar o registro, que só valeria no próximo login

.NOTAS
    Rodando via launcher central puxado do Git.
#>

# ===================== CONFIGURAÇÃO =====================
$RepoOwner   = "itsmatheus145"
$RepoName    = "popup-scripts-test"
$Branch      = "main"
$GitHubToken = "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"
                     # ATENÇÃO: token de teste. Revogar após validação e mover
                     # para Windows Credential Manager antes de produção.

$WallpaperRepoPath = "images/wallpaper.jpg"   # caminho da imagem dentro do repo
$CacheDir           = "C:\ProgramData\FirstDecision\Wallpaper"
$LogDir              = "C:\ProgramData\FirstDecision\PopupScripts\Logs"
$WallpaperLocalPath  = Join-Path $CacheDir "wallpaper.jpg"
$ApiBase             = "https://api.github.com/repos/$RepoOwner/$RepoName/contents"
# ==========================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
    $logFile = Join-Path $LogDir ("execucao_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $line = "[{0}] [{1}] [Wallpaper] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

# ====================== P/Invoke: SystemParametersInfo ======================
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

$SPI_SETDESKWALLPAPER = 0x0014
$SPIF_UPDATEINIFILE   = 0x01
$SPIF_SENDCHANGE      = 0x02

# --------------------------------------------------------
Write-Log "===== Início da atualização de wallpaper ====="

if (-not (Test-Path $CacheDir)) {
    New-Item -Path $CacheDir -ItemType Directory -Force | Out-Null
}

$headers = @{
    "User-Agent" = "FirstDecision-WallpaperUpdater"
    "Accept"     = "application/vnd.github.v3+json"
}
if ($GitHubToken) {
    $headers["Authorization"] = "token $GitHubToken"
}

try {
    $url  = "$ApiBase/${WallpaperRepoPath}?ref=$Branch"
    $resp = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 20
    $bytes = [System.Convert]::FromBase64String($resp.content)
} catch {
    Write-Log "Falha ao baixar o wallpaper: $($_.Exception.Message)" "ERROR"
    exit 1
}

try {
    [System.IO.File]::WriteAllBytes($WallpaperLocalPath, $bytes)
    Write-Log "Wallpaper baixado e salvo em $WallpaperLocalPath ($($bytes.Length) bytes)."
} catch {
    Write-Log "Falha ao salvar o wallpaper localmente: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ====================== Aplica o wallpaper ======================
try {
    # Estilo "Preencher" (Fill) - ajusta antes de aplicar
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"

    $result = [Wallpaper]::SystemParametersInfo(
        $SPI_SETDESKWALLPAPER,
        0,
        $WallpaperLocalPath,
        ($SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    )

    if ($result -ne 0) {
        Write-Log "Wallpaper aplicado com sucesso."
    } else {
        Write-Log "SystemParametersInfo retornou falha ao aplicar o wallpaper." "ERROR"
    }
} catch {
    Write-Log "Erro ao aplicar o wallpaper: $($_.Exception.Message)" "ERROR"
}

Write-Log "===== Fim da atualização de wallpaper ====="
