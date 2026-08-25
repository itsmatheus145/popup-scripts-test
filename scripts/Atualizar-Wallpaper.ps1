<#
.SYNOPSIS
    Aplica o wallpaper mais recente na área de trabalho do usuário atual.

.DESCRIPTION
    Este script NÃO baixa nada da internet e NÃO carrega nenhuma credencial -
    ele apenas aplica a imagem que o launcher central (Update-PopupScripts.ps1)
    já baixou para C:\ProgramData\FirstDecision\PopupScripts\Assets\wallpaper.jpg.

    Por não ter nenhum segredo embutido, este arquivo pode ficar no repositório
    Git normalmente, sem risco de exposição de credenciais.

.NOTAS
    Aplica via SystemParametersInfo (P/Invoke) - reflete na hora, ao contrário
    de só alterar o registro, que só valeria no próximo login.
#>

# ===================== CONFIGURAÇÃO =====================
$ImagePath = "C:\ProgramData\FirstDecision\PopupScripts\Assets\wallpaper.jpg"
$LogDir    = "C:\ProgramData\FirstDecision\PopupScripts\Logs"
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

$PolicyKeyPaths = @(
    "HKCU:\Software\Policies\Microsoft\Windows\System",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
)

# --------------------------------------------------------
Write-Log "===== Início da aplicação de wallpaper ====="

if (-not (Test-Path $ImagePath)) {
    Write-Log "Imagem não encontrada em $ImagePath (o launcher deveria tê-la baixado antes de chamar este script)." "ERROR"
    exit 1
}

# Remove eventuais chaves de política que forcem um wallpaper fixo
# (ex: GPO "Desktop Wallpaper" ainda não desativada / não propagada).
# Sem isso, o valor definido via SystemParametersInfo é sobrescrito, pois
# valores em ...\Policies\... têm precedência sobre HKCU:\Control Panel\Desktop.
# Windows grava em caminhos diferentes conforme a versão - por isso os dois.
$algumaChaveRemovida = $false

foreach ($PolicyKeyPath in $PolicyKeyPaths) {
    if (-not (Test-Path $PolicyKeyPath)) { continue }

    foreach ($nomeValor in @("Wallpaper", "WallpaperStyle")) {
        try {
            $valorAtual = Get-ItemProperty -Path $PolicyKeyPath -Name $nomeValor -ErrorAction SilentlyContinue
            if ($valorAtual) {
                Remove-ItemProperty -Path $PolicyKeyPath -Name $nomeValor -ErrorAction Stop
                Write-Log "Valor de política '$nomeValor' removido de $PolicyKeyPath."
                $algumaChaveRemovida = $true
            }
        } catch {
            Write-Log "Não foi possível remover '$nomeValor' de ${PolicyKeyPath}: $($_.Exception.Message)" "WARN"
        }
    }
}

if (-not $algumaChaveRemovida) {
    Write-Log "Nenhuma chave de política de wallpaper encontrada (esperado após a GPO ser desativada)."
}

# ====================== Aplica o wallpaper ======================
try {
    # Estilo "Preencher" (Fill)
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"

    $result = [Wallpaper]::SystemParametersInfo(
        $SPI_SETDESKWALLPAPER,
        0,
        $ImagePath,
        ($SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    )

    if ($result -ne 0) {
        Write-Log "Wallpaper aplicado com sucesso ($ImagePath)."
    } else {
        Write-Log "SystemParametersInfo retornou falha ao aplicar o wallpaper." "ERROR"
    }

    # Reforço: força o Windows a reler os parâmetros do usuário. Necessário em
    # alguns casos para a mudança refletir imediatamente na sessão ativa.
    Start-Process -FilePath "rundll32.exe" -ArgumentList "user32.dll,UpdatePerUserSystemParameters" -NoNewWindow -Wait
    Write-Log "Parâmetros de usuário atualizados (refresh da área de trabalho)."
} catch {
    Write-Log "Erro ao aplicar o wallpaper: $($_.Exception.Message)" "ERROR"
}

Write-Log "===== Fim da aplicação de wallpaper ====="
