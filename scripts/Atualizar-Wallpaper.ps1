<#
.SYNOPSIS
    Aplica a imagem mais recente como wallpaper (contexto de usuário) e/ou
    como tela de bloqueio (contexto SYSTEM), dependendo de como este script
    for invocado.

.DESCRIPTION
    Este único arquivo cobre os dois casos, detectando sozinho o contexto:

      - Rodando como usuário comum (chamado pela tarefa "At log on"):
        aplica o WALLPAPER (HKCU - não precisa de privilégio elevado).

      - Rodando como SYSTEM (chamado pela tarefa de Computer Configuration):
        aplica a TELA DE BLOQUEIO (HKLM - precisa de privilégio de máquina,
        por isso só funciona quando disparado nesse contexto elevado).

    Não baixa nada da internet e não carrega nenhuma credencial - reaproveita
    a imagem que o launcher central já baixou em:
        C:\ProgramData\FirstDecision\PopupScripts\Assets\wallpaper.jpg
    (pasta compartilhada, legível por qualquer usuário/SYSTEM na máquina).

.NOTAS
    Ainda são necessárias DUAS Tarefas Agendadas na GPO (uma em User
    Configuration, outra em Computer Configuration rodando como SYSTEM) -
    isso é uma limitação do Windows (wallpaper é HKCU, tela de bloqueio é
    HKLM), não uma escolha de arquitetura. Um processo de usuário comum não
    tem permissão para escrever em HKLM, não importa como o código esteja
    organizado. Mas ambas as tarefas chamam este MESMO arquivo.
#>

# ===================== CONFIGURAÇÃO =====================
$ImagePath = "C:\ProgramData\FirstDecision\PopupScripts\Assets\wallpaper.jpg"
$LogDir    = "C:\ProgramData\FirstDecision\PopupScripts\Logs"
# ==========================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [string]$Origem = "Wallpaper")
    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
    $logFile = Join-Path $LogDir ("execucao_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $line = "[{0}] [{1}] [{2}] {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Origem, $Message
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

# Detecta o contexto de execução: SYSTEM ou usuário comum
$identidadeAtual = [Security.Principal.WindowsIdentity]::GetCurrent()
$rodandoComoSystem = $identidadeAtual.IsSystem

$origemLog = if ($rodandoComoSystem) { "TelaBloqueio" } else { "Wallpaper" }
Write-Log "===== Início da execução (contexto: $(if ($rodandoComoSystem) {'SYSTEM'} else {'usuário'})) =====" "INFO" $origemLog

if (-not (Test-Path $ImagePath)) {
    Write-Log "Imagem não encontrada em $ImagePath (o launcher ainda não baixou o asset nesta máquina)." "WARN" $origemLog
    exit 0
}

if ($rodandoComoSystem) {
    # ====================== CONTEXTO SYSTEM: TELA DE BLOQUEIO (HKLM) ======================
    $PolicyKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"

    try {
        if (-not (Test-Path $PolicyKeyPath)) {
            New-Item -Path $PolicyKeyPath -Force | Out-Null
        }
        # Diferentes builds/versões do Windows leem nomes de valor diferentes
        # para a mesma política - grava todas as variantes conhecidas para
        # maximizar compatibilidade.
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenImage" -Value $ImagePath -Type String
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenImagePath" -Value $ImagePath -Type String
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenImageUrl" -Value $ImagePath -Type String
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenImageStatus" -Value 1 -Type DWord
        Set-ItemProperty -Path $PolicyKeyPath -Name "NoChangingLockScreen" -Value 1 -Type DWord
        Write-Log "Chave de política Personalization configurada com sucesso (todas as variantes de nome de valor, LockScreenImage* = $ImagePath)." "INFO" $origemLog
    } catch {
        Write-Log "Falha ao configurar a chave de política Personalization: $($_.Exception.Message)" "ERROR" $origemLog
    }

    try {
        $SpotlightKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        if (-not (Test-Path $SpotlightKeyPath)) {
            New-Item -Path $SpotlightKeyPath -Force | Out-Null
        }
        Set-ItemProperty -Path $SpotlightKeyPath -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord
        Write-Log "Windows Spotlight desativado na tela de bloqueio." "INFO" $origemLog
    } catch {
        Write-Log "Falha ao desativar o Windows Spotlight: $($_.Exception.Message)" "WARN" $origemLog
    }

    try {
        $DefaultLockScreenPath = "C:\Windows\Web\Screen\img100.jpg"
        if (Test-Path (Split-Path $DefaultLockScreenPath)) {
            Copy-Item -Path $ImagePath -Destination $DefaultLockScreenPath -Force
            Write-Log "Arquivo padrão de tela de bloqueio sobrescrito (reforço para SKUs onde a política sozinha não é suficiente)." "INFO" $origemLog
        }
    } catch {
        Write-Log "Não foi possível sobrescrever o arquivo padrão (reforço opcional): $($_.Exception.Message)" "WARN" $origemLog
    }

} else {
    # ====================== CONTEXTO USUÁRIO: WALLPAPER (HKCU) ======================
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

    $algumaChaveRemovida = $false
    foreach ($PolicyKeyPath in $PolicyKeyPaths) {
        if (-not (Test-Path $PolicyKeyPath)) { continue }
        foreach ($nomeValor in @("Wallpaper", "WallpaperStyle")) {
            try {
                $valorAtual = Get-ItemProperty -Path $PolicyKeyPath -Name $nomeValor -ErrorAction SilentlyContinue
                if ($valorAtual) {
                    Remove-ItemProperty -Path $PolicyKeyPath -Name $nomeValor -ErrorAction Stop
                    Write-Log "Valor de política '$nomeValor' removido de $PolicyKeyPath." "INFO" $origemLog
                    $algumaChaveRemovida = $true
                }
            } catch {
                Write-Log "Não foi possível remover '$nomeValor' de ${PolicyKeyPath}: $($_.Exception.Message)" "WARN" $origemLog
            }
        }
    }
    if (-not $algumaChaveRemovida) {
        Write-Log "Nenhuma chave de política de wallpaper encontrada (esperado após a GPO ser desativada)." "INFO" $origemLog
    }

    try {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"

        $result = [Wallpaper]::SystemParametersInfo(
            $SPI_SETDESKWALLPAPER, 0, $ImagePath, ($SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
        )

        if ($result -ne 0) {
            Write-Log "Wallpaper aplicado com sucesso ($ImagePath)." "INFO" $origemLog
        } else {
            Write-Log "SystemParametersInfo retornou falha ao aplicar o wallpaper." "ERROR" $origemLog
        }

        Start-Process -FilePath "rundll32.exe" -ArgumentList "user32.dll,UpdatePerUserSystemParameters" -NoNewWindow -Wait
        Write-Log "Parâmetros de usuário atualizados (refresh da área de trabalho)." "INFO" $origemLog
    } catch {
        Write-Log "Erro ao aplicar o wallpaper: $($_.Exception.Message)" "ERROR" $origemLog
    }
}

Write-Log "===== Fim da execução =====" "INFO" $origemLog
