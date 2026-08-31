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
    $CspKeyPath    = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

    try {
        if (-not (Test-Path $PolicyKeyPath)) {
            New-Item -Path $PolicyKeyPath -Force | Out-Null
        }
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenImage" -Value $ImagePath -Type String
        Set-ItemProperty -Path $PolicyKeyPath -Name "NoChangingLockScreen" -Value 1 -Type DWord
        Set-ItemProperty -Path $PolicyKeyPath -Name "LockScreenOverlaysDisabled" -Value 1 -Type DWord
        Write-Log "Chave de política Personalization configurada com sucesso (LockScreenImage = $ImagePath)." "INFO" $origemLog
    } catch {
        Write-Log "Falha ao configurar a chave de política Personalization: $($_.Exception.Message)" "ERROR" $origemLog
    }

    # A chave CSP (PersonalizationCSP) é de MÁQUINA (HKLM), não por usuário -
    # não precisa carregar hive de ninguém, ao contrário do que tentamos antes.
    try {
        if (-not (Test-Path $CspKeyPath)) {
            New-Item -Path $CspKeyPath -Force | Out-Null
        }
        Set-ItemProperty -Path $CspKeyPath -Name "LockScreenImagePath" -Value $ImagePath -Type String
        Set-ItemProperty -Path $CspKeyPath -Name "LockScreenImageUrl" -Value $ImagePath -Type String
        Set-ItemProperty -Path $CspKeyPath -Name "LockScreenImageStatus" -Value 1 -Type DWord
        Write-Log "Chave PersonalizationCSP configurada com sucesso." "INFO" $origemLog
    } catch {
        Write-Log "Falha ao configurar a chave PersonalizationCSP: $($_.Exception.Message)" "ERROR" $origemLog
    }

    try {
        $SpotlightKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        if (-not (Test-Path $SpotlightKeyPath)) {
            New-Item -Path $SpotlightKeyPath -Force | Out-Null
        }
        Set-ItemProperty -Path $SpotlightKeyPath -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord
        Set-ItemProperty -Path $SpotlightKeyPath -Name "DisableWindowsSpotlightOnLockScreen" -Value 1 -Type DWord

        $SystemPolicyKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
        if (-not (Test-Path $SystemPolicyKeyPath)) {
            New-Item -Path $SystemPolicyKeyPath -Force | Out-Null
        }
        Set-ItemProperty -Path $SystemPolicyKeyPath -Name "DisableLockScreenAppNotifications" -Value 1 -Type DWord

        Write-Log "Windows Spotlight e notificações na tela de bloqueio desativados." "INFO" $origemLog
    } catch {
        Write-Log "Falha ao desativar Spotlight/notificações: $($_.Exception.Message)" "WARN" $origemLog
    }

    try {
        $DefaultLockScreenPath = "C:\Windows\Web\Screen\img100.jpg"
        if (Test-Path $DefaultLockScreenPath) {
            # Esse arquivo é protegido pelo Windows Resource Protection - nem
            # o SYSTEM consegue sobrescrever sem antes tomar posse dele
            # explicitamente (mesma técnica já usada para limpar SystemData).
            & takeown.exe /F $DefaultLockScreenPath *> $null
            & icacls.exe $DefaultLockScreenPath /grant "SYSTEM:F" *> $null
        }
        if (Test-Path (Split-Path $DefaultLockScreenPath)) {
            Copy-Item -Path $ImagePath -Destination $DefaultLockScreenPath -Force
            Write-Log "Arquivo padrão de tela de bloqueio sobrescrito (reforço)." "INFO" $origemLog
        }
    } catch {
        Write-Log "Não foi possível sobrescrever o arquivo padrão (reforço opcional): $($_.Exception.Message)" "WARN" $origemLog
    }

    # ====================== Limpeza agressiva do cache da tela de bloqueio ======================
    # O Windows guarda dados renderizados/cacheados em SystemData, protegidos
    # por TrustedInstaller. Uma limpeza "suave" (sem tomar posse) não remove
    # esse cache - é preciso takeown + icacls antes de apagar. O Windows
    # recria essa pasta sozinho na próxima vez que precisar dela.
    try {
        $systemDataPath = "C:\ProgramData\Microsoft\Windows\SystemData"
        if (Test-Path $systemDataPath) {
            & takeown.exe /F $systemDataPath /R /D Y *> $null
            & icacls.exe $systemDataPath /grant "Administrators:F" /T *> $null
            Remove-Item -Path $systemDataPath -Recurse -Force -ErrorAction Stop
            Write-Log "Cache de tela de bloqueio (SystemData) limpo com sucesso." "INFO" $origemLog
        }
    } catch {
        Write-Log "Falha ao limpar o cache SystemData: $($_.Exception.Message)" "WARN" $origemLog
    }

    # ====================== Limpeza de tarefas obsoletas (privilégio SYSTEM) ======================
    # Algumas tarefas antigas (criadas por GPOs já aposentadas, como "TASK
    # POPUP - FD") ficam com uma ACL que só Administrador/SYSTEM consegue
    # apagar - um processo de usuário comum (como a sincronização normal do
    # launcher) recebe "Acesso negado" pra sempre. Como este script já roda
    # como SYSTEM nesta tarefa (Tela de Bloqueio, a cada hora), aproveitamos
    # esse privilégio pra garantir a remoção em qualquer máquina, sem
    # precisar de intervenção manual como Administrador.
    $tarefasObsoletasComPrivilegio = @(
        "TASK POPUP - FD",   # sistema antigo de popup de bem-estar (GPO "POP_UP", substituído por Notificacao-BemEstar.ps1)
        "teste",             # tarefa de teste manual antiga, criada como Administrador
        "teste popup"        # tarefa de teste manual antiga, criada como Administrador
    )
    foreach ($nomeAntigo in $tarefasObsoletasComPrivilegio) {
        if (Get-ScheduledTask -TaskName $nomeAntigo -ErrorAction SilentlyContinue) {
            try {
                Unregister-ScheduledTask -TaskName $nomeAntigo -Confirm:$false -ErrorAction Stop
                Write-Log "Tarefa obsoleta '$nomeAntigo' removida (privilégio SYSTEM)." "INFO" $origemLog
            } catch {
                Write-Log "Falha ao remover a tarefa obsoleta '$nomeAntigo' mesmo como SYSTEM: $($_.Exception.Message)" "ERROR" $origemLog
            }
        }
    }

    # ====================== Reescreve a tarefa de logon (GPO/GPP) para caminho local ======================
    # Tarefas criadas via GPO (GPP Scheduled Tasks) são registradas por um
    # processo do próprio Windows rodando como SYSTEM - mesmo "pertencendo"
    # ao usuário que a executa, um processo comum não consegue modificá-la
    # (Acesso negado, sempre, não importa quantas vezes tente). Só SYSTEM
    # consegue reescrever. Sem isso, a tarefa continua apontando pro
    # caminho de rede (SYSVOL) pra sempre, mostrando erro em todo logon sem
    # conexão com o domínio (ex: home office sem VPN).
    $taskNameLogonGpo = "FirstDecision - Launcher Popup Central"
    $launcherLocalDirFixo  = "C:\ProgramData\FirstDecision\Launcher"
    $launcherLocalPathFixo = Join-Path $launcherLocalDirFixo "Update-PopupScripts.ps1"
    $vbsLocalPathFixo      = Join-Path $launcherLocalDirFixo "RunHidden.vbs"

    try {
        $tarefaLogon = Get-ScheduledTask -TaskName $taskNameLogonGpo -ErrorAction SilentlyContinue
        if ($tarefaLogon) {
            $argumentosAtuais = $tarefaLogon.Actions[0].Arguments
            $jaEhLocal = $argumentosAtuais -like "*$launcherLocalDirFixo*"
            $arquivosLocaisExistem = (Test-Path $launcherLocalPathFixo) -and (Test-Path $vbsLocalPathFixo)

            if (-not $jaEhLocal -and $arquivosLocaisExistem) {
                $actionLogonLocal = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsLocalPathFixo`" `"$launcherLocalPathFixo`""
                Set-ScheduledTask -TaskName $taskNameLogonGpo -Action $actionLogonLocal -ErrorAction Stop | Out-Null
                Write-Log "Tarefa de logon '$taskNameLogonGpo' reescrita para caminho local (privilégio SYSTEM)." "INFO" $origemLog
            }
        }
    } catch {
        Write-Log "Falha ao reescrever a tarefa de logon '$taskNameLogonGpo' mesmo como SYSTEM: $($_.Exception.Message)" "ERROR" $origemLog
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
        # Limpa o cache "TranscodedWallpaper" do Windows - sem isso, o Windows
        # às vezes reutiliza a versão processada antiga, mesmo com o arquivo
        # de origem já trocado no mesmo caminho (bug conhecido do Explorer).
        $transcodedCachePath = Join-Path $env:APPDATA "Microsoft\Windows\Themes\TranscodedWallpaper"
        if (Test-Path $transcodedCachePath) {
            Remove-Item -Path $transcodedCachePath -Force -ErrorAction SilentlyContinue
            Write-Log "Cache TranscodedWallpaper removido (forçando regeneração)." "INFO" $origemLog
        }
        $cachedFilesPath = Join-Path $env:APPDATA "Microsoft\Windows\Themes\CachedFiles"
        if (Test-Path $cachedFilesPath) {
            Remove-Item -Path (Join-Path $cachedFilesPath "*") -Force -ErrorAction SilentlyContinue
        }

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
