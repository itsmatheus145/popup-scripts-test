<#
.SYNOPSIS
    Launcher central - baixa manifest.json do repositório público do suporte
    no GitHub, mantém os scripts e a imagem de wallpaper atualizados
    localmente, e garante que cada script tenha sua PRÓPRIA Tarefa Agendada
    dedicada, no horário definido no manifest.

.DESCRIPTION
    Modelo de execução:
      - Chamado SEM -RunScript: modo "sincronização completa". Faz tudo:
        instala a cópia local (se necessário), baixa manifest/scripts/assets
        que mudaram (comparando o SHA-256 de cada arquivo com o manifest),
        aplica o wallpaper, reconcilia (cria/atualiza/remove) uma Tarefa
        Agendada dedicada para cada script do manifest e manda o check-in
        diário da máquina.
        É este modo que a GPO chama (via SYSVOL) e que a tarefa local
        "FirstDecision-LauncherSync" chama.

      - Chamado COM -RunScript <nome.ps1>: modo "execução direta e rápida".
        Não baixa nada, não mexe em tarefas - só executa o arquivo já
        cacheado localmente. É este modo que cada tarefa dedicada
        (FD-Task-<script>) usa para disparar seu próprio popup, sem
        depender de rede nem repetir toda a lógica de sincronização.

    Sem token: o repositório é público e não guarda nenhum segredo. Os
    arquivos são baixados de raw.githubusercontent.com (CDN, sem o limite
    de 60 chamadas/hora da API para quem não tem token - importante porque
    no escritório todas as máquinas saem pelo mesmo IP).

    Integridade: o manifest traz o SHA-256 de cada arquivo (inclusive deste
    launcher). Arquivo baixado com hash diferente é descartado e o cache
    anterior é mantido - nunca grava um arquivo pela metade ou adulterado.

    "Auto-cura": toda sincronização reconcilia de novo - tarefa apagada à
    mão é recriada, horário mudado no manifest é aplicado, script removido
    ou com run=false tem sua tarefa apagada, e arquivo local alterado volta
    a ser o do manifest.

.NOTES
    Autor: Matheus Itacaramby Souza - First Decision
#>

param(
    [string]$RunScript = "",   # se informado: executa só esse script cacheado e sai (sem sync)
    [switch]$Bootstrap          # marca explicitamente "isto é uma instalação inicial vinda do SYSVOL/GPO" -
                                # não depende de comparar caminhos (frágil: o BootstrapSilencioso.vbs já
                                # copia o arquivo pra local ANTES de rodar, então $PSCommandPath nunca
                                # reflete a origem real de forma confiável)
)

# ===================== CONFIGURAÇÃO =====================
$RepoOwner   = "firstdecisionltda"   # conta do suporte no GitHub (suporte@firstdecision.com.br)
$RepoName    = "popups-fd"
$Branch      = "main"
# Repositório PÚBLICO e sem segredos: as máquinas leem direto, sem token. O
# raw.githubusercontent.com não tem o limite de 60 chamadas/hora por IP da API
# (importante: no escritório todas as máquinas saem pelo mesmo IP). Em troca,
# guarda cópias por ~5 min, então uma mudança pode levar esse tempo a valer;
# o hash do manifest garante que nunca se grava arquivo "meio velho".
$RawBase     = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch"

$CacheDir       = "C:\ProgramData\FirstDecision\PopupScripts"
$AssetsDir      = Join-Path $CacheDir "Assets"
$LogDir         = Join-Path $CacheDir "Logs"
$VersionFile    = Join-Path $CacheDir "current_version.txt"
# Cópia local do manifest - os scripts leem dela as URLs de config (ex: fluxo
# de aniversários), assim dá pra trocar uma URL sem mexer em script nenhum.
$ManifestLocal  = Join-Path $CacheDir "manifest.json"
$CheckinFile    = Join-Path $CacheDir "ultimo_checkin.txt"

# Caminho local fixo do launcher. Depois da primeira execução via GPO/SYSVOL,
# todas as tarefas locais apontam pra cá - nunca mais precisam do domínio.
$LauncherLocalDir  = "C:\ProgramData\FirstDecision\Launcher"
$LauncherLocalPath = Join-Path $LauncherLocalDir "Update-PopupScripts.ps1"

# Nome do script de wallpaper - tratado à parte dos demais (não tem tarefa
# dedicada própria; é aplicado a cada sincronização, porque deve refletir
# qualquer imagem nova o quanto antes).
$WallpaperScriptName = "Atualizar-Wallpaper.ps1"

# Intervalo (em minutos) da tarefa local recorrente de sincronização completa.
$IntervaloMinutosSincronizacao = 20

# Atraso (em segundos) só quando esta execução partiu de um caminho remoto
# (SYSVOL/GPO) - dá tempo da rede subir logo após conectar/logar.
$DelayPosConexaoSegundos = 60

# Nome do arquivo do próprio launcher dentro do repositório (raiz, separado
# da pasta scripts/ que é só para os scripts "de carga" tipo popups).
$LauncherRepoPath = "Update-PopupScripts.ps1"

$TaskPrefixDedicada = "FD-Task-"
# ==========================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        # Mesma proteção já aplicada ao diretório do Launcher: garante que
        # usuários comuns sempre consigam escrever aqui, mesmo que esta
        # criação específica tenha acontecido numa sessão elevada.
        try { & icacls.exe $LogDir /grant "*S-1-5-32-545:(OI)(CI)M" /T *> $null } catch { }
    }
    $logFile = Join-Path $LogDir ("execucao_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    try {
        Add-Content -Path $logFile -Value $line -ErrorAction Stop
    } catch {
        # Auto-correção: se a escrita falhar (ex: arquivo criado antes numa
        # sessão elevada, com ACL restrita), tenta liberar a permissão e
        # escrever de novo, uma vez, antes de desistir silenciosamente.
        try {
            & icacls.exe $logFile /grant "*S-1-5-32-545:M" *> $null
            Add-Content -Path $logFile -Value $line -ErrorAction Stop
        } catch { }
    }
    Write-Host $line
}

function Get-Sha256 {
    param([byte[]]$Bytes)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

function Get-LocalFileSha256 {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return "" }
    try {
        return Get-Sha256 -Bytes ([System.IO.File]::ReadAllBytes($Path))
    } catch {
        return ""
    }
}

function Get-RemoteFileBytes {
    param([string]$RepoPath, [string]$CacheBuster = "")
    # WebClient.DownloadData devolve os bytes exatos (Invoke-WebRequest no
    # PowerShell 5.1 converte texto e pode mexer em BOM/encoding, o que
    # quebraria a comparação de hash).
    $segmentosCodificados = $RepoPath -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }
    $url = "$RawBase/$($segmentosCodificados -join '/')"
    # Um parâmetro que muda junto com o conteúdo (o hash esperado, ou o horário
    # no caso do manifest) ajuda a não pegar cópia velha do CDN.
    if ($CacheBuster) { $url = "${url}?v=$CacheBuster" }
    $wc = New-Object System.Net.WebClient
    try {
        $wc.Headers.Add("User-Agent", "FirstDecision-PopupLauncher")
        return $wc.DownloadData($url)
    } catch {
        Write-Log "Falha ao baixar $RepoPath : $($_.Exception.Message)" "ERROR"
        return $null
    } finally {
        $wc.Dispose()
    }
}

function Save-VerifiedFile {
    # Baixa um arquivo do repo e só grava se o SHA-256 bater com o manifest.
    # Retorna $true se o arquivo local ficou igual ao do manifest.
    param([string]$RepoPath, [string]$LocalPath, [string]$ExpectedSha256)

    $expected = "$ExpectedSha256".ToUpper()
    if ($expected -and ((Get-LocalFileSha256 -Path $LocalPath) -eq $expected)) {
        return $true   # já está igual - nada a baixar
    }

    $bytes = Get-RemoteFileBytes -RepoPath $RepoPath -CacheBuster $expected
    if (-not $bytes) {
        Write-Log "Mantendo a versão em cache de $RepoPath (se existir)." "WARN"
        return $false
    }

    if ($expected) {
        $hashRecebido = Get-Sha256 -Bytes $bytes
        if ($hashRecebido -ne $expected) {
            Write-Log "Hash de $RepoPath não confere (esperado $expected, recebido $hashRecebido). Arquivo descartado; mantendo cache. Se acabou de publicar, o CDN pode estar atrasado - tenta de novo na próxima sincronização." "ERROR"
            return $false
        }
    }

    try {
        [System.IO.File]::WriteAllBytes($LocalPath, $bytes)
        Write-Log "$RepoPath atualizado -> $LocalPath"
        return $true
    } catch {
        Write-Log "Falha ao gravar $LocalPath : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Update-LauncherFromGitHub {
    param($Manifest)

    # Compara o hash do launcher local com o informado no manifest. Garante
    # que máquinas que nunca mais voltam a tocar no domínio/SYSVOL recebem
    # correções no próprio launcher, não só nos scripts que ele gerencia.
    if (-not (Test-Path $LauncherLocalPath)) {
        return  # Install-LocalLauncher ainda não rodou nesta execução - nada a comparar
    }

    $hashEsperado = "$($Manifest.launcher.sha256)".ToUpper()
    if (-not $hashEsperado) {
        Write-Log "Manifest sem launcher.sha256 - atualização do próprio launcher pulada." "WARN"
        return
    }

    if ((Get-LocalFileSha256 -Path $LauncherLocalPath) -eq $hashEsperado) {
        Write-Log "Launcher já está na versão mais recente (hash confere com o manifest)."
        return
    }

    if (Save-VerifiedFile -RepoPath $LauncherRepoPath -LocalPath $LauncherLocalPath -ExpectedSha256 $hashEsperado) {
        Write-Log "Launcher atualizado (o código do próprio Update-PopupScripts.ps1 mudou). A versão nova entra em vigor na próxima execução agendada."
    }
}

function Invoke-CachedScript {
    param([string]$ScriptPath)
    try {
        Write-Log "Executando $ScriptPath..."
        Start-Process -FilePath "powershell.exe" `
            -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $ScriptPath) `
            -WindowStyle Hidden -Wait
        Write-Log "$ScriptPath executado."
    } catch {
        Write-Log "Erro ao executar $ScriptPath : $($_.Exception.Message)" "ERROR"
    }
}

function Install-RunHiddenWrapper {
    # O flag "Hidden" da Tarefa Agendada e o "-WindowStyle Hidden" do
    # PowerShell às vezes não evitam um flash rápido do console (comum em
    # builds mais novas do Windows com Windows Terminal). Um wrapper VBScript
    # que chama WScript.Shell.Run com estilo 0 suprime a janela ANTES dela
    # ser criada, não depois - é a forma que realmente elimina o flash.
    # Auto-suficiente: cria a pasta local se ainda não existir.
    if (-not (Test-Path $LauncherLocalDir)) {
        New-Item -Path $LauncherLocalDir -ItemType Directory -Force | Out-Null
    }
    $vbsPath = Join-Path $LauncherLocalDir "RunHidden.vbs"

    # So (re)cria se ainda nao existir. Evita tentar sobrescrever um
    # arquivo que o bootstrap via SYSTEM (BootstrapTelaBloqueio.ps1) ja
    # criou antes - um processo de usuario comum nao tem permissao pra
    # sobrescrever um arquivo cujo dono e SYSTEM, e como o conteudo nunca
    # muda, nao ha necessidade real de recriar toda vez.
    if (Test-Path $vbsPath) {
        return $vbsPath
    }

    $vbsContent = @'
Dim shell, fso, cmd, i
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(WScript.Arguments(0)) Then
    WScript.Quit
End If

cmd = "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File """ & WScript.Arguments(0) & """"
For i = 1 To WScript.Arguments.Count - 1
    cmd = cmd & " " & WScript.Arguments(i)
Next
shell.Run cmd, 0, True
'@
    try {
        Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force
    } catch {
        Write-Log "Falha ao instalar o wrapper RunHidden.vbs: $($_.Exception.Message)" "ERROR"
    }
    return $vbsPath
}

function Install-LocalLauncher {
    param([string]$VbsPath, [bool]$ForcarBootstrap = $false)

    # Roda a instalação completa se: (a) o -Bootstrap foi passado
    # explicitamente pelo BootstrapSilencioso.vbs (sinal confiável, não
    # depende de comparar caminhos), OU (b) $PSCommandPath é mesmo diferente
    # do caminho local (cobre testes manuais chamando o UNC diretamente).
    # Isso é necessário porque o BootstrapSilencioso.vbs JÁ copia o arquivo
    # pra local antes de rodar o PowerShell - então $PSCommandPath nunca
    # reflete "veio do SYSVOL" de forma confiável nessa cadeia.
    if (-not $ForcarBootstrap -and $PSCommandPath -and ($PSCommandPath -eq $LauncherLocalPath)) {
        return $false
    }

    Write-Log "Instalando/atualizando cópia local (bootstrap $(if ($ForcarBootstrap) {'via sinal explícito'} else {'via caminho remoto: ' + $PSCommandPath}))..."

    try {
        if (-not (Test-Path $LauncherLocalDir)) {
            New-Item -Path $LauncherLocalDir -ItemType Directory -Force | Out-Null
        }
        if ($PSCommandPath -ne $LauncherLocalPath) {
            Copy-Item -Path $PSCommandPath -Destination $LauncherLocalPath -Force
        }

        # Reforço: garante que usuários comuns sempre consigam escrever nessa
        # pasta, mesmo que esta execução específica tenha rodado elevada (o
        # que normalmente restringiria a ACL só a Administradores). Sem isso,
        # a próxima sincronização não-elevada falha com "Acesso negado".
        # Usa o SID do grupo Usuários (S-1-5-32-545) em vez do nome, que muda
        # conforme o idioma do Windows ("Users" em inglês).
        try {
            & icacls.exe $LauncherLocalDir /grant "*S-1-5-32-545:(OI)(CI)M" /T *> $null
        } catch { }

        Write-Log "Launcher copiado/atualizado em $LauncherLocalPath."
    } catch {
        Write-Log "Falha ao copiar o launcher para local: $($_.Exception.Message)" "ERROR"
        return $true
    }

    # Tarefa local de sincronização periódica (contexto usuário, sem GPO)
    $taskNameSync = "FirstDecision-LauncherSync"
    $actionSync = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsPath`" `"$LauncherLocalPath`""
    $triggerSync = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $IntervaloMinutosSincronizacao) -RepetitionDuration (New-TimeSpan -Days 3650)
    $settingsSync = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    try {
        if (Get-ScheduledTask -TaskName $taskNameSync -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $taskNameSync -Confirm:$false -ErrorAction Stop
        }
        Register-ScheduledTask -TaskName $taskNameSync -Action $actionSync -Trigger $triggerSync -Settings $settingsSync -ErrorAction Stop | Out-Null
        Write-Log "Tarefa local '$taskNameSync' (sincronização a cada $IntervaloMinutosSincronizacao min, oculta) criada/atualizada."
    } catch {
        Write-Log "Falha ao criar a tarefa '$taskNameSync': $($_.Exception.Message)" "ERROR"
    }

    # Reescreve a AÇÃO da própria tarefa de logon criada pela GPO
    # ("FirstDecision - Launcher Popup Central"), trocando o caminho de
    # rede (UNC/SYSVOL) por um caminho 100% LOCAL. Sem isso, todo logon sem
    # conexão com o domínio (ex: home office sem VPN) mostra um erro nativo
    # do Windows ("O nome da rede especificado não está mais disponível"),
    # porque o wscript.exe nem consegue ABRIR o arquivo pela rede - isso
    # acontece ANTES de qualquer lógica interna de tolerância a falha rodar.
    $taskNameLogonGpo = "FirstDecision - Launcher Popup Central"
    try {
        if (Get-ScheduledTask -TaskName $taskNameLogonGpo -ErrorAction SilentlyContinue) {
            $actionLogonLocal = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsPath`" `"$LauncherLocalPath`""
            Set-ScheduledTask -TaskName $taskNameLogonGpo -Action $actionLogonLocal -ErrorAction Stop | Out-Null
            Write-Log "Tarefa de logon '$taskNameLogonGpo' reescrita para apontar ao caminho local (elimina dependência de rede nesse gatilho)."
        }
    } catch {
        Write-Log "Falha ao reescrever a tarefa de logon '$taskNameLogonGpo' para caminho local: $($_.Exception.Message)" "WARN"
    }

    return $true
}

function Remove-TarefasObsoletas {
    # Roda em TODA sincronização (não só no bootstrap) - garante que tarefas
    # antigas/aposentadas sejam removidas mesmo em máquinas que nunca mais
    # voltam a tocar no domínio.
    $nomesObsoletos = @(
        "FirstDecision-LauncherLocal-Logon", "FirstDecision-LauncherLocal-Periodico", "FirstDecision-LauncherLocal",
        "TASK POPUP - FD"   # sistema antigo de popup de bem-estar (GPO "POP_UP", substituído por Notificacao-BemEstar.ps1)
    )
    foreach ($nomeAntigo in $nomesObsoletos) {
        if (Get-ScheduledTask -TaskName $nomeAntigo -ErrorAction SilentlyContinue) {
            try {
                Unregister-ScheduledTask -TaskName $nomeAntigo -Confirm:$false -ErrorAction Stop
                Write-Log "Tarefa obsoleta '$nomeAntigo' removida."
            } catch {
                Write-Log "Falha ao remover a tarefa obsoleta '$nomeAntigo': $($_.Exception.Message)" "ERROR"
            }
        }
    }
    Get-ScheduledTask -TaskName "FD-Popups-*" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false -ErrorAction Stop
            Write-Log "Tarefa obsoleta (modelo anterior) '$($_.TaskName)' removida."
        } catch {
            Write-Log "Falha ao remover a tarefa obsoleta (modelo anterior) '$($_.TaskName)': $($_.Exception.Message)" "ERROR"
        }
    }
}

function Sync-DedicatedScriptTasks {
    param($Manifest, [string]$VbsPath)

    $schedulesByName = @{}
    foreach ($s in $Manifest.schedules) { $schedulesByName[$s.name] = $s }

    # Scripts desejados: precisam de run=true, um trigger que resolva pra um
    # schedule existente, e não serem o script de wallpaper (que é tratado à parte).
    $desired = @{}
    foreach ($item in $Manifest.scripts) {
        if ($item.run -ne $true) { continue }
        if ($item.name -eq $WallpaperScriptName) { continue }

        $sched = $schedulesByName[$item.trigger]
        if (-not $sched) {
            Write-Log "Script $($item.name) tem trigger '$($item.trigger)' sem schedule correspondente no manifest. Pulando tarefa dedicada." "WARN"
            continue
        }

        $shortName = [IO.Path]::GetFileNameWithoutExtension($item.name)
        $desired[$shortName] = @{ Item = $item; Schedule = $sched }
    }

    # Tarefas dedicadas existentes
    $existingTasks = Get-ScheduledTask -TaskName "$TaskPrefixDedicada*" -ErrorAction SilentlyContinue
    $existingNames = @($existingTasks | ForEach-Object { $_.TaskName })

    # Sem -StartWhenAvailable de propósito: se a máquina estiver desligada
    # no horário programado, o popup deve ser PULADO, não disparar tarde
    # quando a pessoa ligar depois (ex: sábado à noite). Diferente do
    # LauncherSync, que é manutenção silenciosa e pode recuperar atraso
    # sem problema, popups são visíveis e só fazem sentido no horário certo.
    $settingsDedicada = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    # Cria/atualiza as desejadas
    foreach ($shortName in $desired.Keys) {
        $d = $desired[$shortName]
        $taskName = "$TaskPrefixDedicada$shortName"

        $trigger = switch ($d.Schedule.type) {
            "logon" { New-ScheduledTaskTrigger -AtLogOn }
            "daily" {
                # $d.Schedule.time pode ser um horário único ("10:30") ou uma
                # lista de horários (["10:30", "15:30"]) - nesse caso, cria
                # múltiplos triggers na MESMA tarefa, em vez de uma tarefa por
                # horário. É assim que um popup roda 2x ao dia sem duplicar
                # a tarefa dedicada.
                $horarios = @($d.Schedule.time)
                $triggersDoDia = foreach ($h in $horarios) {
                    New-ScheduledTaskTrigger -Daily -At $h
                }
                $triggersDoDia
            }
            "weekly" { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $d.Schedule.day -At $d.Schedule.time }
            default {
                Write-Log "Tipo de schedule desconhecido '$($d.Schedule.type)' para '$($d.Item.name)'. Pulando." "WARN"
                $null
            }
        }
        if (-not $trigger) { continue }

        $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsPath`" `"$LauncherLocalPath`" -RunScript `"$($d.Item.name)`""

        try {
            if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
            }
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settingsDedicada -ErrorAction Stop | Out-Null
            Write-Log "Tarefa dedicada '$taskName' sincronizada ($($d.Schedule.type), $($d.Schedule.time), oculta)."
        } catch {
            Write-Log "Falha ao criar/atualizar a tarefa '$taskName': $($_.Exception.Message)" "ERROR"
        }
    }

    # Remove tarefas dedicadas que não existem mais no manifest (script
    # removido, desativado, ou trigger sem schedule válido)
    foreach ($existingName in $existingNames) {
        $shortName = $existingName -replace [regex]::Escape($TaskPrefixDedicada), ""
        if (-not $desired.ContainsKey($shortName)) {
            Unregister-ScheduledTask -TaskName $existingName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "Tarefa dedicada obsoleta '$existingName' removida (não está mais no manifest)."
        }
    }
}

function Send-CheckIn {
    # Uma vez por dia por máquina/usuário, avisa o fluxo de check-in (Power
    # Automate) que esta máquina está com o sistema de popups funcionando.
    # Serve só pra TI saber quem tem e quem falta. Falha aqui nunca
    # interrompe o resto da sincronização.
    param($Manifest)

    $url = $Manifest.config.checkinUrl
    if (-not $url) { return }

    $hoje = Get-Date -Format "yyyy-MM-dd"
    $marca = "$hoje|$($env:USERNAME)"
    if ((Test-Path $CheckinFile) -and ((Get-Content -Path $CheckinFile -Raw -ErrorAction SilentlyContinue) -match [regex]::Escape($marca))) {
        return   # já fez check-in hoje com este usuário
    }

    # Patrimônio = nome do computador sem o prefixo "NOTE" (ex: NOTE-1100 -> 1100).
    # Fora desse padrão, manda o nome inteiro pra não perder a linha.
    $computador = $env:COMPUTERNAME
    $patrimonio = if ($computador -match '^NOTE[-_ ]?(.+)$') { $Matches[1] } else { $computador }

    $upn = ""
    try { $upn = (& whoami.exe /upn 2>$null | Out-String).Trim() } catch { }

    $corpo = [ordered]@{
        computador      = $computador
        patrimonio      = $patrimonio
        usuario         = $env:USERNAME
        upn             = $upn
        versaoManifest  = "$($Manifest.version)"
        versaoLauncher  = (Get-LocalFileSha256 -Path $LauncherLocalPath).Substring(0, [Math]::Min(12, (Get-LocalFileSha256 -Path $LauncherLocalPath).Length))
        wallpaper       = (Get-LocalFileSha256 -Path (Join-Path $AssetsDir "wallpaper.jpg"))
        windows         = [System.Environment]::OSVersion.Version.ToString()
        dataHora        = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    } | ConvertTo-Json -Compress

    try {
        Invoke-RestMethod -Uri $url -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($corpo)) -ContentType "application/json; charset=utf-8" -TimeoutSec 30 | Out-Null
        Set-Content -Path $CheckinFile -Value $marca -Encoding UTF8
        Write-Log "Check-in diário enviado ($computador / $($env:USERNAME))."
    } catch {
        Write-Log "Falha ao enviar o check-in diário: $($_.Exception.Message)" "WARN"
    }
}

# ============================================================
# MODO RÁPIDO: -RunScript informado -> só executa e sai, sem tocar em rede.
# ============================================================
if ($RunScript) {
    Write-Log "===== Execução direta (RunScript: $RunScript) ====="
    $localScriptPath = Join-Path $CacheDir $RunScript
    if (-not (Test-Path $localScriptPath)) {
        Write-Log "$RunScript não encontrado no cache local ($localScriptPath)." "ERROR"
        exit 1
    }
    Invoke-CachedScript -ScriptPath $localScriptPath
    Write-Log "===== Fim da execução direta ====="
    exit 0
}

# ============================================================
# MODO COMPLETO: instala local (se preciso), sincroniza manifest/scripts/
# assets, aplica wallpaper, reconcilia tarefas dedicadas, faz check-in.
# ============================================================
Write-Log "===== Início da sincronização completa ====="

# raw.githubusercontent.com exige TLS 1.2; o PowerShell 5.1 nem sempre usa
# por padrão em Windows mais antigos.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ===================== PONTE: ANEL DE MIGRAÇÃO =====================
# O repositório antigo tem o arquivo ponte-anel.txt com os nomes dos
# computadores liberados para migrar ("*" = todos). Quem não está na lista
# continua exatamente como antes: não sincroniza com o repositório novo, não
# mexe em tarefas e só mantém a própria ponte atualizada pelo repositório
# antigo. Se o arquivo não puder ser lido, ninguém migra.
$OldRawBase = "https://raw.githubusercontent.com/itsmatheus145/popup-scripts-test/main"
function Get-OldRepoBytes {
    param([string]$Path)
    $wc = New-Object System.Net.WebClient
    try {
        $wc.Headers.Add("User-Agent", "FirstDecision-PopupLauncher")
        return $wc.DownloadData("$OldRawBase/${Path}?v=$(Get-Date -Format 'yyyyMMddHHmmss')")
    } catch {
        return $null
    } finally { $wc.Dispose() }
}
function Test-AnelMigracao {
    $b = Get-OldRepoBytes "ponte-anel.txt"
    if (-not $b) { return $false }
    $linhas = [System.Text.Encoding]::UTF8.GetString($b).TrimStart([char]0xFEFF) -split "`r?`n" |
        ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith("#") }
    return (($linhas -contains "*") -or ($linhas -contains $env:COMPUTERNAME))
}
if (-not (Test-AnelMigracao)) {
    Write-Log "Ponte: computador fora do anel de migração - nada a fazer nesta execução."
    $remoto = Get-OldRepoBytes "Update-PopupScripts.ps1"
    if ($remoto -and ((Get-Sha256 -Bytes $remoto) -ne (Get-LocalFileSha256 -Path $LauncherLocalPath))) {
        try {
            [System.IO.File]::WriteAllBytes($LauncherLocalPath, $remoto)
            Write-Log "Ponte: atualizada a partir do repositório antigo."
        } catch {
            Write-Log "Ponte: não foi possível atualizar: $($_.Exception.Message)" "WARN"
        }
    }
    Write-Log "===== Fim da sincronização (fora do anel) ====="
    exit 0
}
$vbsPath = Install-RunHiddenWrapper
$vieioDeRemoto = Install-LocalLauncher -VbsPath $vbsPath -ForcarBootstrap $Bootstrap.IsPresent

Remove-TarefasObsoletas

if ($vieioDeRemoto -and $DelayPosConexaoSegundos -gt 0) {
    Write-Log "Aguardando $DelayPosConexaoSegundos segundos (rede subindo pós-conexão)."
    Start-Sleep -Seconds $DelayPosConexaoSegundos
}

if (-not (Test-Path $CacheDir)) { New-Item -Path $CacheDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $AssetsDir)) { New-Item -Path $AssetsDir -ItemType Directory -Force | Out-Null }

# Manifest sempre fresco: cache-buster pelo horário, pra ligar/desligar um
# popup ou mudar horário valer na próxima sincronização, sem esperar o CDN.
$manifestBytes = Get-RemoteFileBytes -RepoPath "manifest.json" -CacheBuster (Get-Date -Format "yyyyMMddHHmmss")
if (-not $manifestBytes) {
    Write-Log "Não foi possível obter o manifest remoto. Abortando execução (cache e tarefas atuais continuam valendo)." "ERROR"
    exit 1
}

$bom = [System.Text.Encoding]::UTF8.GetPreamble()
if ($manifestBytes.Length -ge 3 -and $manifestBytes[0] -eq $bom[0] -and $manifestBytes[1] -eq $bom[1] -and $manifestBytes[2] -eq $bom[2]) {
    $manifestBytes = $manifestBytes[3..($manifestBytes.Length - 1)]
}
$manifestRaw = [System.Text.Encoding]::UTF8.GetString($manifestBytes)

try {
    $manifest = $manifestRaw | ConvertFrom-Json
} catch {
    Write-Log "Manifest inválido (JSON malformado). Abortando (cache e tarefas atuais continuam valendo)." "ERROR"
    exit 1
}

Write-Log "Versão do manifest remoto: $($manifest.version)"

try {
    [System.IO.File]::WriteAllText($ManifestLocal, $manifestRaw, (New-Object System.Text.UTF8Encoding($false)))
} catch {
    Write-Log "Falha ao salvar a cópia local do manifest: $($_.Exception.Message)" "WARN"
}

Update-LauncherFromGitHub -Manifest $manifest

# Scripts: baixa só os que mudaram (hash local diferente do manifest). Os
# com run=false não são baixados, mas a cópia em cache não é apagada - ao
# religar, volta a valer na hora.
$tudoOk = $true
foreach ($item in $manifest.scripts) {
    if ($item.run -ne $true) { continue }
    $ok = Save-VerifiedFile -RepoPath "scripts/$($item.name)" -LocalPath (Join-Path $CacheDir $item.name) -ExpectedSha256 $item.sha256
    if (-not $ok) { $tudoOk = $false }
}

foreach ($asset in $manifest.assets) {
    $ok = Save-VerifiedFile -RepoPath $asset.repoPath -LocalPath (Join-Path $AssetsDir $asset.localName) -ExpectedSha256 $asset.sha256
    if (-not $ok) { $tudoOk = $false }
}

if ($tudoOk) {
    Set-Content -Path $VersionFile -Value $manifest.version -Encoding UTF8
} else {
    Write-Log "Algum arquivo não pôde ser atualizado nesta sincronização - tenta de novo na próxima." "WARN"
}

# Aplica o wallpaper agora (contexto usuário desta execução) - garante que
# qualquer imagem nova baixada acima já reflita na área de trabalho.
$wallpaperScriptPath = Join-Path $CacheDir $WallpaperScriptName
if (Test-Path $wallpaperScriptPath) {
    Invoke-CachedScript -ScriptPath $wallpaperScriptPath
} else {
    Write-Log "$WallpaperScriptName não encontrado no cache - pulando aplicação de wallpaper desta vez." "WARN"
}

# Reconcilia as tarefas dedicadas (uma por script agendado) com o manifest -
# roda SEMPRE, mudando ou não a versão, garantindo "auto-cura" a cada
# sincronização (tarefa apagada na mão volta, horário mudado é corrigido,
# script desligado ou removido do manifest tem sua tarefa apagada).
Sync-DedicatedScriptTasks -Manifest $manifest -VbsPath $vbsPath

Send-CheckIn -Manifest $manifest

Write-Log "===== Fim da sincronização completa ====="
