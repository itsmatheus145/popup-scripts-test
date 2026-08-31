<#
.SYNOPSIS
    Launcher central - baixa manifest.json de um repositório GitHub, mantém os
    scripts e a imagem de wallpaper atualizados localmente, e garante que
    cada script tenha sua PRÓPRIA Tarefa Agendada dedicada, no horário
    definido no manifest.

.DESCRIPTION
    Modelo de execução:
      - Chamado SEM -RunScript: modo "sincronização completa". Faz tudo:
        instala a cópia local (se necessário), baixa manifest/scripts/assets
        se a versão mudou, aplica o wallpaper, e reconcilia (cria/atualiza/
        remove) uma Tarefa Agendada dedicada para cada script do manifest.
        É este modo que a GPO chama (via SYSVOL) e que a tarefa horária local
        "FirstDecision-LauncherSync" chama.

      - Chamado COM -RunScript <nome.ps1>: modo "execução direta e rápida".
        Não baixa nada, não mexe em tarefas - só executa o arquivo já
        cacheado localmente. É este modo que cada tarefa dedicada
        (FD-Task-<script>) usa para disparar seu próprio popup, sem
        depender de rede nem repetir toda a lógica de sincronização.

    "Auto-cura": toda vez que a máquina se conecta à GPO (ou a tarefa local
    horária dispara), a reconciliação roda de novo - qualquer tarefa apagada
    manualmente é recriada, qualquer mudança de horário no manifest é
    aplicada, e qualquer script removido do manifest tem sua tarefa apagada.

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
$RepoOwner   = "itsmatheus145"
$RepoName    = "popup-scripts-test"
$Branch      = "main"

# O token NÃO fica mais escrito aqui dentro - esse arquivo agora pode ir pro
# Git sem risco (é código puro, sem segredo). O token vive num arquivo à
# parte, que só existe no SYSVOL e é copiado pra máquina uma vez (assim como
# este script) - nunca é versionado, nunca aparece num commit.
$TokenFilePath = Join-Path "C:\ProgramData\FirstDecision\Launcher" "github_token.dat"

$CacheDir       = "C:\ProgramData\FirstDecision\PopupScripts"
$AssetsDir      = Join-Path $CacheDir "Assets"
$LogDir         = Join-Path $CacheDir "Logs"
$VersionFile    = Join-Path $CacheDir "current_version.txt"
$ApiBase        = "https://api.github.com/repos/$RepoOwner/$RepoName/contents"

# Caminho local fixo do launcher. Depois da primeira execução via GPO/SYSVOL,
# todas as tarefas locais apontam pra cá - nunca mais precisam do domínio.
$LauncherLocalDir  = "C:\ProgramData\FirstDecision\Launcher"
$LauncherLocalPath = Join-Path $LauncherLocalDir "Update-PopupScripts.ps1"

# Nome do script de wallpaper - tratado à parte dos demais (não tem tarefa
# dedicada própria; é baixado e aplicado a cada sincronização, porque deve
# refletir qualquer imagem nova o quanto antes).
$WallpaperScriptName = "Atualizar-Wallpaper.ps1"

# Intervalo (em horas) da tarefa local recorrente de sincronização completa.
$IntervaloHorasSincronizacao = 1

# Atraso (em segundos) só quando esta execução partiu de um caminho remoto
# (SYSVOL/GPO) - dá tempo da rede subir logo após conectar/logar.
$DelayPosConexaoSegundos = 60

# Nome do arquivo do próprio launcher dentro do repositório (raiz, separado
# da pasta scripts/ que é só para os scripts "de carga" tipo popups).
# Usado para o launcher se auto-atualizar via GitHub, sem depender de nunca
# mais tocar no domínio/SYSVOL. Este arquivo em si NÃO carrega o token -
# só código.
# mais tocar no domínio/SYSVOL.
$LauncherRepoPath = "Update-PopupScripts.ps1"

$TaskPrefixDedicada = "FD-Task-"
# ==========================================================

function Get-GitHubToken {
    if (-not (Test-Path $TokenFilePath)) {
        Write-Log "Arquivo de token não encontrado em $TokenFilePath. Sincronização com o GitHub vai falhar até isso ser corrigido." "ERROR"
        return $null
    }
    try {
        return (Get-Content -Path $TokenFilePath -Raw).Trim()
    } catch {
        Write-Log "Falha ao ler o arquivo de token: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
        # Mesma proteção já aplicada ao diretório do Launcher: garante que
        # usuários comuns sempre consigam escrever aqui, mesmo que esta
        # criação específica tenha acontecido numa sessão elevada.
        try { & icacls.exe $LogDir /grant "Usuários:(OI)(CI)M" /T *> $null } catch { }
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
            & icacls.exe $logFile /grant "Usuários:M" *> $null
            Add-Content -Path $logFile -Value $line -ErrorAction Stop
        } catch { }
    }
    Write-Host $line
}

function Get-GitHubFileBytes {
    param([string]$RepoPath, [hashtable]$Headers)
    try {
        $segmentosCodificados = $RepoPath -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }
        $repoPathCodificado = $segmentosCodificados -join '/'
        $url = "$ApiBase/${repoPathCodificado}?ref=$Branch"
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -TimeoutSec 15
        return [System.Convert]::FromBase64String($resp.content)
    } catch {
        Write-Log "Falha ao baixar $RepoPath : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-GitHubFileContent {
    param([string]$RepoPath, [hashtable]$Headers)
    $bytes = Get-GitHubFileBytes -RepoPath $RepoPath -Headers $Headers
    if (-not $bytes) { return $null }
    $bom = [System.Text.Encoding]::UTF8.GetPreamble()
    if ($bytes.Length -ge $bom.Length -and (Compare-Object $bytes[0..($bom.Length-1)] $bom -SyncWindow 0) -eq $null) {
        $bytes = $bytes[$bom.Length..($bytes.Length-1)]
    }
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Update-LauncherFromGitHub {
    param([hashtable]$Headers)

    # Compara o conteúdo do launcher no GitHub com a cópia local (por hash,
    # não por número de versão - qualquer mudança de conteúdo já dispara a
    # atualização). Isso garante que máquinas que nunca mais voltam a tocar
    # no domínio/SYSVOL ainda recebem correções futuras no próprio launcher,
    # não só nos scripts que ele gerencia.
    $remoteBytes = Get-GitHubFileBytes -RepoPath $LauncherRepoPath -Headers $Headers
    if (-not $remoteBytes) {
        Write-Log "Não foi possível verificar atualização do próprio launcher (falha ao baixar). Mantendo versão atual." "WARN"
        return
    }

    if (-not (Test-Path $LauncherLocalPath)) {
        return  # Install-LocalLauncher ainda não rodou nesta execução - nada a comparar
    }

    $localBytes = [System.IO.File]::ReadAllBytes($LauncherLocalPath)
    $hashRemoto = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash($remoteBytes))
    $hashLocal  = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash($localBytes))

    if ($hashRemoto -eq $hashLocal) {
        Write-Log "Launcher já está na versão mais recente (hash confere com o GitHub)."
        return
    }

    try {
        [System.IO.File]::WriteAllBytes($LauncherLocalPath, $remoteBytes)
        Write-Log "Launcher atualizado via GitHub (o código do próprio Update-PopupScripts.ps1 mudou). A versão nova entra em vigor na próxima execução agendada."
    } catch {
        Write-Log "Falha ao gravar a nova versão do launcher: $($_.Exception.Message)" "ERROR"
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
    $vbsContent = @'
Dim shell, cmd, i
Set shell = CreateObject("WScript.Shell")
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
        Copy-Item -Path $PSCommandPath -Destination $LauncherLocalPath -Force

        # Reforço: garante que usuários comuns sempre consigam escrever nessa
        # pasta, mesmo que esta execução específica tenha rodado elevada (o
        # que normalmente restringiria a ACL só a Administradores). Sem isso,
        # a próxima sincronização não-elevada falha com "Acesso negado".
        try {
            & icacls.exe $LauncherLocalDir /grant "Usuários:(OI)(CI)M" /T *> $null
        } catch { }

        Write-Log "Launcher copiado/atualizado em $LauncherLocalPath."
    } catch {
        Write-Log "Falha ao copiar o launcher para local: $($_.Exception.Message)" "ERROR"
        return $true
    }

    # Copia o arquivo de token (vive como "arquivo irmão" do script no SYSVOL,
    # NUNCA no Git) para o local fixo - mesma lógica do próprio script: só
    # precisa existir uma vez, na primeira conexão com o domínio.
    try {
        $tokenSourcePath = Join-Path (Split-Path -Parent $PSCommandPath) "github_token.dat"
        if (Test-Path $tokenSourcePath) {
            Copy-Item -Path $tokenSourcePath -Destination $TokenFilePath -Force
            Write-Log "Arquivo de token copiado/atualizado a partir do SYSVOL."
        } else {
            Write-Log "Arquivo 'github_token.dat' não encontrado ao lado do launcher no SYSVOL ($tokenSourcePath). Verifique se ele foi publicado." "ERROR"
        }
    } catch {
        Write-Log "Falha ao copiar o arquivo de token: $($_.Exception.Message)" "ERROR"
    }

    # Tarefa local de sincronização periódica (contexto usuário, sem GPO)
    $taskNameSync = "FirstDecision-LauncherSync"
    $actionSync = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsPath`" `"$LauncherLocalPath`""
    $triggerSync = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours $IntervaloHorasSincronizacao) -RepetitionDuration (New-TimeSpan -Days 3650)
    $settingsSync = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    try {
        if (Get-ScheduledTask -TaskName $taskNameSync -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $taskNameSync -Confirm:$false -ErrorAction Stop
        }
        Register-ScheduledTask -TaskName $taskNameSync -Action $actionSync -Trigger $triggerSync -Settings $settingsSync -ErrorAction Stop | Out-Null
        Write-Log "Tarefa local '$taskNameSync' (sincronização a cada $IntervaloHorasSincronizacao h, oculta) criada/atualizada."
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
    # Como esta tarefa é de propriedade do próprio usuário (não SYSTEM/GPO
    # admin), consegue ser reescrita sem precisar de privilégio elevado.
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
    # voltam a tocar no domínio. Se uma GPO antiga ainda estiver ativa e
    # recriando alguma dessas tarefas, elas voltam a ser removidas na
    # próxima sincronização (a cada 1h) - só "vencem" de vez quando a GPO
    # de origem for desativada/desvinculada.
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
        if (-not $item.run) { continue }
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

    $settingsDedicada = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

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
# assets, aplica wallpaper, reconcilia tarefas dedicadas.
# ============================================================
Write-Log "===== Início da sincronização completa ====="

$vbsPath = Install-RunHiddenWrapper
$vieioDeRemoto = Install-LocalLauncher -VbsPath $vbsPath -ForcarBootstrap $Bootstrap.IsPresent

Remove-TarefasObsoletas

if ($vieioDeRemoto -and $DelayPosConexaoSegundos -gt 0) {
    Write-Log "Aguardando $DelayPosConexaoSegundos segundos (rede subindo pós-conexão)."
    Start-Sleep -Seconds $DelayPosConexaoSegundos
}

if (-not (Test-Path $CacheDir)) { New-Item -Path $CacheDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $AssetsDir)) { New-Item -Path $AssetsDir -ItemType Directory -Force | Out-Null }

$headers = @{
    "User-Agent" = "FirstDecision-PopupLauncher"
    "Accept"     = "application/vnd.github.v3+json"
}
$GitHubToken = Get-GitHubToken
if (-not $GitHubToken) {
    Write-Log "Sem token disponível - abortando sincronização." "ERROR"
    exit 1
}
$headers["Authorization"] = "token $GitHubToken"

Update-LauncherFromGitHub -Headers $headers

$manifestRaw = Get-GitHubFileContent -RepoPath "manifest.json" -Headers $headers
if (-not $manifestRaw) {
    Write-Log "Não foi possível obter o manifest remoto. Abortando execução." "ERROR"
    exit 1
}

try {
    $manifest = $manifestRaw | ConvertFrom-Json
} catch {
    Write-Log "Manifest inválido (JSON malformado). Abortando." "ERROR"
    exit 1
}

Write-Log "Versão do manifest remoto: $($manifest.version)"

$localVersion = if (Test-Path $VersionFile) { (Get-Content $VersionFile -Raw).Trim() } else { "" }
$versaoMudou = ($localVersion -ne $manifest.version.Trim())

if (-not $versaoMudou) {
    Write-Log "Versão local já está atualizada ($($manifest.version)). Usando cache para scripts/assets."
} else {
    Write-Log "Nova versão detectada (local: '$localVersion' -> remoto: '$($manifest.version)'). Baixando scripts e assets..."

    foreach ($item in $manifest.scripts) {
        if (-not $item.run) {
            Write-Log "Script $($item.name) marcado como run=false. Pulando download."
            continue
        }
        Write-Log "Baixando $($item.name)..."
        $rawBytes = Get-GitHubFileBytes -RepoPath "scripts/$($item.name)" -Headers $headers
        if (-not $rawBytes) {
            Write-Log "Falha ao baixar $($item.name). Mantendo versão em cache anterior (se existir)." "ERROR"
            continue
        }
        $localScriptPath = Join-Path $CacheDir $item.name
        [System.IO.File]::WriteAllBytes($localScriptPath, $rawBytes)
        Write-Log "$($item.name) baixado com sucesso."
    }

    foreach ($asset in $manifest.assets) {
        Write-Log "Baixando asset $($asset.repoPath)..."
        $rawBytes = Get-GitHubFileBytes -RepoPath $asset.repoPath -Headers $headers
        if (-not $rawBytes) {
            Write-Log "Falha ao baixar asset $($asset.repoPath). Mantendo versão em cache anterior (se existir)." "ERROR"
            continue
        }
        $localAssetPath = Join-Path $AssetsDir $asset.localName
        [System.IO.File]::WriteAllBytes($localAssetPath, $rawBytes)
        Write-Log "Asset $($asset.repoPath) baixado com sucesso -> $localAssetPath"
    }

    Set-Content -Path $VersionFile -Value $manifest.version -Encoding UTF8
}

# Baixa o script de wallpaper sempre (independente de mudança de versão, é
# leve e garante que ele exista mesmo em cenários de recuperação)
$wallpaperScriptPath = Join-Path $CacheDir $WallpaperScriptName
$rawWallpaperScript = Get-GitHubFileBytes -RepoPath "scripts/$WallpaperScriptName" -Headers $headers
if ($rawWallpaperScript) {
    [System.IO.File]::WriteAllBytes($wallpaperScriptPath, $rawWallpaperScript)
}

# Aplica o wallpaper agora (contexto usuário desta execução) - garante que
# qualquer imagem nova baixada acima já reflita na área de trabalho.
if (Test-Path $wallpaperScriptPath) {
    Invoke-CachedScript -ScriptPath $wallpaperScriptPath
} else {
    Write-Log "$WallpaperScriptName não encontrado no cache - pulando aplicação de wallpaper desta vez." "WARN"
}

# Reconcilia as tarefas dedicadas (uma por script agendado) com o manifest -
# roda SEMPRE, mudando ou não a versão, garantindo "auto-cura" a cada
# sincronização (tarefa apagada na mão volta, horário mudado é corrigido,
# script removido do manifest tem sua tarefa apagada).
Sync-DedicatedScriptTasks -Manifest $manifest -VbsPath $vbsPath

Write-Log "===== Fim da sincronização completa ====="
