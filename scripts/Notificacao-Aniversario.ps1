<#
.SYNOPSIS
    Popup de aniversario - so aparece para o colaborador que faz aniversario
    no dia, consultando uma planilha Excel no SharePoint/OneDrive via Graph API.

.DESCRIPTION
    1. Autentica no Graph API via Client Credentials (app registration dedicado,
       somente leitura).
    2. Le a Tabela do Excel "Table1" via Tables API (independe do nome da aba).
    3. Compara a coluna "Nome de Usuario" com o usuario de rede logado na maquina.
    4. Se o dia/mes da coluna "Data de Nascimento" bater com hoje (ano ignorado),
       mostra o popup. Caso contrario, o script termina silenciosamente.

.NOTAS
    Rodando via launcher central puxado do Git (trigger diario, ex: daily_0900).
    Arquivo salvo em UTF-8 com BOM.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ====================== CONFIGURACAO - GRAPH API ======================
$TenantId     = "b240a0cd-a9bc-4363-8632-cbf35140a023"
$ClientId     = "700884ff-a0bc-4dbe-892d-ca5c94b116de"
$ClientSecret = "duX8Q~N-q1YTonIyqmHiyRM8uY8T7W-n6gu4mbVT"
                # ATENCAO: mover para Windows Credential Manager ou variavel
                # de ambiente de maquina antes de qualquer uso em producao.

# Link de compartilhamento da planilha (OneDrive pessoal do RH).
# Cole aqui o mesmo link que você usa para abrir a planilha no navegador.
$SharingUrl = "https://firstdecisioncombr-my.sharepoint.com/:x:/g/personal/dayanne_ferreira_firstdecision_com_br/IQDHMx49JdY3Ro_Szy5UbYwgAf2TTnVdx23DBJcj52QoBQY?e=51aaGF"

$TableName       = "Table1"                 # nome da Tabela do Excel (aparece na aba embaixo da planilha)
$ColUsuarioNome  = "Nome de Usuario"         # cabeçalho exato da coluna de usuário
$ColDataNome     = "Data de Nascimento"      # cabeçalho exato da coluna de data (formato DD/MM/AAAA)

# ====================== CONFIGURACAO - MENSAGEM DO POPUP ======================
$Linha1         = "Feliz aniversário!"
$Linha2         = "Que seu novo ciclo seja repleto de alegrias, conquistas e bons momentos."
$TempoAutoFecha = 60

$CorFundo       = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)
$CorTextoTitulo = [System.Drawing.Color]::FromArgb(255, 1, 28, 83)
$CorX           = [System.Drawing.Color]::FromArgb(255, 150, 160, 180)
$CorXHover      = [System.Drawing.Color]::FromArgb(255, 1, 28, 83)
$CorBorda       = [System.Drawing.Color]::FromArgb(255, 225, 227, 235)

$LogDir = "C:\ProgramData\FirstDecision\PopupScripts\Logs"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
    $logFile = Join-Path $LogDir ("execucao_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $line = "[{0}] [{1}] [Aniversario] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $logFile -Value $line
}

# ====================== 1. AUTENTICACAO NO GRAPH (Client Credentials) ======================
function Get-GraphToken {
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    try {
        $resp = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $resp.access_token
    } catch {
        Write-Log "Falha ao obter token do Graph: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

$token = Get-GraphToken
if (-not $token) {
    Write-Log "Sem token - abortando verificação de aniversário." "ERROR"
    exit 1
}

$headers = @{ Authorization = "Bearer $token" }

# ====================== 2. RESOLVE O ARQUIVO A PARTIR DO LINK DE COMPARTILHAMENTO ======================
function ConvertTo-GraphShareId {
    param([string]$Url)
    # Codifica a URL em base64url no formato que o Graph exige para /shares/{id}
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Url)
    $b64   = [System.Convert]::ToBase64String($bytes)
    $b64url = $b64.TrimEnd('=').Replace('/', '_').Replace('+', '-')
    return "u!$b64url"
}

try {
    $shareId  = ConvertTo-GraphShareId -Url $SharingUrl
    $itemUrl  = "https://graph.microsoft.com/v1.0/shares/$shareId/driveItem?`$select=id,parentReference"
    $driveItem = Invoke-RestMethod -Uri $itemUrl -Headers $headers -Method Get

    $driveId = $driveItem.parentReference.driveId
    $itemId  = $driveItem.id
} catch {
    Write-Log "Falha ao resolver o arquivo a partir do link de compartilhamento: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ====================== 3. DESCOBRE A TABELA E LE O CONTEUDO (Tables API) ======================
try {
    $tablesUrl  = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/workbook/tables"
    $tablesList = Invoke-RestMethod -Uri $tablesUrl -Headers $headers -Method Get
} catch {
    Write-Log "Falha ao listar as tabelas do arquivo: $($_.Exception.Message)" "ERROR"
    exit 1
}

$nomesDisponiveis = $tablesList.value | ForEach-Object { $_.name }
Write-Log "Tabelas encontradas no arquivo: $($nomesDisponiveis -join ', ')"

if ($tablesList.value.Count -eq 1) {
    $TableId = $tablesList.value[0].id
    Write-Log "Usando a única tabela existente (id: $TableId, nome: '$($tablesList.value[0].name)')."
} else {
    $match = $tablesList.value | Where-Object { $_.name -eq $TableName }
    if ($match) {
        $TableId = $match.id
    } else {
        Write-Log "Tabela '$TableName' não encontrada entre as disponíveis: $($nomesDisponiveis -join ', '). Ajuste `$TableName no script." "ERROR"
        exit 1
    }
}

try {
    $headerUrl  = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/workbook/tables/$TableId/headerRowRange"
    $headerData = Invoke-RestMethod -Uri $headerUrl -Headers $headers -Method Get

    $rowsUrl  = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/workbook/tables/$TableId/rows?`$select=values"
    $rowsData = Invoke-RestMethod -Uri $rowsUrl -Headers $headers -Method Get
} catch {
    Write-Log "Falha ao ler a tabela (id: $TableId): $($_.Exception.Message)" "ERROR"
    exit 1
}

$headerRow = $headerData.text[0]
$colUsuario = [Array]::IndexOf($headerRow, $ColUsuarioNome)
$colData    = [Array]::IndexOf($headerRow, $ColDataNome)

if ($colUsuario -lt 0 -or $colData -lt 0) {
    Write-Log "Colunas '$ColUsuarioNome' e/ou '$ColDataNome' não encontradas no cabeçalho da tabela. Cabeçalho real: $($headerRow -join ' | ')" "ERROR"
    exit 1
}

$dataRows = $rowsData.value
if (-not $dataRows -or $dataRows.Count -eq 0) {
    Write-Log "Tabela sem linhas de dados." "WARN"
    exit 0
}

# ====================== 4. IDENTIFICA O USUARIO LOGADO E A DATA DE HOJE ======================
function Get-DiaMesDaCelula {
    param($Valor)

    # Caso 1: número serial de data do Excel (dias desde 30/12/1899)
    $numero = 0.0
    if ([double]::TryParse($Valor.ToString(), [ref]$numero)) {
        try {
            $dt = [DateTime]::FromOADate($numero)
            return @{ dia = $dt.Day; mes = $dt.Month }
        } catch {
            return $null
        }
    }

    # Caso 2: texto em formato DD/MM/AAAA ou ISO AAAA-MM-DD
    $t = $Valor.ToString().Trim()
    if ($t -match '^(\d{1,2})/(\d{1,2})/\d{2,4}') {
        return @{ dia = [int]$Matches[1]; mes = [int]$Matches[2] }
    }
    if ($t -match '^\d{4}-(\d{1,2})-(\d{1,2})') {
        return @{ dia = [int]$Matches[2]; mes = [int]$Matches[1] }
    }
    return $null
}

$usuarioAtual = $env:USERNAME.ToLower()
$hoje = Get-Date
$hojeDia = $hoje.Day
$hojeMes = $hoje.Month

Write-Log "Verificando aniversário para usuário '$usuarioAtual' (hoje: $($hoje.ToString('dd/MM')))..."

$encontrouAniversario = $false

foreach ($row in $dataRows) {
    $valores = $row.values[0]   # cada "row" tem values como matriz 1xN

    $linhaUsuarioRaw = $valores[$colUsuario]
    if (-not $linhaUsuarioRaw) { continue }
    $linhaUsuario = $linhaUsuarioRaw.ToString().Trim().ToLower()

    $linhaDataRaw = $valores[$colData]

    # DIAGNOSTICO: loga o valor bruto exato da célula quando é a linha do usuário atual
    if ($linhaUsuario -eq $usuarioAtual) {
        Write-Log "DIAGNOSTICO - linha do usuário '$usuarioAtual' encontrada. Valor bruto da data: '$linhaDataRaw' (tipo: $($linhaDataRaw.GetType().Name))"
    }

    if (-not $linhaDataRaw) { continue }

    $dm = Get-DiaMesDaCelula -Valor $linhaDataRaw
    if (-not $dm) {
        if ($linhaUsuario -eq $usuarioAtual) {
            Write-Log "DIAGNOSTICO - não foi possível interpretar a data '$linhaDataRaw' com os formatos esperados (serial Excel, DD/MM/AAAA ou ISO)." "WARN"
        }
        continue
    }

    if ($linhaUsuario -eq $usuarioAtual) {
        Write-Log "DIAGNOSTICO - data interpretada: dia=$($dm.dia) mes=$($dm.mes) | hoje: dia=$hojeDia mes=$hojeMes"
    }

    if ($linhaUsuario -eq $usuarioAtual -and $dm.dia -eq $hojeDia -and $dm.mes -eq $hojeMes) {
        $encontrouAniversario = $true
        break
    }
}

if (-not $encontrouAniversario) {
    Write-Log "Hoje não é aniversário de '$usuarioAtual'. Encerrando sem exibir popup."
    exit 0
}

Write-Log "Aniversário confirmado para '$usuarioAtual'. Exibindo popup..."

# ====================== 5. MONTA E EXIBE O POPUP (mesmo estilo do FirstNews) ======================
$LogoBase64 = "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"

$logoBytes  = [Convert]::FromBase64String($LogoBase64)
$logoStream = New-Object System.IO.MemoryStream(,$logoBytes)
$logoImage  = [System.Drawing.Image]::FromStream($logoStream)

$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.StartPosition   = 'Manual'
$form.TopMost          = $true
$form.ShowInTaskbar    = $false
$form.Size             = New-Object System.Drawing.Size(380, 175)
$form.BackColor        = $CorFundo

$screen  = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$xFinal  = $screen.Right  - $form.Width  - 16
$yFinal  = $screen.Bottom - $form.Height - 16
$yInicio = $screen.Bottom + 10

$form.Location = New-Object System.Drawing.Point($xFinal, $yInicio)

# ====================== FOGOS DE ARTIFICIO (janela separada, acima do popup) ======================
# Como o popup em si tem fundo branco, os fogos precisam de uma janela propria,
# transparente (TransparencyKey), posicionada acima, com leve sobreposicao no
# topo do popup para dar a sensacao de que os fogos "saem" dele.
$FogosAltura = 150
$FogosLargura = $form.Width + 20
$overlay = New-Object System.Windows.Forms.Form
$overlay.FormBorderStyle = 'None'
$overlay.StartPosition   = 'Manual'
$overlay.TopMost         = $true
$overlay.ShowInTaskbar   = $false
$overlay.Size            = New-Object System.Drawing.Size($FogosLargura, $FogosAltura)
$overlay.BackColor       = [System.Drawing.Color]::Black
$overlay.TransparencyKey = [System.Drawing.Color]::Black
$overlayX = $xFinal - 10
$overlay.Location        = New-Object System.Drawing.Point($overlayX, ($yFinal - $FogosAltura + 18))

# Ativa double buffering (propriedade protegida por padrao) para eliminar o
# "piscar" do fundo preto real por trás da transparência a cada redesenho.
$overlay.GetType().InvokeMember(
    'DoubleBuffered',
    [System.Reflection.BindingFlags]::SetProperty -bor [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic,
    $null, $overlay, @($true)
) | Out-Null

$CoresFogos = @(
    [System.Drawing.Color]::FromArgb(255, 255, 215, 0),    # dourado
    [System.Drawing.Color]::FromArgb(255, 255, 165, 0),    # laranja-dourado
    [System.Drawing.Color]::FromArgb(255, 255, 255, 200),  # branco-amarelado
    [System.Drawing.Color]::FromArgb(255, 255, 190, 60)    # amarelo-ouro
)

$Particulas = New-Object System.Collections.Generic.List[hashtable]
$Foguetes   = New-Object System.Collections.Generic.List[hashtable]

function Add-ExplosaoFogos {
    param([double]$Cx, [double]$Cy, [double]$Escala = 1.0)
    $qtdParticulas = [int]((Get-Random -Minimum 45 -Maximum 70) * $Escala)
    for ($i = 0; $i -lt $qtdParticulas; $i++) {
        $angulo     = (Get-Random -Minimum 0 -Maximum 360) * [math]::PI / 180
        $velocidade = ((Get-Random -Minimum 25 -Maximum 75) / 10.0) * $Escala
        $vida       = Get-Random -Minimum 35 -Maximum 55
        $tam        = (Get-Random -Minimum 5 -Maximum 10) * $Escala
        $script:Particulas.Add(@{
            X        = $Cx
            Y        = $Cy
            VX       = [math]::Cos($angulo) * $velocidade
            VY       = [math]::Sin($angulo) * $velocidade
            Vida     = $vida
            VidaMax  = $vida
            Tamanho  = $tam
            Cor      = $CoresFogos[(Get-Random -Minimum 0 -Maximum $CoresFogos.Count)]
        })
    }
}

function Add-Foguete {
    $xLancamento = Get-Random -Minimum 40 -Maximum ($overlay.Width - 40)
    $yInicial    = $FogosAltura - 10   # nasce bem perto do topo do popup
    $yAlvo       = Get-Random -Minimum ($FogosAltura * 0.15) -Maximum ($FogosAltura * 0.40)
    $script:Foguetes.Add(@{
        X = $xLancamento; Y = $yInicial; YAlvo = $yAlvo; VY = -6.5
    })
}

$overlay.Add_Paint({
    param($s, $e)

    # Rastro dos foguetes subindo
    foreach ($f in $Foguetes) {
        $brushRastro = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 255, 230, 140))
        $e.Graphics.FillEllipse($brushRastro, [single]($f.X - 2), [single]$f.Y, 4, 10)
        $brushRastro.Dispose()
    }

    foreach ($p in $Particulas) {
        $fracaoVida = $p.Vida / [double]$p.VidaMax
        $alpha = [int](255 * $fracaoVida)
        if ($alpha -lt 0) { $alpha = 0 }
        if ($alpha -gt 255) { $alpha = 255 }
        $corBase = $p.Cor
        $tamanhoAtual = [single]($p.Tamanho * $fracaoVida)
        if ($tamanhoAtual -lt 0.5) { continue }

        # Brilho (glow) por trás - um circulo maior e bem translucido
        $alphaGlow = [int]($alpha * 0.35)
        $corGlow = [System.Drawing.Color]::FromArgb($alphaGlow, $corBase.R, $corBase.G, $corBase.B)
        $brushGlow = New-Object System.Drawing.SolidBrush($corGlow)
        $tamGlow = $tamanhoAtual * 2.2
        $e.Graphics.FillEllipse($brushGlow, [single]($p.X - $tamGlow/2 + $tamanhoAtual/2), [single]($p.Y - $tamGlow/2 + $tamanhoAtual/2), $tamGlow, $tamGlow)
        $brushGlow.Dispose()

        # Nucleo da particula
        $corComAlpha = [System.Drawing.Color]::FromArgb($alpha, $corBase.R, $corBase.G, $corBase.B)
        $brush = New-Object System.Drawing.SolidBrush($corComAlpha)
        $e.Graphics.FillEllipse($brush, [single]$p.X, [single]$p.Y, $tamanhoAtual, $tamanhoAtual)
        $brush.Dispose()
    }
})

$script:fogosTickCount = 0
$fireworkTimer = New-Object System.Windows.Forms.Timer
$fireworkTimer.Interval = 30
$fireworkTimer.Add_Tick({
    $script:fogosTickCount++

    # Lanca um novo foguete periodicamente
    if ($script:fogosTickCount % 32 -eq 0) {
        Add-Foguete
    }

    # Atualiza foguetes subindo; ao atingir o alvo, explode
    for ($i = $Foguetes.Count - 1; $i -ge 0; $i--) {
        $f = $Foguetes[$i]
        $f.Y += $f.VY
        if ($f.Y -le $f.YAlvo) {
            Add-ExplosaoFogos -Cx $f.X -Cy $f.Y -Escala 1.0
            $Foguetes.RemoveAt($i)
        }
    }

    # Atualiza particulas existentes (fisica simples: leve gravidade)
    for ($i = $Particulas.Count - 1; $i -ge 0; $i--) {
        $p = $Particulas[$i]
        $p.X += $p.VX
        $p.Y += $p.VY
        $p.VY += 0.06   # gravidade leve
        $p.Vida -= 1
        if ($p.Vida -le 0) { $Particulas.RemoveAt($i) }
    }

    $overlay.Invalidate()
})

# ====================== CONFETES (decoração animada) ======================
# X/Y iniciais + velocidade de queda (VY), leve balanço horizontal (VX) e
# rotação (Rot/RotSpeed) - atualizados a cada tick do timer de animação.
$CoresConfete = @(
    [System.Drawing.Color]::FromArgb(255, 255, 99, 132),   # rosa
    [System.Drawing.Color]::FromArgb(255, 255, 205, 86),   # amarelo
    [System.Drawing.Color]::FromArgb(255, 75, 192, 192),   # ciano
    [System.Drawing.Color]::FromArgb(255, 153, 102, 255),  # roxo
    [System.Drawing.Color]::FromArgb(255, 255, 159, 64),   # laranja
    [System.Drawing.Color]::FromArgb(255, 1, 28, 83)       # azul da marca
)
$Confetes = @(
    @{X=12.0;  Y=8.0;   S=14; Shape='circ'; C=0; VX=0.15;  VY=0.55; Rot=0.0;  RotSpeed=2.5}
    @{X=38.0;  Y=24.0;  S=11; Shape='rect'; C=1; VX=-0.20; VY=0.70; Rot=20.0; RotSpeed=-3.0}
    @{X=70.0;  Y=6.0;   S=13; Shape='tri';  C=2; VX=0.25;  VY=0.45; Rot=0.0;  RotSpeed=2.0}
    @{X=100.0; Y=20.0;  S=11; Shape='circ'; C=3; VX=-0.15; VY=0.60; Rot=0.0;  RotSpeed=-2.0}
    @{X=130.0; Y=4.0;   S=13; Shape='rect'; C=4; VX=0.20;  VY=0.50; Rot=45.0; RotSpeed=3.5}
    @{X=255.0; Y=6.0;   S=13; Shape='rect'; C=0; VX=-0.20; VY=0.65; Rot=10.0; RotSpeed=-2.5}
    @{X=282.0; Y=22.0;  S=11; Shape='circ'; C=1; VX=0.15;  VY=0.50; Rot=0.0;  RotSpeed=2.0}
    @{X=310.0; Y=5.0;   S=14; Shape='tri';  C=5; VX=-0.25; VY=0.55; Rot=0.0;  RotSpeed=-3.0}
    @{X=336.0; Y=20.0;  S=11; Shape='rect'; C=2; VX=0.20;  VY=0.60; Rot=30.0; RotSpeed=3.0}
    @{X=358.0; Y=7.0;   S=13; Shape='circ'; C=4; VX=-0.15; VY=0.45; Rot=0.0;  RotSpeed=-2.5}
    @{X=6.0;   Y=142.0; S=13; Shape='tri';  C=3; VX=0.20;  VY=0.40; Rot=0.0;  RotSpeed=2.5}
    @{X=28.0;  Y=155.0; S=11; Shape='circ'; C=5; VX=-0.15; VY=0.50; Rot=0.0;  RotSpeed=-2.0}
    @{X=54.0;  Y=145.0; S=13; Shape='rect'; C=0; VX=0.25;  VY=0.55; Rot=15.0; RotSpeed=3.0}
    @{X=316.0; Y=148.0; S=13; Shape='circ'; C=2; VX=-0.20; VY=0.60; Rot=0.0;  RotSpeed=-2.5}
    @{X=342.0; Y=140.0; S=11; Shape='tri';  C=1; VX=0.15;  VY=0.45; Rot=0.0;  RotSpeed=2.0}
    @{X=364.0; Y=155.0; S=13; Shape='rect'; C=4; VX=-0.20; VY=0.50; Rot=25.0; RotSpeed=-3.5}
)

$form.Add_Paint({
    param($s, $e)

    $pen = New-Object System.Drawing.Pen($CorBorda, 1)
    $e.Graphics.DrawRectangle($pen, 0, 0, $form.Width - 1, $form.Height - 1)

    try {
        foreach ($cf in $Confetes) {
            $brush = New-Object System.Drawing.SolidBrush($CoresConfete[$cf.C])
            $metade = $cf.S / 2.0
            switch ($cf.Shape) {
                'circ' { $e.Graphics.FillEllipse($brush, [single]$cf.X, [single]$cf.Y, [single]$cf.S, [single]$cf.S) }
                'rect' {
                    $state = $e.Graphics.Save()
                    $e.Graphics.TranslateTransform([single]($cf.X + $metade), [single]($cf.Y + $metade))
                    $e.Graphics.RotateTransform([single]$cf.Rot)
                    $e.Graphics.FillRectangle($brush, [single](-$metade), [single](-$metade), [single]$cf.S, [single]$cf.S)
                    $e.Graphics.Restore($state)
                }
                'tri' {
                    $state = $e.Graphics.Save()
                    $e.Graphics.TranslateTransform([single]($cf.X + $metade), [single]($cf.Y + $metade))
                    $e.Graphics.RotateTransform([single]$cf.Rot)
                    $pts = [System.Drawing.PointF[]]@(
                        (New-Object System.Drawing.PointF([single](-$metade), [single]$metade)),
                        (New-Object System.Drawing.PointF([single]0, [single](-$metade))),
                        (New-Object System.Drawing.PointF([single]$metade, [single]$metade))
                    )
                    $e.Graphics.FillPolygon($brush, $pts)
                    $e.Graphics.Restore($state)
                }
            }
            $brush.Dispose()
        }
    } catch {
        Add-Content -Path "C:\ProgramData\FirstDecision\PopupScripts\Logs\confete_erro.log" -Value "[$(Get-Date)] Erro ao desenhar confetes: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
    }
})

# Timer que anima a queda/rotação dos confetes, reiniciando cada um no topo
# (com X levemente aleatorio) quando sai da area visivel do popup.
$confettiTimer = New-Object System.Windows.Forms.Timer
$confettiTimer.Interval = 40
$confettiTimer.Add_Tick({
    foreach ($cf in $Confetes) {
        $cf.X += $cf.VX
        $cf.Y += $cf.VY
        $cf.Rot += $cf.RotSpeed

        if ($cf.Y -gt $form.Height) {
            $cf.Y = -($cf.S)
            $cf.X = Get-Random -Minimum 0 -Maximum ($form.Width - $cf.S)
        }
        if ($cf.X -lt -20 -or $cf.X -gt ($form.Width + 20)) {
            $cf.X = [double]([int]$cf.X % [int]$form.Width)
        }
    }
    $form.Invalidate()
})

$totalPassos = 45
$passoAtual  = 0
$animTimer = New-Object System.Windows.Forms.Timer
$animTimer.Interval = 14
$animTimer.Add_Tick({
    $script:passoAtual++
    $progresso = $script:passoAtual / $totalPassos
    if ($progresso -gt 1) { $progresso = 1 }
    $suavizado = 1 - [math]::Pow((1 - $progresso), 3)
    $yAtual = $yInicio + (($yFinal - $yInicio) * $suavizado)
    $form.Location = New-Object System.Drawing.Point($xFinal, [int]$yAtual)
    if ($script:passoAtual -ge $totalPassos) {
        $animTimer.Stop()
        $form.Location = New-Object System.Drawing.Point($xFinal, $yFinal)
        $form.Invalidate()
        $form.Refresh()
    }
})
$form.Add_Shown({
    $animTimer.Start()
    $confettiTimer.Start()
    $overlay.Show()
    Add-Foguete
    Add-Foguete
    $fireworkTimer.Start()
    $form.Invalidate()
    $form.Refresh()
})
$form.Add_FormClosed({
    $confettiTimer.Stop()
    $confettiTimer.Dispose()
    $fireworkTimer.Stop()
    $fireworkTimer.Dispose()
    $overlay.Close()
    $overlay.Dispose()
})

$btnFechar = New-Object System.Windows.Forms.Label
$btnFechar.Text = [char]0x2715
$btnFechar.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$btnFechar.ForeColor = $CorX
$btnFechar.BackColor = [System.Drawing.Color]::Transparent
$btnFechar.AutoSize = $true
$btnFechar.Location = New-Object System.Drawing.Point(($form.Width - 30), 12)
$btnFechar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnFechar.Add_Click({ $form.Close() })
$btnFechar.Add_MouseEnter({ $btnFechar.ForeColor = $CorXHover })
$btnFechar.Add_MouseLeave({ $btnFechar.ForeColor = $CorX })

$textoY = 30

$lblLinha1 = New-Object System.Windows.Forms.Label
$lblLinha1.Text = $Linha1
$lblLinha1.Font = New-Object System.Drawing.Font("Segoe UI", 13.5, [System.Drawing.FontStyle]::Bold)
$lblLinha1.ForeColor = $CorTextoTitulo
$lblLinha1.BackColor = [System.Drawing.Color]::Transparent
$lblLinha1.Location = New-Object System.Drawing.Point(10, $textoY)
$lblLinha1.Size = New-Object System.Drawing.Size(360, 30)
$lblLinha1.TextAlign = 'MiddleCenter'

$lblLinha2 = New-Object System.Windows.Forms.Label
$lblLinha2.Text = $Linha2
$lblLinha2.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$lblLinha2.ForeColor = $CorTextoTitulo
$lblLinha2.BackColor = [System.Drawing.Color]::Transparent
$lblLinha2.Location = New-Object System.Drawing.Point(15, ($textoY + 30))
$lblLinha2.Size = New-Object System.Drawing.Size(350, 65)
$lblLinha2.TextAlign = 'MiddleCenter'

# Icone (marca) centralizado no rodape do popup
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Image = $logoImage
$picLogo.SizeMode = 'Zoom'
$IconeLargura = 32
$IconeAltura  = [int]($IconeLargura * $logoImage.Height / $logoImage.Width)
$picLogo.Size = New-Object System.Drawing.Size($IconeLargura, $IconeAltura)
$iconeX = [int](($form.Width - $IconeLargura) / 2)
$iconeY = $form.Height - $IconeAltura - 14
$picLogo.Location = New-Object System.Drawing.Point($iconeX, $iconeY)
$picLogo.BackColor = [System.Drawing.Color]::Transparent

$form.Controls.AddRange(@($picLogo, $btnFechar, $lblLinha1, $lblLinha2))

if ($TempoAutoFecha -gt 0) {
    $autoTimer = New-Object System.Windows.Forms.Timer
    $autoTimer.Interval = $TempoAutoFecha * 1000
    $autoTimer.Add_Tick({
        $autoTimer.Stop()
        $form.Close()
    })
    $autoTimer.Start()
}

[System.Windows.Forms.Application]::Run($form)
