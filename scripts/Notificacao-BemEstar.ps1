<#
.SYNOPSIS
    Popup de pausa de bem-estar, no mesmo padrao visual do FirstNews.

.DESCRIPTION
    - Anima deslizando de baixo para cima ao aparecer
    - Nao tem botao de acao/link - clicar em qualquer lugar (ou esperar o
      timer) fecha o popup, igual ao comportamento do popup.ps1 original
    - Botao "X" no canto tambem fecha
    - Auto-close por timer apos 1 minuto
    - Logo First Decision embutida em Base64 (sem depender de arquivo externo)

.NOTAS
    Rodando via Agendador de Tarefas com script puxado do Git:
      - Acao: powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "caminho\script.ps1"
      - Este arquivo esta salvo em UTF-8 com BOM. Se reeditar, mantenha essa
        codificacao, senao os acentos voltam a bugar.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ====================== CONFIGURACAO ======================
$Titulo         = "Vamos fazer uma pausa de bem-estar?"
$Descricao      = "Levante-se, alongue-se e beba água."
$TempoAutoFecha = 60  # segundos; 0 = não fecha sozinho

$CorFundo         = [System.Drawing.Color]::FromArgb(255, 1, 28, 83)     # mesmo azul do modelo original
$CorTextoTitulo   = [System.Drawing.Color]::White
$CorTextoDesc     = [System.Drawing.Color]::White
$CorX             = [System.Drawing.Color]::FromArgb(255, 190, 200, 220)
$CorXHover        = [System.Drawing.Color]::White

# ====================== LOGO (Base64) ======================
$LogoBase64 = "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"

$logoBytes  = [Convert]::FromBase64String($LogoBase64)
$logoStream = New-Object System.IO.MemoryStream(,$logoBytes)
$logoImage  = [System.Drawing.Image]::FromStream($logoStream)

# ====================== FORM PRINCIPAL ======================
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'None'
$form.StartPosition   = 'Manual'
$form.TopMost          = $true
$form.ShowInTaskbar    = $false
$form.Size             = New-Object System.Drawing.Size(380, 190)
$form.BackColor        = $CorFundo

# posicao final (canto inferior direito, acima da barra de tarefas)
$screen  = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$xFinal  = $screen.Right  - $form.Width  - 16
$yFinal  = $screen.Bottom - $form.Height - 16
$yInicio = $screen.Bottom + 10   # comeca totalmente fora da tela, embaixo

$form.Location = New-Object System.Drawing.Point($xFinal, $yInicio)

# ====================== ANIMACAO DE ENTRADA (slide-in de baixo pra cima) ======================
$totalPassos  = 45
$passoAtual   = 0
$animTimer = New-Object System.Windows.Forms.Timer
$animTimer.Interval = 14

$animTimer.Add_Tick({
    $script:passoAtual++
    $progresso = $script:passoAtual / $totalPassos
    if ($progresso -gt 1) { $progresso = 1 }
    # ease-out cubico: comeca rapido, desacelera perto do destino
    $suavizado = 1 - [math]::Pow((1 - $progresso), 3)
    $yAtual = $yInicio + (($yFinal - $yInicio) * $suavizado)
    $form.Location = New-Object System.Drawing.Point($xFinal, [int]$yAtual)

    if ($script:passoAtual -ge $totalPassos) {
        $animTimer.Stop()
        $form.Location = New-Object System.Drawing.Point($xFinal, $yFinal)
    }
})

$form.Add_Shown({ $animTimer.Start() })

# ====================== LOGO (PictureBox, centralizada) ======================
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Image = $logoImage
$picLogo.SizeMode = 'Zoom'
$picLogo.Size = New-Object System.Drawing.Size($logoImage.Width, $logoImage.Height)
$logoX = [int](($form.Width - $logoImage.Width) / 2)
$picLogo.Location = New-Object System.Drawing.Point($logoX, 22)
$picLogo.BackColor = [System.Drawing.Color]::Transparent

# ====================== BOTAO FECHAR (X) ======================
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

# ====================== TITULO (centralizado) ======================
$lblTitulo = New-Object System.Windows.Forms.Label
$lblTitulo.Text = $Titulo
$lblTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblTitulo.ForeColor = $CorTextoTitulo
$lblTitulo.BackColor = [System.Drawing.Color]::Transparent
$lblTitulo.Location = New-Object System.Drawing.Point(10, 95)
$lblTitulo.Size = New-Object System.Drawing.Size(360, 30)
$lblTitulo.TextAlign = 'MiddleCenter'

# ====================== DESCRICAO (centralizada) ======================
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = $Descricao
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lblDesc.ForeColor = $CorTextoDesc
$lblDesc.BackColor = [System.Drawing.Color]::Transparent
$lblDesc.Location = New-Object System.Drawing.Point(10, 128)
$lblDesc.Size = New-Object System.Drawing.Size(360, 30)
$lblDesc.TextAlign = 'MiddleCenter'

# ====================== MONTAGEM ======================
$form.Controls.AddRange(@($picLogo, $btnFechar, $lblTitulo, $lblDesc))

# ====================== AUTO-FECHAR ======================
if ($TempoAutoFecha -gt 0) {
    $autoTimer = New-Object System.Windows.Forms.Timer
    $autoTimer.Interval = $TempoAutoFecha * 1000
    $autoTimer.Add_Tick({
        $autoTimer.Stop()
        $form.Close()
    })
    $autoTimer.Start()
}

# ====================== EXIBICAO ======================
[System.Windows.Forms.Application]::Run($form)
