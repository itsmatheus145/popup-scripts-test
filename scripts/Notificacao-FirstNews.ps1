<#
.SYNOPSIS
    Popup de notificacao clicavel, estilo toast, para divulgar noticias do FirstNews.

.DESCRIPTION
    - Anima deslizando de baixo para cima ao aparecer
    - Todo o corpo do popup e clicavel (abre o link do FirstNews)
    - Botao dedicado "Clique aqui para conferir" com cantos arredondados
      (desenhado manualmente com anti-aliasing, para bordas suaves)
    - Botao "X" no canto: fecha o popup SEM abrir o link
    - Auto-close por timer (fallback)
    - Logo First Decision embutida em Base64 (sem depender de arquivo externo)

.NOTAS
    Rodando via Agendador de Tarefas com script puxado do Git:
      - Acao: powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "caminho\script.ps1"
      - Configure "Executar somente quando o usuario estiver conectado"
        (senao o WinForms nao tem sessao interativa pra desenhar a janela)
      - Este arquivo esta salvo em UTF-8 com BOM. Se reeditar, mantenha essa
        codificacao (no VS Code: "Save with Encoding" -> UTF-8 with BOM),
        senao os acentos voltam a bugar.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ====================== CONFIGURACAO ======================
$Titulo         = "A nova edição do FirstNews já está no ar!"
$TextoBotao     = "Clique aqui para conferir"
$UrlAoClicar    = "https://youtube.com"
$TempoAutoFecha = 20  # segundos; 0 = não fecha sozinho
$RaioBotao      = 18  # raio do arredondamento do botão, em pixels

$CorFundo         = [System.Drawing.Color]::FromArgb(255, 1, 28, 83)     # azul do modelo original
$CorTextoTitulo   = [System.Drawing.Color]::White
$CorBotaoFundo    = [System.Drawing.Color]::White
$CorBotaoTexto    = [System.Drawing.Color]::FromArgb(255, 1, 28, 83)
$CorBotaoHover    = [System.Drawing.Color]::FromArgb(255, 225, 230, 240)
$CorX             = [System.Drawing.Color]::FromArgb(255, 190, 200, 220)
$CorXHover        = [System.Drawing.Color]::White

# ====================== LOGO (Base64) ======================
$LogoBase64 = "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"

$logoBytes  = [Convert]::FromBase64String($LogoBase64)
$logoStream = New-Object System.IO.MemoryStream(,$logoBytes)
$logoImage  = [System.Drawing.Image]::FromStream($logoStream)

# ====================== FUNCAO: RETANGULO COM CANTOS ARREDONDADOS ======================
function Get-RoundedRectPath {
    param(
        [System.Drawing.Rectangle]$Rect,
        [int]$Radius
    )
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    $path.AddArc($Rect.X, $Rect.Y, $d, $d, 180, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Y, $d, $d, 270, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

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

# borda sutil
$form.Paint.Add({
    param($s, $e)
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 1)
    $e.Graphics.DrawRectangle($pen, 0, 0, $form.Width - 1, $form.Height - 1)
})

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

# ====================== ACAO DE CLIQUE (abre o link) ======================
$AcaoClicar = {
    try {
        Start-Process $UrlAoClicar
    } catch {
        # log opcional
    }
    $form.Close()
}

# ====================== LOGO (PictureBox, centralizada) ======================
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Image = $logoImage
$picLogo.SizeMode = 'Zoom'
$picLogo.Size = New-Object System.Drawing.Size($logoImage.Width, $logoImage.Height)
$logoX = [int](($form.Width - $logoImage.Width) / 2)
$picLogo.Location = New-Object System.Drawing.Point($logoX, 20)
$picLogo.BackColor = [System.Drawing.Color]::Transparent

# ====================== BOTAO FECHAR (X) -- fecha SEM abrir o link ======================
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
$lblTitulo.Location = New-Object System.Drawing.Point(10, 80)
$lblTitulo.Size = New-Object System.Drawing.Size(360, 44)
$lblTitulo.TextAlign = 'MiddleCenter'

# ====================== BOTAO "CLIQUE AQUI PARA CONFERIR" ======================
$btnAcao = New-Object System.Windows.Forms.Label
$btnAcao.Text = $TextoBotao
$btnAcao.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnAcao.ForeColor = $CorBotaoTexto
$btnAcao.BackColor = $CorFundo
$btnAcao.Size = New-Object System.Drawing.Size(230, 36)
$btnAcaoX = [int](($form.Width - $btnAcao.Width) / 2)
$btnAcao.Location = New-Object System.Drawing.Point($btnAcaoX, 132)
$btnAcao.Cursor = [System.Windows.Forms.Cursors]::Hand

$script:corBotaoAtual = $CorBotaoFundo

$btnAcao.Add_Paint({
    param($sender, $e)
    $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $e.Graphics.Clear($CorFundo)

    $rect = New-Object System.Drawing.Rectangle(0, 0, ($sender.Width - 1), ($sender.Height - 1))
    $path = Get-RoundedRectPath -Rect $rect -Radius $RaioBotao
    $brush = New-Object System.Drawing.SolidBrush($script:corBotaoAtual)
    $e.Graphics.FillPath($brush, $path)
    $brush.Dispose()

    $rectF = New-Object System.Drawing.RectangleF(0, 0, $sender.Width, $sender.Height)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $textBrush = New-Object System.Drawing.SolidBrush($sender.ForeColor)
    $e.Graphics.DrawString($sender.Text, $sender.Font, $textBrush, $rectF, $sf)
    $textBrush.Dispose()
})

$btnAcao.Add_MouseEnter({
    $script:corBotaoAtual = $CorBotaoHover
    $btnAcao.Invalidate()
})
$btnAcao.Add_MouseLeave({
    $script:corBotaoAtual = $CorBotaoFundo
    $btnAcao.Invalidate()
})
$btnAcao.Add_Click($AcaoClicar)

# ====================== MONTAGEM ======================
$form.Controls.AddRange(@($picLogo, $btnFechar, $lblTitulo, $btnAcao))

# apenas o botao branco ($btnAcao) abre o link -- clicar em qualquer outra
# parte do popup (fundo, logo, titulo) nao faz nada. Para fechar, so o X.
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
