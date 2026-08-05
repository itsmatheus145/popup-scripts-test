# 1. Configuração de DPI e Estilo Visual
try {
   [System.Windows.Forms.Application]::EnableVisualStyles()
   [System.Windows.Forms.Application]::SetHighDpiMode("PerMonitorV2")
} catch {}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# 2. Criar janela
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size(340,150)
$form.StartPosition = "Manual"
$form.FormBorderStyle = "None"
$form.BackColor = [Drawing.Color]::FromArgb(0,28,80)
$form.TopMost = $true
$form.Cursor = "Hand"
# --- AQUI ESTÁ A CORREÇÃO PARA A BARRA DE TAREFAS ---
$form.ShowInTaskbar = $false
# ----------------------------------------------------
$form.Add_Click({ $form.Close() })
$screen = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position)
$workingArea = $screen.WorkingArea
$finalX = $workingArea.Right - $form.Width - 15
$finalY = $workingArea.Bottom - $form.Height - 15
$form.Location = New-Object Drawing.Point($finalX, $workingArea.Bottom)
# --- CONTEÚDO ---
$picture = New-Object Windows.Forms.PictureBox
$picture.Size = New-Object Drawing.Size(110,35)
$picture.Location = New-Object Drawing.Point([int](($form.Width - $picture.Width) / 2), 20)
$picture.SizeMode = "Zoom"
$picture.BackColor = $form.BackColor
$picture.Image = [Drawing.Image]::FromFile("C:\Scripts\logo.png")
$picture.Add_Click({ $form.Close() })
$title = New-Object Windows.Forms.Label
$title.Text = "Vamos fazer uma pausa de bem-estar?"
$title.ForeColor = "White"
$title.Font = New-Object Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$title.Size = New-Object Drawing.Size(320,30)
$title.TextAlign = "MiddleCenter"
$title.Location = New-Object Drawing.Point([int](($form.Width - $title.Width) / 2), 65)
$title.Add_Click({ $form.Close() })
$desc = New-Object Windows.Forms.Label
$desc.Text = "Levante-se, alongue-se e beba " + [char]225 + "gua."
$desc.ForeColor = "White"
$desc.Font = New-Object Drawing.Font("Segoe UI",10)
$desc.Size = New-Object Drawing.Size(320,25)
$desc.TextAlign = "MiddleCenter"
$desc.Location = New-Object Drawing.Point([int](($form.Width - $desc.Width) / 2), 95)
$desc.Add_Click({ $form.Close() })
$closeArea = New-Object Windows.Forms.Panel
$closeArea.Size = New-Object Drawing.Size(28,28)
$closeArea.Location = New-Object Drawing.Point(($form.Width - 35), 8)
$closeBtn = New-Object Windows.Forms.Label
$closeBtn.Text = [char]0x2715
$closeBtn.ForeColor = "White"
$closeBtn.Font = New-Object Drawing.Font("Segoe UI Symbol",9,[System.Drawing.FontStyle]::Bold)
$closeBtn.AutoSize = $true
$closeBtn.Location = New-Object Drawing.Point(6,5)
$closeArea.Add_Click({ $form.Close() })
$closeBtn.Add_Click({ $form.Close() })
$closeArea.Controls.Add($closeBtn)
# --- ANIMAÇÕES ---
$animUp = New-Object System.Windows.Forms.Timer
$animUp.Interval = 1
$animUp.Add_Tick({
  if ($form.Top -gt $finalY) { $form.Top -= 3 }
  else { $animUp.Stop(); $displayTimer.Start() }
})
$displayTimer = New-Object System.Windows.Forms.Timer
$displayTimer.Interval = 10000
$displayTimer.Add_Tick({ $displayTimer.Stop(); $animDown.Start() })
$animDown = New-Object System.Windows.Forms.Timer
$animDown.Interval = 1
$animDown.Add_Tick({
  if ($form.Top -lt $workingArea.Bottom) { $form.Top += 3 }
  else { $animDown.Stop(); $form.Close() }
})
$form.Controls.Add($picture)
$form.Controls.Add($title)
$form.Controls.Add($desc)
$form.Controls.Add($closeArea)
$form.Add_Shown({ $animUp.Start() })
$form.ShowDialog()
