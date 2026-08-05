<#
.SYNOPSIS
    Script de teste - popup simples para validar o fluxo do launcher central.
#>

Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.MessageBox]::Show(
    "Este popup foi baixado e executado via launcher central!`n`nData/Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
    "Teste - First Decision",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
