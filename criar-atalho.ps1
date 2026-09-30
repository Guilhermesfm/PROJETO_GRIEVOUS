# Cria um atalho do Grievous na Area de Trabalho, com o icone do General.
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
$atalho = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Grievous.lnk'

$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($atalho)
$link.TargetPath = Join-Path $raiz 'Grievous.bat'
$link.WorkingDirectory = $raiz
$link.IconLocation = Join-Path $raiz 'assets\grievous.ico'
$link.Description = 'Automacoes do Field Control'
$link.Save()

Write-Host "Atalho criado em: $atalho"
