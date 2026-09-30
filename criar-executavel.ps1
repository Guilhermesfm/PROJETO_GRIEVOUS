# Compila Grievous.exe DENTRO desta pasta, com o icone do General.
# O executavel nao e copiado para lugar nenhum: mova-o voce mesmo para a
# Area de Trabalho, ou crie um atalho para ele.
#
# Uso:  powershell -ExecutionPolicy Bypass -File criar-executavel.ps1

$ErrorActionPreference = 'Stop'
$raiz  = Split-Path -Parent $MyInvocation.MyCommand.Path
$icone = Join-Path $raiz 'assets\grievous.ico'
$saida = Join-Path $raiz 'Grievous.exe'

if (-not (Test-Path $icone)) { throw "Icone nao encontrado: $icone" }

# Compilador do .NET Framework, presente em qualquer Windows moderno.
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) {
  $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path $csc)) {
  throw "Compilador C# nao encontrado. Instale o .NET Framework 4."
}

# O executavel so localiza a pasta do projeto e chama "node iniciar.js",
# deixando a janela aberta para os erros aparecerem.
$fonte = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

class Grievous
{
    static int Main(string[] args)
    {
        string pasta = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string script = Path.Combine(pasta, "iniciar.js");

        // O .exe pode ter sido movido (a Area de Trabalho, por exemplo).
        // Nesse caso o caminho do projeto fica gravado em Grievous.txt.
        if (!File.Exists(script))
        {
            string ponteiro = Path.Combine(pasta, "Grievous.txt");
            if (File.Exists(ponteiro))
            {
                pasta = File.ReadAllText(ponteiro).Trim();
                script = Path.Combine(pasta, "iniciar.js");
            }
        }

        if (!File.Exists(script))
        {
            Console.Error.WriteLine("Nao encontrei iniciar.js.");
            Console.Error.WriteLine("Mantenha o Grievous.exe na pasta do projeto,");
            Console.Error.WriteLine("ou crie um atalho para ele em vez de mover o arquivo.");
            Console.Error.WriteLine();
            Console.Error.WriteLine("Pressione qualquer tecla para fechar.");
            Console.ReadKey();
            return 1;
        }

        ProcessStartInfo info = new ProcessStartInfo("node", "\"" + script + "\"");
        info.WorkingDirectory = pasta;
        info.UseShellExecute = false;

        try
        {
            Process p = Process.Start(info);
            p.WaitForExit();
            return p.ExitCode;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Nao consegui executar o Node.js: " + e.Message);
            Console.Error.WriteLine("Instale em https://nodejs.org e tente de novo.");
            Console.Error.WriteLine();
            Console.Error.WriteLine("Pressione qualquer tecla para fechar.");
            Console.ReadKey();
            return 1;
        }
    }
}
'@

$arquivoFonte = Join-Path $env:TEMP 'GrievousLauncher.cs'
Set-Content -Path $arquivoFonte -Value $fonte -Encoding UTF8

& $csc /nologo /target:exe /platform:anycpu /win32icon:"$icone" /out:"$saida" "$arquivoFonte"
if ($LASTEXITCODE -ne 0) { throw "Falha ao compilar." }

Remove-Item $arquivoFonte -ErrorAction SilentlyContinue

# Guarda o caminho do projeto, para o .exe continuar funcionando se for movido.
Set-Content -Path (Join-Path $raiz 'Grievous.txt') -Value $raiz -Encoding UTF8

Write-Host ""
Write-Host "Grievous.exe criado em: $saida"
Write-Host ""
Write-Host "Para usar na Area de Trabalho, o recomendado e criar um atalho:"
Write-Host "  clique com o botao direito no Grievous.exe > Enviar para > Area de trabalho"
Write-Host ""
Write-Host "Se preferir mover o proprio .exe, leve o Grievous.txt junto."
