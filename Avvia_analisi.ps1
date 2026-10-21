# Usa un Python funzionante senza modificare l'ambiente del logger.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$candidates = @(
    (Join-Path $taskRoot '.venv\Scripts\python.exe'),
    (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe')
)
$analysisPython = $null
foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
        $savedErrorPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $probe = & $candidate --version 2>&1
        $ErrorActionPreference = $savedErrorPreference
        if ($LASTEXITCODE -eq 0) {
            $analysisPython = $candidate
            break
        }
    }
}
if (-not $analysisPython) {
    throw 'Nessun Python funzionante trovato. Esegui analisi_lab6.py con un Python e le librerie di requirements_analisi.txt.'
}
& $analysisPython (Join-Path $taskRoot 'analisi_lab6.py') @args
if ($LASTEXITCODE -ne 0) { throw "Analisi interrotta (codice $LASTEXITCODE)." }
