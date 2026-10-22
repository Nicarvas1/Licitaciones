<#
.SYNOPSIS
  Corrida completa y desatendida: filtra, descarga, extrae y audita todos los meses del catalogo.

.DESCRIPTION
  1. Verifica el entorno (Tesseract, LM Studio, portal, disco...). Si hay errores, no arranca.
  2. Corre automatizar_pipeline.py con la configuracion recomendada. Mientras extrae un mes descarga el
     siguiente, repara descargas incompletas y vuelve a intentar los proveedores que fallaron.
  3. Si la corrida se corta (LM Studio se cae, se pierde internet...), espera y la retoma donde quedo.
     Todo es reanudable: volver a ejecutar este mismo comando nunca repite lo que ya esta bien.
  4. Al final audita cada mes (sin IA) y deja el Excel consolidado.

.EXAMPLE
  .\correr_fin_de_semana.ps1 -Catalogo .\ListaLicitaciones.csv -Salida D:\lotes
  .\correr_fin_de_semana.ps1 -Catalogo .\ListaLicitaciones.csv -Salida D:\lotes -Desde 2025-11 -Hasta 2025-11
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Catalogo,
    [Parameter(Mandatory = $true)][string]$Salida,
    [string]$Modelo = "qwen/qwen3.6-35b-a3b",
    [string]$Desde = "",
    [string]$Hasta = "",
    [int]$Intentos = 6,
    [int]$EsperaMinutos = 10,
    [switch]$SinAuditoria
)

$ErrorActionPreference = "Continue"
$Raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
$Py = Join-Path $Raiz ".venv\Scripts\python.exe"
if (-not (Test-Path $Py)) { $Py = "python" }
$env:PYTHONIOENCODING = "utf-8"
Set-Location $Raiz
New-Item -ItemType Directory -Force $Salida | Out-Null
$registro = Join-Path $Salida "fin_de_semana.log"

function Escribir($texto) {
    $linea = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $texto
    Write-Host $linea
    Add-Content -Path $registro -Value $linea -Encoding utf8
}

Escribir "Inicio. Catalogo: $($Catalogo -join ', ') | Salida: $Salida | Modelo: $Modelo"

# 1. Verificacion (tambien la hace automatizar_pipeline; aqui se ve antes de empezar)
& $Py verificar_entorno.py --modelo $Modelo --salida $Salida
if ($LASTEXITCODE -ne 0) {
    Escribir "La verificacion encontro errores. Corrigelos y vuelve a ejecutar."
    exit 1
}

# 2. Pipeline con reintentos
$argumentos = @("automatizar_pipeline.py", "--in") + $Catalogo + @("--salida", $Salida, "--recomendado",
    "--modelo", $Modelo, "--max-tokens", "4096", "--timeout", "600", "--reintentos-modelo", "2", "--sin-verificar")
if ($Desde) { $argumentos += @("--desde", $Desde) }
if ($Hasta) { $argumentos += @("--hasta", $Hasta) }

$terminado = $false
for ($intento = 1; $intento -le $Intentos; $intento++) {
    Escribir "Pipeline, intento $intento de $Intentos"
    $lock = Join-Path $Salida ".pipeline.lock"
    if (Test-Path $lock) {
        Escribir "Habia un lock de una corrida anterior interrumpida; se elimina."
        Remove-Item $lock -Force
    }
    & $Py -u @argumentos
    if ($LASTEXITCODE -eq 0) { $terminado = $true; Escribir "Pipeline completado."; break }
    Escribir "El pipeline termino con codigo $LASTEXITCODE. Se reintenta en $EsperaMinutos min (retoma donde quedo)."
    Start-Sleep -Seconds ($EsperaMinutos * 60)
}
if (-not $terminado) { Escribir "El pipeline no termino sin incidencias tras $Intentos intentos. Revisa $Salida\pipeline.log." }

# 3. Auditoria por mes (sin IA)
if (-not $SinAuditoria) {
    Get-ChildItem $Salida -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}$' -and (Test-Path (Join-Path $_.FullName "ofertas")) } |
        Sort-Object Name | ForEach-Object {
            Escribir "Auditoria $($_.Name)"
            & $Py 5_auditar_extraccion.py --dir (Join-Path $_.FullName "ofertas") --excel (Join-Path $_.FullName "auditoria.xlsx") --workers 2 |
                Out-File -FilePath (Join-Path $_.FullName "auditoria.log") -Encoding utf8
        }
}
Escribir "Fin. Excel consolidado: $(Join-Path $Salida 'resultados_consolidados.xlsx')"
