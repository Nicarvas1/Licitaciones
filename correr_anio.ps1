<#
.SYNOPSIS
  Corre el pipeline de extraccion para todos los meses de una carpeta de datos, mes por mes y de forma reanudable.

.DESCRIPTION
  Espera la estructura  <Datos>\<YYYY-MM>\{para_scrapear.csv, ofertas\}  o  <Datos>\lotes_*\<YYYY-MM>\...
  Por mes: (1) repara descargas, (2) extrae con LM Studio, (3) audita sin IA.
  Si se interrumpe, volver a ejecutar el mismo comando: los meses terminados (.anio_ok) se saltan y el mes en
  curso retoma donde quedo (la primera vez rehace la extraccion; despues solo continua).

.EXAMPLE
  .\correr_anio.ps1 -Datos D:\licitaciones_datos
  .\correr_anio.ps1 -Datos D:\licitaciones_datos -Desde 2025-11 -Hasta 2025-11 -SinScraping
#>
param(
    [Parameter(Mandatory = $true)][string]$Datos,
    [string]$Modelo = "qwen/qwen3.6-35b-a3b",
    [string]$Desde = "0000-00",
    [string]$Hasta = "9999-99",
    [int]$Paralelo = 4,
    [switch]$SinScraping,
    [switch]$SinAuditoria,
    [switch]$Revision
)

$ErrorActionPreference = "Stop"
$Raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
$Py = Join-Path $Raiz ".venv\Scripts\python.exe"
if (-not (Test-Path $Py)) { $Py = "python" }
$env:PYTHONIOENCODING = "utf-8"
Set-Location $Raiz

# --- Comprobaciones previas -------------------------------------------------------------------------
try {
    $modelos = (Invoke-RestMethod -Uri "http://127.0.0.1:1234/v1/models" -TimeoutSec 10).data.id
} catch {
    throw "LM Studio no responde en http://127.0.0.1:1234. Enciendelo (y LM Link si corresponde) y reintenta."
}
if ($modelos -notcontains $Modelo) {
    throw "El modelo '$Modelo' no esta cargado en LM Studio. Modelos disponibles: $($modelos -join ', ')"
}

$meses = Get-ChildItem -Path $Datos -Directory -Recurse -Depth 2 |
    Where-Object { $_.Name -match '^\d{4}-\d{2}$' -and (Test-Path (Join-Path $_.FullName "ofertas")) } |
    Where-Object { $_.Name -ge $Desde -and $_.Name -le $Hasta } |
    Sort-Object Name
if (-not $meses) { throw "No encontre meses (carpetas YYYY-MM con 'ofertas') en $Datos" }

$logs = Join-Path $Datos "logs"
New-Item -ItemType Directory -Force $logs | Out-Null
"Meses a procesar: $($meses.Name -join ', ')"

function Ejecutar($etapa, $mes, $argumentos) {
    $log = Join-Path $logs ("{0}_{1}.log" -f $mes, $etapa)
    "[{0}] {1} {2} -> {3}" -f (Get-Date -Format "HH:mm:ss"), $mes, $etapa, $log
    & $Py -u @argumentos 2>&1 | Out-File -FilePath $log -Encoding utf8 -Append
    if ($LASTEXITCODE -ne 0) { throw "Fallo la etapa '$etapa' de $mes (codigo $LASTEXITCODE). Ver $log" }
}

foreach ($m in $meses) {
    $mes = $m.Name
    $dir = $m.FullName
    $ok = Join-Path $dir ".anio_ok"
    $iniciado = Join-Path $dir ".anio_iniciado"
    if (Test-Path $ok) { "[$mes] ya terminado, se salta"; continue }

    if (-not $SinScraping) {
        Ejecutar "scraping" $mes @("2_scraper_ofertas.py", "--csv", "$dir\para_scrapear.csv",
            "--salida", "$dir\ofertas", "--reparar", "--pausa-licitaciones", "2")
    }

    $extraccion = @("3_extraer_ia.py", "--dir", "$dir\ofertas", "--backend", "lmstudio", "--modelo", $Modelo,
        "--metadata-csv", "$dir\para_scrapear.csv", "--excel", "$dir\resultado_$mes.xlsx",
        "--por-proveedor", "--max-chars-proveedor", "24000", "--paralelo", "$Paralelo",
        "--max-tokens", "4096", "--timeout", "600", "--reintentos-modelo", "2", "--ocr")
    if (-not (Test-Path $iniciado)) {
        $extraccion += "--rehacer"      # primera vez: se rehace con el codigo y el modelo actuales
        Set-Content -Path $iniciado -Value (Get-Date -Format "s")
    }
    Ejecutar "extraccion" $mes $extraccion

    if (-not $SinAuditoria) {
        Ejecutar "auditoria" $mes @("5_auditar_extraccion.py", "--dir", "$dir\ofertas",
            "--excel", "$dir\auditoria_$mes.xlsx", "--workers", "2")
    }
    if ($Revision) {
        Ejecutar "revision" $mes @("4_revisar_extraccion.py", "--dir", "$dir\ofertas", "--fuente", "ia",
            "--salida", "$dir\revision_ia")
    }
    Set-Content -Path $ok -Value (Get-Date -Format "s")
    "[$mes] terminado"
}

"Uniendo los resultados mensuales..."
& $Py unir_resultados.py --dir $Datos --excel (Join-Path $Datos "resultado_anio.xlsx")
"Listo. Excel del periodo: $(Join-Path $Datos 'resultado_anio.xlsx')"
