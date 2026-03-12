#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_excel_rw.ps1
# @Description : Diagnostico de lectura/escritura avanzada en Excel.
#                Usa las funciones Get-ExcelSheet, Add-ExcelSheet,
#                Get-ExcelUsedRange, Get-ExcelCellValue e Invoke-ExcelAutoFit
#                para garantizar liberacion correcta de todos los COM objects.
# @Category    : Utilidad
# @Version     : 1.1.0
# @Param       : Extra  string optional "Nombre del .xlsx en input\ a usar como fuente"
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Lee un .xlsx de input\, transforma datos y escribe hoja de resumen
task diag_excel_rw {
    $ctx = New-TaskContext `
        -TaskName 'diag_excel_rw' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent) `
        -Params   @{ Archivo = $Extra }

    Write-BuildLog $ctx 'INFO' "Iniciando diag_excel_rw | Archivo=$Extra"

    $errores = 0
    $xl      = $null
    $wbSrc   = $null
    $wbOut   = $null
    $wsSrc   = $null
    $wsDatos = $null
    $wsRes   = $null

    # ---- Determinar archivo fuente -----------------------------------------
    $archivoFuente = $null

    if (-not [string]::IsNullOrWhiteSpace($Extra)) {
        $ruta = Join-Path $ctx.Paths.Input $Extra
        if (Test-Path $ruta) {
            $archivoFuente = $ruta
            Write-Build Cyan  "  [SRC]  Usando archivo de input\ : $Extra"
        } else {
            throw "Archivo no encontrado en input\: $Extra"
        }
    } else {
        Write-Build Cyan  "  [SRC]  Sin -Extra: datos sinteticos en memoria"
    }

    try {
        if (-not (Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 15)) {
            throw 'Excel COM no disponible'
        }

        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30
        if ($null -eq $xl) { throw 'New-ExcelApp devolvio null' }
        Write-Build Green "  [COM]  Instancia Excel    : OK"

        # ---- Fase A: cargar o crear datos fuente ---------------------------
        $datos   = @()
        $cols    = @()

        if ($null -ne $archivoFuente) {
            $wbSrc = Open-ExcelWorkbook -Context $ctx -ExcelApp $xl `
                -Path $archivoFuente -ReadOnly $true
            $wsSrc = Get-ExcelSheet -Workbook $wbSrc -Index 1

            $dim = Get-ExcelUsedRange -Sheet $wsSrc
            Write-Build Cyan  "  [READ] Hoja '$($wsSrc.Name)' : $($dim.Rows) filas x $($dim.Cols) columnas"

            if ($dim.Rows -lt 2) {
                Write-Build Yellow "  [READ] Menos de 2 filas - sin datos utiles"
            } else {
                $headers = @()
                for ($c = 1; $c -le $dim.Cols; $c++) {
                    $h = Get-ExcelCellValue -Sheet $wsSrc -Row 1 -Col $c
                    if ([string]::IsNullOrWhiteSpace($h)) { $h = "Col$c" }
                    $headers += $h
                }
                $cols = $headers

                for ($r = 2; $r -le $dim.Rows; $r++) {
                    $fila = @{}
                    for ($c = 1; $c -le $dim.Cols; $c++) {
                        $fila[$headers[$c - 1]] = Get-ExcelCellValue -Sheet $wsSrc -Row $r -Col $c
                    }
                    $datos += $fila
                }
                Write-Build Green "  [READ] Leidas $($datos.Count) filas"
            }

            # Liberar hoja y workbook fuente
            Release-ComObject $wsSrc ; $wsSrc  = $null
            Close-ExcelWorkbook -Workbook $wbSrc -Save $false
            $wbSrc = $null

        } else {
            $datos = @(
                @{ ID = '1'; Articulo = 'Teclado';   Unidades = '10'; PrecioUnit = '45.00'  },
                @{ ID = '2'; Articulo = 'Monitor';   Unidades = '3';  PrecioUnit = '350.00' },
                @{ ID = '3'; Articulo = 'Raton';     Unidades = '15'; PrecioUnit = '18.00'  },
                @{ ID = '4'; Articulo = 'Auricular'; Unidades = '7';  PrecioUnit = '62.00'  }
            )
            $cols = @('ID','Articulo','Unidades','PrecioUnit')
            Write-Build Cyan  "  [SRC]  Datos sinteticos : $($datos.Count) filas"
        }

        # ---- Fase B: workbook de salida ------------------------------------
        $wbOut = New-ExcelWorkbook -Context $ctx -ExcelApp $xl

        # Hoja Datos (indice 1, ya existe)
        $wsDatos = Get-ExcelSheet -Workbook $wbOut -Index 1
        $wsDatos.Name = 'Datos'

        if ($datos.Count -gt 0) {
            Write-ExcelData -Context $ctx -Sheet $wsDatos -Data $datos -Columns $cols
            Invoke-ExcelAutoFit -Sheet $wsDatos
            Write-Build Green "  [WRITE] Hoja Datos : OK ($($datos.Count) filas)"
        }

        Release-ComObject $wsDatos ; $wsDatos = $null

        # Hoja Resumen (nueva)
        $wsRes = Add-ExcelSheet -Workbook $wbOut
        $wsRes.Name = 'Resumen'

        $resumen = @(
            @{ Metrica = 'Tarea';         Valor = 'diag_excel_rw'                           },
            @{ Metrica = 'RunId';         Valor = $ctx.RunId                                 },
            @{ Metrica = 'Fecha';         Valor = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')   },
            @{ Metrica = 'FilasLeidas';   Valor = "$($datos.Count)"                          },
            @{ Metrica = 'ArchivoFuente'; Valor = if ($archivoFuente) {
                [System.IO.Path]::GetFileName($archivoFuente) } else { 'sintetico' }         }
        )
        Write-ExcelData -Context $ctx -Sheet $wsRes -Data $resumen -Columns @('Metrica','Valor')
        Invoke-ExcelAutoFit -Sheet $wsRes
        Write-Build Green "  [WRITE] Hoja Resumen : OK"

        Release-ComObject $wsRes ; $wsRes = $null

        $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outFile = Join-Path $ctx.Paths.Output "diag_excel_rw_${ts}.xlsx"
        Save-ExcelWorkbook -Context $ctx -Workbook $wbOut -Path $outFile

        if (Test-Path $outFile) {
            Write-Build Green "  [SAVE] OK : $((Get-Item $outFile).Length) bytes -> $outFile"
        } else {
            Write-Build Red   "  [SAVE] FAIL : archivo no encontrado"
            $errores++
        }

        Write-RunResult -Context $ctx -Success ($errores -eq 0)

    } catch {
        $errores++
        Write-Build Red "  [ERR]  Excepcion: $_"
        Write-BuildLog $ctx 'ERROR' "Excepcion en diag_excel_rw: $_"
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
        throw

    } finally {
        # Liberar hojas si quedaron abiertas por una excepcion
        if ($null -ne $wsSrc)   { Release-ComObject $wsSrc;   $wsSrc   = $null }
        if ($null -ne $wsDatos) { Release-ComObject $wsDatos; $wsDatos = $null }
        if ($null -ne $wsRes)   { Release-ComObject $wsRes;   $wsRes   = $null }

        # Workbooks -> App
        if ($null -ne $wbSrc) { Close-ExcelWorkbook -Workbook $wbSrc -Save $false ; $wbSrc = $null }
        Close-ExcelWorkbook -Workbook $wbOut -Save $false
        Close-ExcelApp -ExcelApp $xl
        $wbOut = $null
        $xl    = $null

        $zombis = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($zombis -eq 0) {
            Write-Build Green "  [GC]   Sin zombis Excel OK"
        } else {
            Write-Build Yellow "  [GC]   Zombis EXCEL: $zombis -> ejecutar limpiar_com"
        }
    }

    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es)"
        throw "diag_excel_rw detecto $errores error(es)"
    }
    Write-BuildLog $ctx 'INFO' "diag_excel_rw completado. Errores=$errores"
}
