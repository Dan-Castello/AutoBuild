#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_com_stress.ps1
# @Description : Prueba de estres del subsistema COM: abre y cierra instancias
#                Excel y Word en ciclos repetidos, verificando tras cada ciclo
#                que el proceso OS termino realmente (cero zombis).
#                Un fallo aqui indica fuga de ref-count en las librerias.
# @Category    : Diagnostico
# @Version     : 1.0.0
# @Param       : Extra  int optional "Numero de ciclos (defecto 3)"
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Ciclos repetidos Excel+Word verificando cero zombis tras cada cierre
task diag_com_stress {
    $ctx = New-TaskContext `
        -TaskName 'diag_com_stress' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent) `
        -Params   @{ Ciclos = $Extra }

    $ciclos = 3
    if (-not [string]::IsNullOrWhiteSpace($Extra)) {
        $n = 0
        if ([int]::TryParse($Extra, [ref]$n) -and $n -ge 1 -and $n -le 10) {
            $ciclos = $n
        }
    }

    Write-BuildLog $ctx 'INFO' "Iniciando stress COM: $ciclos ciclos"
    Write-Build Cyan "  Ciclos planificados: $ciclos"
    Write-Build Cyan ""

    $errores   = 0
    $zombisTot = 0

    # ---- Limpiar estado previo ---------------------------------------------
    $prev = Remove-ZombieCom
    if ($prev -gt 0) {
        Write-Build Yellow "  [PRE]  Zombis previos eliminados: $prev"
        Start-Sleep -Seconds 2
    }

    # ========================================================================
    # CICLOS EXCEL
    # ========================================================================
    Write-Build Cyan  "  [XL]  --- Ciclos Excel ---"
    for ($i = 1; $i -le $ciclos; $i++) {
        Write-Build Cyan "  [XL]  Ciclo $i/$ciclos ..."
        $xl = $null
        $wb = $null
        $ws = $null
        try {
            $xl = New-ExcelApp -Context $ctx -TimeoutSec 20
            if ($null -eq $xl) { throw 'Excel no disponible' }

            $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
            $ws = Get-ExcelSheet -Workbook $wb -Index 1
            $ws.Name = "Stress$i"

            # Escribir una celda y leerla de vuelta
            $data = @(@{ K = "Ciclo"; V = "$i" })
            Write-ExcelData -Context $ctx -Sheet $ws -Data $data -Columns @('K','V')
            $leido = Get-ExcelCellValue -Sheet $ws -Row 1 -Col 1
            if ($leido -ne 'K') {
                Write-Build Red "  [XL]  Ciclo $i : lectura inesperada '$leido'"
                $errores++
            }

            Invoke-ExcelAutoFit -Sheet $ws
            $outF = Join-Path $ctx.Paths.Output "stress_xl_${i}_$($ctx.RunId).xlsx"
            Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $outF

        } catch {
            Write-Build Red "  [XL]  Ciclo $i FAIL: $_"
            $errores++
        } finally {
            if ($null -ne $ws) { Release-ComObject $ws ; $ws = $null }
            Close-ExcelWorkbook -Workbook $wb -Save $false
            Close-ExcelApp -ExcelApp $xl
            $wb = $null ; $xl = $null
        }

        # Verificacion post-ciclo
        $z = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($z -eq 0) {
            Write-Build Green "  [XL]  Ciclo $i : cerrado OK (0 zombis)"
        } else {
            Write-Build Red   "  [XL]  Ciclo $i : $z ZOMBI(S) tras cierre - FUGA DETECTADA"
            $errores++
            $zombisTot += $z
            # Limpiar para que el siguiente ciclo no acumule
            Remove-ZombieCom | Out-Null
            Start-Sleep -Seconds 2
        }
    }

    Write-Build Cyan ""

    # ========================================================================
    # CICLOS WORD
    # ========================================================================
    Write-Build Cyan  "  [WD]  --- Ciclos Word ---"
    for ($i = 1; $i -le $ciclos; $i++) {
        Write-Build Cyan "  [WD]  Ciclo $i/$ciclos ..."
        $wd  = $null
        $doc = $null
        $sel = $null
        try {
            $wd = New-WordApp -Context $ctx -TimeoutSec 20
            if ($null -eq $wd) { throw 'Word no disponible' }

            $doc = New-WordDocument -Context $ctx -WordApp $wd
            $sel = Get-WordSelection -WordApp $wd
            Add-WordParagraph -Context $ctx -Selection $sel -Text "Stress ciclo $i"
            Release-ComObject $sel ; $sel = $null

            $cc = Get-WordCharCount -Document $doc
            if ($cc -lt 5) {
                Write-Build Red "  [WD]  Ciclo $i : contenido inesperado ($cc chars)"
                $errores++
            }

            $outF = Join-Path $ctx.Paths.Output "stress_wd_${i}_$($ctx.RunId).docx"
            Save-WordDocument -Context $ctx -Document $doc -Path $outF

        } catch {
            Write-Build Red "  [WD]  Ciclo $i FAIL: $_"
            $errores++
        } finally {
            if ($null -ne $sel) { Release-ComObject $sel ; $sel = $null }
            Close-WordDocument -Document $doc -Save $false
            Close-WordApp -WordApp $wd
            $doc = $null ; $wd = $null
        }

        $z = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($z -eq 0) {
            Write-Build Green "  [WD]  Ciclo $i : cerrado OK (0 zombis)"
        } else {
            Write-Build Red   "  [WD]  Ciclo $i : $z ZOMBI(S) - FUGA DETECTADA"
            $errores++
            $zombisTot += $z
            Remove-ZombieCom | Out-Null
            Start-Sleep -Seconds 2
        }
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    Write-Build Cyan "  Ciclos completados : $ciclos Excel + $ciclos Word"
    Write-Build Cyan "  Zombis acumulados  : $zombisTot"

    Write-BuildLog $ctx 'INFO' "diag_com_stress: errores=$errores zombis=$zombisTot"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        Write-Build Red "  RESULTADO: $errores error(es) - hay fugas COM"
        throw "diag_com_stress detecto $errores error(es)"
    } else {
        Write-Build Green "  RESULTADO: OK - todos los procesos cerraron sin zombis"
    }
}
