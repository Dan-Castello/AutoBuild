#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_excel.ps1
# @Description : Diagnostico completo de Excel COM: instancia, workbook,
#                escritura, guardado, liberacion y verificacion de zombis
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Prueba ciclo completo Excel COM (instancia, escritura, guardado, liberacion)
task diag_excel {
    $ctx = New-TaskContext `
        -TaskName 'diag_excel' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de Excel COM'

    $errores = 0
    $xl      = $null
    $wb      = $null

    # ---- 0. Limpiar zombis previos -----------------------------------------
    $zombisPrev = Remove-ZombieCom
    if ($zombisPrev -gt 0) {
        Write-Build Yellow "  [PRE]  Zombis previos eliminados: $zombisPrev"
        Start-Sleep -Seconds 2
    }

    # ---- 1. Verificar que Excel esta instalado (sin instanciar) ------------
    $exeExcel = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($exeExcel) {
        Write-Build Green "  [INST] Excel.exe         : $exeExcel"
    } else {
        Write-Build Yellow "  [INST] Excel.exe         : no encontrado en rutas conocidas (puede estar en ruta distinta)"
    }

    # ---- 2. Comprobar servidor COM antes de instanciar ---------------------
    Write-Build Cyan  "  [COM]  Comprobando disponibilidad del servidor COM (hasta 20s)..."
    $comOk = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 20
    if ($comOk) {
        Write-Build Green "  [COM]  Servidor COM       : disponible"
    } else {
        Write-Build Red   "  [COM]  Servidor COM       : NO responde"
        Write-Build Red   "         Abrir Excel manualmente, cerrar dialogos pendientes y reintentar"
        $errores++
        Write-BuildLog $ctx 'ERROR' 'Servidor COM de Excel no disponible'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'COM no disponible'
        throw 'Excel COM no disponible - diagnostico abortado'
    }

    # ---- 3. Ciclo completo COM (instancia -> workbook -> datos -> guardado) -
    try {
        # 3a. Instancia
        Write-Build Cyan  "  [COM]  Instanciando Excel..."
        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30
        if ($null -eq $xl) {
            $errores++
            throw 'New-ExcelApp devolvio null'
        }
        Write-Build Green "  [COM]  Instancia          : OK"

        # 3b. Version
        try {
            $ver = $xl.Version
            Write-Build Green "  [COM]  Version Excel      : $ver"
        } catch {
            Write-Build Yellow "  [COM]  Version Excel      : no disponible"
        }

        # 3c. Crear workbook
        Write-Build Cyan  "  [WB]   Creando workbook..."
        $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        Write-Build Green "  [WB]   Workbook           : OK"

        # 3d. Acceder a hoja
        $ws = Get-ExcelSheet -Workbook $wb -Index 1
        $ws.Name = 'DiagExcel'
        Write-Build Green "  [WS]   Hoja activa        : OK"

        # 3e. Escritura de datos de prueba
        Write-Build Cyan  "  [DATA] Escribiendo datos de prueba..."
        $prueba = @(
            @{ Prueba = 'Texto';    Valor = 'AutoBuild OK'                    },
            @{ Prueba = 'Numero';   Valor = '42'                              },
            @{ Prueba = 'Fecha';    Valor = (Get-Date -Format 'yyyy-MM-dd')   },
            @{ Prueba = 'RunId';    Valor = $ctx.RunId                        }
        )
        Write-ExcelData -Context $ctx -Sheet $ws -Data $prueba -Columns @('Prueba','Valor')

        # Verificar que se escribio - leer y liberar la celda inmediatamente
        $celdaObj = $ws.Cells(1, 1)
        $celda    = $celdaObj.Value2
        Release-ComObject $celdaObj
        $celdaObj = $null
        if ($celda -eq 'Prueba') {
            Write-Build Green "  [DATA] Escritura          : OK (encabezado verificado)"
        } else {
            Write-Build Red   "  [DATA] Escritura          : FAIL (encabezado inesperado: '$celda')"
            $errores++
        }

        # Liberar hoja antes de guardar (ya no se necesita)
        Release-ComObject $ws
        $ws = $null

        # 3f. Guardado
        $outFile = Join-Path $ctx.Paths.Output "diag_excel_$($ctx.RunId).xlsx"
        Write-Build Cyan  "  [SAVE] Guardando en output\..."
        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $outFile

        if (Test-Path $outFile) {
            $size = (Get-Item $outFile).Length
            Write-Build Green "  [SAVE] Guardado           : OK ($size bytes) -> $outFile"
        } else {
            Write-Build Red   "  [SAVE] Guardado           : FAIL (archivo no encontrado tras SaveAs)"
            $errores++
        }

        Write-RunResult -Context $ctx -Success ($errores -eq 0)

    } catch {
        $errores++
        Write-Build Red "  [ERR]  Excepcion: $_"
        Write-BuildLog $ctx 'ERROR' "Excepcion en diag_excel: $_"
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
        throw

    } finally {
        # Liberar intermedios antes del workbook
        if ($null -ne $ws) { Release-ComObject $ws; $ws = $null }

        # Orden correcto: workbook -> app -> GC
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp      -ExcelApp $xl
        $xl = $null
        $wb = $null

        # Verificar cierre real
        $zombis = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($zombis -eq 0) {
            Write-Build Green "  [GC]   Procesos zombi     : ninguno OK"
        } else {
            Write-Build Yellow "  [GC]   Procesos zombi     : $zombis proceso(s) sin ventana detectados"
            Write-Build Yellow "         Ejecutar: .\Run.ps1 limpiar_com"
        }
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - Excel COM funciona correctamente"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en el ciclo COM de Excel"
        throw "diag_excel detecto $errores error(es)"
    }

    Write-BuildLog $ctx 'INFO' "diag_excel completado. Errores=$errores"
}
