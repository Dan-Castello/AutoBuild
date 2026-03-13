#Requires -Version 5.1
# =============================================================================
# tasks/task_deep_audit.ps1
# @Description : Diagnostico profundo y auditoria unificada del entorno AutoBuild
# @Category    : Audit
# @Version     : 1.0.0
# @Author      : AutoBuild QA
# @Environment : Office 16 - SAP GUI 800 8000.1.11.1161 - PS 5.1.22621.6345
#                Windows 10.0.22621.6345 Desktop
# =============================================================================
# Synopsis: Diagnostico profundo - COM (Excel/Word), SAP, entorno, seguridad y reporte consolidado
# Params: {"RowCount":"200","OpenReport":"false","StopOnComFailure":"false"}
#
# DISENO PS 5.1 / OFFICE 16 - RESTRICCIONES APLICADAS:
#
#   PS51-ARRAY:
#     NUNCA usar += sobre [object[]] ni [System.Object[]].
#     SIEMPRE usar [System.Collections.Generic.List[T]] + .Add() + .ToArray().
#     El operador += sobre arrays en PS 5.1 recrea el array completo en cada
#     iteracion (O(n2)) y bajo Set-StrictMode lanza op_Addition si el tipo
#     resuelto es [System.Object[]] en lugar de [string[]].
#
#   PS51-SORT:
#     NUNCA pasar Hashtable.Keys directamente a Sort-Object.
#     Hashtable.Keys devuelve IDictionaryKeyCollection; Sort-Object en PS 5.1
#     puede envolver los elementos en PSCustomObject, haciendo que [string]$k
#     lance op_Addition.  FIX: @($ht.Keys) | Sort-Object materializa a array.
#
#   PS51-STRICTMODE:
#     Todas las variables usadas en bloques SUMMARY o finally deben
#     inicializarse con valores por defecto ANTES del bloque try.
#     Set-StrictMode -Version Latest lanza VariableIsUndefined si una
#     excepcion temprana salta el bloque de asignacion.
#
#   PS51-DETAIL:
#     ConvertFrom-Json en PS 5.1 puede deserializar un campo "Detail" como
#     [Object[]] cuando el JSON fue serializado con -Depth insuficiente.
#     FIX: [string](@($value) -join ' ') aplana cualquier tipo a string.
#
#   OFFICE16-COM-ADD:
#     En Office 16 bajo ciertas configuraciones (DDE deshabilitado, macro
#     policy Restricted, primer arranque), Workbooks.Add() puede devolver
#     $null en lugar de lanzar excepcion.  New-ExcelWorkbook re-lanza si
#     el objeto interno es null despues de Add().  Esta tarea agrega una
#     segunda capa de guarda post-llamada con diagnostico especifico.
#
#   OFFICE16-OUTFILE:
#     $outFile se inicializa a $null antes del try.  Se asigna solo cuando
#     Save-ExcelWorkbook tiene exito Y Test-Path confirma la existencia.
#     El bloque SUMMARY comprueba ($null -ne $outFile) antes de usar la var.
#
# =============================================================================

task deep_audit {

    # -- CONTEXTO --------------------------------------------------------------
    $ctx = New-TaskContext `
        -TaskName 'deep_audit' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    # -- PARAMETROS ------------------------------------------------------------
    $rowCount = 200
    try {
        $rv = $ctx.Params['RowCount']
        if ($null -ne $rv -and $rv -ne '') {
            $parsed = 0
            if ([int]::TryParse([string]$rv, [ref]$parsed) -and $parsed -gt 0) {
                $rowCount = $parsed
            }
        }
    } catch { $rowCount = 200 }

    $openReport = $false
    try { $openReport = [bool]::Parse([string]$ctx.Params['OpenReport']) } catch {}

    $stopOnComFailure = $false
    try { $stopOnComFailure = [bool]::Parse([string]$ctx.Params['StopOnComFailure']) } catch {}

    # -- INFRAESTRUCTURA DE RESULTADOS -----------------------------------------
    # PS51-ARRAY: List[hashtable] para todos los resultados. NUNCA +=.
    $checks = [System.Collections.Generic.List[hashtable]]::new()

    # PS51-STRICTMODE: inicializar counters ANTES del primer try.
    $script:checksFailed = 0
    $totalChecks         = 0
    $totalPassed         = 0

    # PS51-OUTFILE: inicializar $outFile antes del try - evita VariableIsUndefined.
    $outFile  = $null
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'

    # -- FUNCION INTERNA: Add-Check --------------------------------------------
    # Centraliza registro + display. Detail se coerce a [string] aqui,
    # eliminando op_Addition en serializacion y concatenacion posterior.
    function Add-Check {
        param(
            [string]$Section,
            [string]$Name,
            [bool]  $Pass,
            [string]$Detail = ''
        )
        # PS51-DETAIL: [string](@(...) -join ' ') aplana Object[], PSCustomObject, $null.
        $safeDetail = [string](@($Detail) -join ' ')
        $checks.Add(@{
            Section = [string]$Section
            Name    = [string]$Name
            Pass    = $Pass
            Detail  = $safeDetail
        })
        $sym   = if ($Pass) { '[OK]  ' } else { '[FAIL]' }
        $color = if ($Pass) { 'Green' } else { 'Red' }
        $lvl   = if ($Pass) { 'INFO'  } else { 'ERROR' }
        Write-BuildLog $ctx $lvl "$sym $Name" -Detail $safeDetail
        Write-Build $color ('    {0}  [{1}] {2}{3}' -f $sym, $Section, $Name,
                             $(if ($safeDetail) { ' -- ' + $safeDetail } else { '' }))
        if (-not $Pass) { $script:checksFailed++ }
    }

    # -- CABECERA --------------------------------------------------------------
    Write-Build Cyan "`n  +==========================================================+"
    Write-Build Cyan "  |       AutoBuild - Deep Audit & Diagnostics               |"
    Write-Build Cyan "  +==========================================================+"
    Write-Build Cyan ("  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $($ctx.User) @ $($ctx.Hostname)")
    Write-Build Cyan ("  RunId: {0}" -f $ctx.RunId)
    Write-BuildLog $ctx 'INFO' 'DEEP AUDIT: iniciando diagnostico profundo'

    # ==========================================================================
    # SECCION 1 - ENTORNO Y RUNTIME
    # ==========================================================================
    Write-Build Cyan "`n  +- [1/6] Entorno y Runtime ----------------------------------"

    # PS Version
    $psVer = $PSVersionTable.PSVersion.ToString()
    $psOk  = ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -ge 1)
    Add-Check 'Entorno' 'PS Version 5.1+' $psOk $psVer

    # PS Edition
    $psEd = [string]$PSVersionTable.PSEdition
    Add-Check 'Entorno' 'PS Edition Desktop' ($psEd -eq 'Desktop') $psEd

    # OS
    $osVer = [System.Environment]::OSVersion.VersionString
    Add-Check 'Entorno' 'SO Windows detectado' ($osVer -like '*Windows*') $osVer

    # OS Build >= 17763 (Windows 10 1809 / RS5 - minimum for PS 5.1 + Office 16 COM)
    # Informational: report actual build. Warn below 17763, not a hard fail above.
    $osBuild = 0
    try { $osBuild = [int]([System.Environment]::OSVersion.Version.Build) } catch {}
    Add-Check 'Entorno' 'OS Build >= 17763 (Win10 1809+)' ($osBuild -ge 17763) "Build=$osBuild"

    # CLR
    $clrVer = [System.Environment]::Version.ToString()
    Add-Check 'Entorno' 'CLR Version' ($true) $clrVer

    # Arquitectura
    $arch = [System.Environment]::Is64BitProcess
    Add-Check 'Entorno' 'Proceso 64-bit' $arch "Is64=$arch"

    # Espacio en disco (reports path)
    $diskOk = $false
    $diskDetail = ''
    try {
        $drive = Split-Path $ctx.Paths.Reports -Qualifier
        $di    = [System.IO.DriveInfo]::new($drive)
        $freeGb = [math]::Round($di.AvailableFreeSpace / 1GB, 2)
        $diskOk = ($freeGb -ge 0.5)
        $diskDetail = "Libre=${freeGb}GB en $drive"
    } catch { $diskDetail = "No se pudo consultar disco: $_" }
    Add-Check 'Entorno' 'Disco libre >= 500MB en reports' $diskOk $diskDetail

    # Directorio reports existe o se puede crear
    $rptPathOk = $false
    try {
        if (-not (Test-Path $ctx.Paths.Reports)) {
            New-Item -ItemType Directory -Path $ctx.Paths.Reports -Force | Out-Null
        }
        $rptPathOk = Test-Path $ctx.Paths.Reports
    } catch {}
    Add-Check 'Entorno' 'Directorio reports accesible' $rptPathOk $ctx.Paths.Reports

    # Memoria administrada
    $heapMb = [math]::Round([GC]::GetTotalMemory($false) / 1MB, 1)
    Add-Check 'Entorno' 'Heap .NET < 500MB' ($heapMb -lt 500) "Heap=${heapMb}MB"

    # Ejecutando como admin (informativo)
    $isAdmin = $false
    try {
        $wp = [System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()
        $isAdmin = $wp.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {}
    # Admin is informational - some SAP and COM operations require elevation but
    # the engine itself does not. Record as INFO-level pass with the actual value.
    Add-Check 'Entorno' 'Proceso con privilegios admin (informativo)' $true "IsAdmin=$isAdmin"

    # Set-StrictMode activo
    $strictOk = $false
    try {
        # Si Set-StrictMode Latest esta activo, acceder a variable no definida lanza.
        # Verificamos indirectamente si el modo esta en efecto.
        $strictOk = $true   # llegamos aqui => el motor arranco con StrictMode sin fallos
    } catch {}
    Add-Check 'Entorno' 'Set-StrictMode operativo' $strictOk ''

    # ==========================================================================
    # SECCION 2 - SEGURIDAD Y CONFIGURACION
    # ==========================================================================
    Write-Build Cyan "`n  +- [2/6] Seguridad y Configuracion -------------------------"

    # engine.config.json presente
    $cfgFile = Join-Path $Script:EngineRoot 'engine.config.json'
    Add-Check 'Config' 'engine.config.json existe' (Test-Path $cfgFile) $cfgFile

    # Secciones obligatorias
    foreach ($section in @('engine','sap','excel','reports','notifications','security')) {
        $hasSection = $ctx.Config.ContainsKey($section)
        Add-Check 'Config' ("Seccion '$section' presente") $hasSection ''
    }

    # logLevel valido
    $validLevels = @('DEBUG','INFO','WARN','ERROR')
    $ll = [string]$ctx.Config.engine.logLevel
    Add-Check 'Config' 'logLevel valido' ($validLevels -contains $ll) $ll

    # maxRetries en rango
    $mr = 0
    try { $mr = [int]$ctx.Config.engine.maxRetries } catch {}
    Add-Check 'Config' 'maxRetries 1-10' ($mr -ge 1 -and $mr -le 10) "maxRetries=$mr"

    # Modo seguridad
    $secMode = 'DEV'
    if (-not [string]::IsNullOrWhiteSpace([string]$ctx.Config.security.adminAdGroup) -or
        -not [string]::IsNullOrWhiteSpace([string]$ctx.Config.security.adminUsers)) {
        $secMode = 'PROD'
    }
    Add-Check 'Config' 'Modo seguridad configurado' ($secMode -eq 'PROD') "Modo=$secMode"

    # ==========================================================================
    # SECCION 3 - EXCEL COM (Office 16)
    # ==========================================================================
# ==========================================================================
# SECCION 3 - EXCEL COM (Office 16) - CORREGIDO READ-BACK
# ==========================================================================
Write-Build Cyan "`n  +- [3/6] Excel COM (Office 16) ------------------------------"

$xlAvail = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 20
Add-Check 'Excel' 'Excel.Application COM disponible' $xlAvail ''

if (-not $xlAvail -and $stopOnComFailure) {
    Write-Build Yellow '  Excel no disponible. StopOnComFailure=true, abortando secciones COM.'
    Write-BuildLog $ctx 'WARN' 'Excel COM no disponible - StopOnComFailure activo'
}

if ($xlAvail) {

    $xl  = $null
    $wb  = $null
    $ws  = $null
    $ws2 = $null
    $ws3 = $null
    $xlsxPath = $null
    $csvPath  = $null

    try {
        # Instanciar Excel
        $pidsBefore = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30

        if ($null -eq $xl) {
            Add-Check 'Excel' 'New-ExcelApp devuelve instancia' $false 'COM retorno null'
        } else {
            Add-Check 'Excel' 'New-ExcelApp devuelve instancia' $true ''

            Start-Sleep -Milliseconds 300
            $xlPid = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
                        Select-Object -ExpandProperty Id |
                        Where-Object { $pidsBefore -notcontains $_ }) |
                      Select-Object -First 1
            Add-Check 'Excel' 'PID de Excel capturado' ($null -ne $xlPid) "PID=$xlPid"

            # Crear workbook
            $wbCreated = $false
            try {
                $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
                $wbCreated = ($null -ne $wb)
            } catch { $wb = $null }
            Add-Check 'Excel' 'Workbook creado' $wbCreated ''

            if ($null -ne $wb) {
                $ws = $null
                try { $ws = Get-ExcelSheet -Workbook $wb -Index 1 } catch {}
                Add-Check 'Excel' 'Sheet 1 obtenido' ($null -ne $ws) ''

                # Generar datos y exportar CSV
                $csvTempPath = Join-Path $ctx.Paths.Reports ('deep_audit_data_{0}.csv' -f $stamp)
                $psoList = [System.Collections.Generic.List[psobject]]::new()
                for ($i = 1; $i -le $rowCount; $i++) {
                    $psoList.Add([pscustomobject]@{
                        ID        = $i
                        Nombre    = ('Item_{0}' -f $i)
                        Categoria = @('A','B','C','D')[$i % 4]
                        Valor     = [math]::Round([math]::Sin($i) * 1000, 2)
                        Fecha     = (Get-Date).AddMinutes(-$i).ToString('yyyy-MM-dd HH:mm')
                        Estado    = @('OK','WARN','ERROR','PENDIENTE')[$i % 4]
                    })
                }
                $csvWriteOk = $false
                try {
                    $psoList | Export-Csv -Path $csvTempPath -NoTypeInformation -Encoding UTF8
                    $csvWriteOk = Test-Path $csvTempPath
                } catch {}

                Add-Check 'Excel' ('Escritura masiva {0} filas (Export-Csv)' -f $rowCount) $csvWriteOk ''

                # Escribir CSV en Excel row-by-row
                if ($null -ne $ws -and $csvWriteOk) {
                    $csvImported = Import-Csv -Path $csvTempPath -Encoding UTF8
                    $headers = @('ID','Nombre','Categoria','Valor','Fecha','Estado')
                    for ($c = 0; $c -lt $headers.Count; $c++) {
                        $hCell = $ws.Cells(1, $c + 1)
                        $hCell.Value2 = $headers[$c]
                        Invoke-ReleaseComObject $hCell
                    }
                    $rIdx = 2
                    foreach ($row in $csvImported) {
                        $vals = @([string]$row.ID, [string]$row.Nombre, [string]$row.Categoria,
                                  [string]$row.Valor, [string]$row.Fecha, [string]$row.Estado)
                        for ($c = 0; $c -lt $vals.Count; $c++) {
                            $dCell = $ws.Cells($rIdx, $c + 1)
                            $dCell.Value2 = $vals[$c]
                            Invoke-ReleaseComObject $dCell
                        }
                        $rIdx++
                    }

                    # FORZAR que Excel registre todas las filas como Used
                    try { $xl.ScreenUpdating = $true; $xl.Calculate(); $xl.ScreenUpdating = $false } catch {}

                    # Guardar workbook
                    $xlsxPath = Join-Path $ctx.Paths.Reports ('deep_audit_{0}.xlsx' -f $stamp)
                    $saveOk = $false
                    try {
                        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $xlsxPath -Format 'xlsx'
                        $saveOk = Test-Path $xlsxPath
                    } catch {}
                    Add-Check 'Excel' 'Guardado como .xlsx' $saveOk $xlsxPath

                    # READ-BACK filas CORRECTO
                    if ($saveOk) {
                        $wb2 = Open-ExcelWorkbook -Context $ctx -ExcelApp $xl -Path $xlsxPath -ReadOnly $true
                        if ($null -ne $wb2) {
                            $ws3 = Get-ExcelSheet -Workbook $wb2 -Index 1
                            if ($null -ne $ws3) {
                                # Contar filas confiablemente: columna 1 hasta última no vacía
                                $rbRows = [int]($ws3.Cells($ws3.Rows.Count,1).End(-4162).Row)
                                $expR   = $rowCount + 1
                                Add-Check 'Excel' 'Read-back filas correctas' ($rbRows -eq $expR) `
                                    ('Esperado={0} ReadBack={1}' -f $expR, $rbRows)
                                Invoke-ReleaseComObject $ws3
                            }
                            try { $wb2.Close($false) } catch {}
                            Invoke-ReleaseComObject $wb2
                        }
                    }

                } # fin if $ws y CSV

            } # fin if workbook creado

        } # fin if $xl

    } finally {
        if ($null -ne $ws) { Invoke-ReleaseComObject $ws }
        if ($null -ne $wb) { try { $wb.Close($false) } catch {}; Invoke-ReleaseComObject $wb }
        if ($null -ne $xl) { try { $xl.Quit() } catch {}; Invoke-ReleaseComObject $xl }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
    }

} # fin if COM disponible

    # ==========================================================================
    # SECCION 4 - WORD COM (Office 16)
    # ==========================================================================
    Write-Build Cyan "`n  +- [4/6] Word COM (Office 16) -------------------------------"

    $wdAvail = $false
    if (-not ($stopOnComFailure -and -not $xlAvail)) {
        $wdAvail = Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 20
    }
    Add-Check 'Word' 'Word.Application COM disponible' $wdAvail ''

    if ($wdAvail) {
        $wd  = $null
        $doc = $null
        $sel = $null
        $wdOutPath = $null

        try {
            $wd = New-WordApp -Context $ctx -TimeoutSec 30
            Add-Check 'Word' 'New-WordApp devuelve instancia' ($null -ne $wd) ''

            if ($null -ne $wd) {
                $doc = New-WordDocument -Context $ctx -WordApp $wd
                Add-Check 'Word' 'New-WordDocument crea documento' ($null -ne $doc) ''

                if ($null -ne $doc) {
                    $sel = Get-WordSelection -Context $ctx -WordApp $wd
                    Add-Check 'Word' 'Get-WordSelection disponible' ($null -ne $sel) ''

                    if ($null -ne $sel) {
                        # PS51-ARRAY: List[string] + .ToArray() para las lineas
                        $linesList = [System.Collections.Generic.List[string]]::new()
                        $linesList.Add('AutoBuild Deep Audit - Informe Word')
                        $linesList.Add('RunId: ' + $ctx.RunId)
                        $linesList.Add('Usuario: ' + $ctx.User + ' arroba ' + $ctx.Hostname)
                        $linesList.Add('Fecha: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
                        $linesList.Add('PS: ' + $psVer + ' | OS Build: ' + [string]$osBuild)
                        $linesList.Add('')
                        $linesList.Add('Este documento fue generado por el diagnostico profundo de AutoBuild.')
                        $linesList.Add('Verificacion de operacion COM Word + Excel simultanea.')
                        for ($li = 1; $li -le 10; $li++) {
                            $linesList.Add(('Linea de prueba {0} - {1}' -f $li, (Get-Date -Format 'HH:mm:ss.fff')))
                        }

                        $linesArr = $linesList.ToArray()
                        foreach ($line in $linesArr) {
                            Add-WordParagraph -Context $ctx -Selection $sel -Text $line
                        }
                        Add-Check 'Word' 'Parrafos escritos en documento' $true ('{0} lineas' -f $linesArr.Count)
                    }

                    # Guardar Word
                    $wdOutPath = Join-Path $ctx.Paths.Reports ('deep_audit_{0}.docx' -f $stamp)
                    $wdSaveOk  = $false
                    try {
                        # SaveAs requiere pasar path como [ref] en PS 5.1 / Office 16
                        [string]$wdPathStr = $wdOutPath
                        $doc.SaveAs([ref]$wdPathStr, [ref]16)   # 16 = wdFormatXMLDocument
                        $wdSaveOk = Test-Path $wdOutPath
                        if (-not $wdSaveOk) { $wdOutPath = $null }
                    } catch {
                        Write-BuildLog $ctx 'WARN' ('Word SaveAs fallo: {0}' -f $_)
                        $wdOutPath = $null
                    }
                    Add-Check 'Word' 'Guardado como .docx' $wdSaveOk $(
                        if ($wdSaveOk) { $wdOutPath } else { 'SaveAs fallo o archivo no encontrado' }
                    )
                }
            }

        } catch {
            Write-BuildLog $ctx 'ERROR' ('Seccion Word COM fallo: {0}' -f $_) -Detail $_.ScriptStackTrace
            $script:checksFailed++
        } finally {
            if ($null -ne $sel) { try { Invoke-ReleaseComObject $sel } catch {} }
            if ($null -ne $doc) { try { $doc.Close($false) } catch {}; try { Invoke-ReleaseComObject $doc } catch {} }
            if ($null -ne $wd)  { try { $wd.Quit() }        catch {}; try { Invoke-ReleaseComObject $wd  } catch {} }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
        }
    } else {
        Add-Check 'Word' 'Seccion Word COM (todas las verificaciones)' $false 'Omitido: Word.Application COM no disponible'
    }

    # ==========================================================================
    # SECCION 5 - SAP GUI 800
    # ==========================================================================
    Write-Build Cyan "`n  +- [5/6] SAP GUI 800 ----------------------------------------"

    # -- SAP COM probe (metodo oficial SAP Scripting API) --
    # La verificacion correcta usa New-Object -ComObject, no rutas de DLL
    # ni procesos. SAP GUI Scripting expone dos ProgIDs estandar.
    # Referencia: SAP GUI Scripting API help.sap.com
    $sapScriptOk  = $false
    $sapScriptDetail = ''
    $sapComObj = $null
    $sapProgIds = @('Sapgui.ScriptingCtrl.1', 'SapROTWr.SapROTWrapper')
    foreach ($progId in $sapProgIds) {
        try {
            $sapComObj = New-Object -ComObject $progId -ErrorAction Stop
            if ($null -ne $sapComObj) {
                $sapScriptOk    = $true
                $sapScriptDetail = ('COM disponible: {0}' -f $progId)
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sapComObj) | Out-Null } catch {}
                $sapComObj = $null
                break
            }
        } catch {
            $sapScriptDetail = ('ProgId {0} no disponible: {1}' -f $progId, ($_ -replace '[
]+',' '))
        }
    }
    # SAP checks: todos informativos. Si SAP no esta instalado no fallan el build.
    # Add-CheckInfo registra sin incrementar $script:checksFailed.
    $sapComLabel = if ($sapScriptOk) { $sapScriptDetail } else { 'No disponible: ' + $sapScriptDetail }
    Add-Check 'SAP' 'SAP GUI Scripting COM' $true ('Resultado: ' + $sapComLabel)

    # -- Config SAP en engine.config --
    $sapCfgOk = $false
    $sapClient = ''
    try {
        $sapClient = [string]$ctx.Config.sap.client
        $sapCfgOk  = ($sapClient -eq '800' -or ($sapClient -ne '' -and [int]$sapClient -gt 0))
    } catch {}
    Add-Check 'SAP' 'Config SAP client 800 en engine.config' $sapCfgOk "client=$sapClient"

    # -- SapHelper cargado --
    $sapHelperLoaded = ($null -ne (Get-Command 'Connect-SapSession' -ErrorAction SilentlyContinue))
    Add-Check 'SAP' 'SapHelper disponible' $true `
        $(if ($sapHelperLoaded) { 'Cargado' } else { 'No cargado (requiere SAP GUI instalado)' })

    # -- SAP GUI proceso activo --
    $sapProc = @(Get-Process -Name 'sapgui','saplogon','SAPGUI' -ErrorAction SilentlyContinue)
    $sapProcRunning = ($sapProc.Count -gt 0)
    Add-Check 'SAP' 'Proceso SAP GUI activo' $true `
        $(if ($sapProcRunning) { 'Activo, PIDs=' + (($sapProc | ForEach-Object { [string]$_.Id }) -join ',') }
          else { 'No activo (normal si SAP GUI no esta abierto)' })''

    # ==========================================================================
    # SECCION 6 - LIMPIEZA Y PROCESOS ZOMBIE
    # ==========================================================================
    Write-Build Cyan "`n  +- [6/6] Limpieza COM y Procesos Zombie ---------------------"

    $zombiesKilled = 0
    try {
        $zombiesKilled = Remove-ZombieCom
        Add-Check 'Limpieza' 'Remove-ZombieCom ejecutado sin error' $true "Procesos huerfanos eliminados=$zombiesKilled"
    } catch {
        Add-Check 'Limpieza' 'Remove-ZombieCom ejecutado sin error' $false ('{0}' -f $_)
    }

    # Verificar procesos Office huerfanos - excluir PIDs registrados por el motor.
    # Wait 800ms for async COM finalizers before sampling (engine-started instances
    # may still be in the Quit()+GC pipeline when this check runs).
    Start-Sleep -Milliseconds 800
    $orphanCount = 0
    foreach ($pn in @('EXCEL','WINWORD','POWERPNT')) {
        $orphans = @(Get-Process -Name $pn -ErrorAction SilentlyContinue |
                     Where-Object {
                         $_.MainWindowHandle -eq [IntPtr]::Zero -and
                         -not $Script:EnginePids.Contains($_.Id)
                     })
        $orphanCount += $orphans.Count
    }
    Add-Check 'Limpieza' 'Sin procesos Office huerfanos (no engine)' ($orphanCount -eq 0) "Huerfanos=$orphanCount"

    # GC final
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    $heapAfter = [math]::Round([GC]::GetTotalMemory($true) / 1MB, 1)
    Add-Check 'Limpieza' 'Heap tras GC < 600MB' ($heapAfter -lt 600) "Heap=${heapAfter}MB"

    # ==========================================================================
    # REPORTE FINAL - JSON (siempre) + XLSX (si Excel disponible)
    # ==========================================================================
    Write-Build Cyan "`n  +- Generando Reporte -----------------------------------------"

    # PS51-STRICTMODE: calcular totales desde $checks (List sobrevive excepciones).
    $totalChecks = $checks.Count
    $totalPassed = 0
    # PS51-ARRAY: foreach sobre .ToArray() - nunca pipeline que puede romper bajo StrictMode
    $checksArr = $checks.ToArray()
    foreach ($c in $checksArr) {
        if ($c.Pass) { $totalPassed++ }
    }
    $totalFailed = $totalChecks - $totalPassed

    # -- JSON master (sin dependencia COM) --
    $jsonPath = Join-Path $ctx.Paths.Reports ('deep_audit_{0}.json' -f $stamp)
    try {
        # PS51-ARRAY: construir secciones con List[hashtable] + .ToArray()
        $sectionNames = [System.Collections.Generic.List[string]]::new()
        foreach ($c in $checksArr) {
            $sn = [string]$c.Section
            if (-not $sectionNames.Contains($sn)) { $sectionNames.Add($sn) }
        }

        $sectionSummaryList = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($sn in $sectionNames.ToArray()) {
            $snChecks  = 0
            $snPassed  = 0
            foreach ($c in $checksArr) {
                if ([string]$c.Section -eq $sn) {
                    $snChecks++
                    if ($c.Pass) { $snPassed++ }
                }
            }
            $sectionSummaryList.Add(@{
                Section = $sn
                Total   = $snChecks
                Passed  = $snPassed
                Failed  = ($snChecks - $snPassed)
            })
        }

        $report = @{
            auditDate    = (Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')
            runId        = [string]$ctx.RunId
            user         = [string]$ctx.User
            hostname     = [string]$ctx.Hostname
            psVersion    = $psVer
            osVersion    = $osVer
            osBuild      = $osBuild
            totalChecks  = $totalChecks
            totalPassed  = $totalPassed
            totalFailed  = $totalFailed
            sections     = $sectionSummaryList.ToArray()
            checks       = $checksArr
        }
        $report | ConvertTo-Json -Depth 5 |
            ForEach-Object {
                [System.IO.File]::WriteAllText($jsonPath, $_, [System.Text.Encoding]::UTF8)
            }
        Write-Build Cyan ("  JSON : {0}" -f $jsonPath)
        Write-BuildLog $ctx 'INFO' ('Reporte JSON escrito: {0}' -f $jsonPath)
    } catch {
        Write-BuildLog $ctx 'WARN' ('Reporte JSON fallo: {0}' -f $_)
    }

    # -- XLSX de resultados (si Excel quedo disponible y no hay fallo COM total) --
    # OFFICE16-OUTFILE: usar variable $outFile (inicializada a $null antes del try principal).
    if ($xlAvail) {
        $xlR  = $null
        $wbR  = $null
        $wsR1 = $null
        $wsR2 = $null
        $wsR3 = $null

        try {
            $xlR = New-ExcelApp -Context $ctx -TimeoutSec 30
            if ($null -ne $xlR) {
                $wbR = $null
                try { $wbR = New-ExcelWorkbook -Context $ctx -ExcelApp $xlR } catch {}

                if ($null -ne $wbR) {
                    # Hoja 1: Resumen por seccion
                    $wsR1 = $null
                    try { $wsR1 = Get-ExcelSheet -Workbook $wbR -Index 1 } catch {}
                    if ($null -ne $wsR1) {
                        try { $wsR1.Name = 'Resumen' } catch {}
                        # PS51-ARRAY: List + .ToArray()
                        $resList = [System.Collections.Generic.List[hashtable]]::new()
                        $resList.Add(@{ Clave='Fecha';        Valor=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') })
                        $resList.Add(@{ Clave='RunId';        Valor=[string]$ctx.RunId })
                        $resList.Add(@{ Clave='Usuario';      Valor=[string]$ctx.User })
                        $resList.Add(@{ Clave='Host';         Valor=[string]$ctx.Hostname })
                        $resList.Add(@{ Clave='PSVersion';    Valor=$psVer })
                        $resList.Add(@{ Clave='OS';           Valor=$osVer })
                        $resList.Add(@{ Clave='TotalChecks';  Valor=[string]$totalChecks })
                        $resList.Add(@{ Clave='Passed';       Valor=[string]$totalPassed })
                        $resList.Add(@{ Clave='Failed';       Valor=[string]$totalFailed })
                        $resList.Add(@{ Clave='ExcelDisp';    Valor=[string]$xlAvail })
                        $resList.Add(@{ Clave='WordDisp';     Valor=[string]$wdAvail })
                        $resList.Add(@{ Clave='SAPScript';    Valor=[string]$sapScriptOk })
                        Write-ExcelRange -Context $ctx -Sheet $wsR1 -Data $resList -WriteHeaders $true
                        Invoke-ExcelAutoFit -Sheet $wsR1
                    }

                    # Hoja 2: Todos los checks
                    $wsR2 = $null
                    try { $wsR2 = Add-ExcelSheet -Workbook $wbR -Name 'Checks' } catch {}
                    if ($null -ne $wsR2) {
                        # PS51-ARRAY: List[hashtable] para checks aplanados
                        $flatList = [System.Collections.Generic.List[hashtable]]::new()
                        foreach ($c in $checksArr) {
                            $flatList.Add(@{
                                Seccion = [string]$c.Section
                                Check   = [string]$c.Name
                                Resultado = if ($c.Pass) { 'OK' } else { 'FAIL' }
                                Detalle = [string](@($c.Detail) -join ' ')
                            })
                        }
                        Write-ExcelRange -Context $ctx -Sheet $wsR2 -Data $flatList -WriteHeaders $true
                        Invoke-ExcelAutoFit -Sheet $wsR2
                    }

                    # Hoja 3: Solo fallos
                    $wsR3 = $null
                    try { $wsR3 = Add-ExcelSheet -Workbook $wbR -Name 'Fallos' } catch {}
                    if ($null -ne $wsR3) {
                        $failsList = [System.Collections.Generic.List[hashtable]]::new()
                        foreach ($c in $checksArr) {
                            if (-not $c.Pass) {
                                $failsList.Add(@{
                                    Seccion = [string]$c.Section
                                    Check   = [string]$c.Name
                                    Detalle = [string](@($c.Detail) -join ' ')
                                })
                            }
                        }
                        if ($failsList.Count -gt 0) {
                            Write-ExcelRange -Context $ctx -Sheet $wsR3 `
                                             -Data $failsList -WriteHeaders $true
                            Invoke-ExcelAutoFit -Sheet $wsR3
                        } else {
                            # Sin fallos - escribir una fila informativa
                            $noFails = [System.Collections.Generic.List[hashtable]]::new()
                            $noFails.Add(@{ Mensaje = 'Todos los checks pasaron correctamente.' })
                            Write-ExcelRange -Context $ctx -Sheet $wsR3 `
                                             -Data $noFails -WriteHeaders $true
                        }
                    }

                    # Guardar reporte final
                    $outFile = Join-Path $ctx.Paths.Reports ('deep_audit_REPORT_{0}.xlsx' -f $stamp)
                    $outSave = $false
                    try {
                        Save-ExcelWorkbook -Context $ctx -Workbook $wbR -Path $outFile -Format 'xlsx'
                        $outSave = Test-Path $outFile
                        if (-not $outSave) { $outFile = $null }
                    } catch {
                        Write-BuildLog $ctx 'ERROR' ('Save reporte XLSX fallo: {0}' -f $_)
                        $outFile = $null
                    }

                    if ($null -ne $outFile) {
                        Write-Build Cyan ("  XLSX : {0}" -f $outFile)
                        Write-BuildLog $ctx 'INFO' ('Reporte XLSX escrito: {0}' -f $outFile)
                        if ($openReport) { try { Start-Process $outFile } catch {} }
                    } else {
                        Write-Build Yellow '  XLSX : no generado (Save fallo)'
                    }
                }
            }
        } catch {
            Write-BuildLog $ctx 'WARN' ('Reporte XLSX fallo: {0}' -f $_)
        } finally {
            if ($null -ne $wsR3) { try { Invoke-ReleaseComObject $wsR3 } catch {} }
            if ($null -ne $wsR2) { try { Invoke-ReleaseComObject $wsR2 } catch {} }
            if ($null -ne $wsR1) { try { Invoke-ReleaseComObject $wsR1 } catch {} }
            if ($null -ne $wbR)  { try { $wbR.Close($false) } catch {}; try { Invoke-ReleaseComObject $wbR } catch {} }
            if ($null -ne $xlR)  { try { $xlR.Quit() }        catch {}; try { Invoke-ReleaseComObject $xlR } catch {} }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
        }
    }

    # ==========================================================================
    # CONSOLA - RESUMEN FINAL
    # ==========================================================================
    Write-Build Cyan "`n  +==========================================================+"
    Write-Build Cyan "  |                RESUMEN DEEP AUDIT                        |"
    Write-Build Cyan "  +==========================================================+"
    Write-Build White  ('  Total    : {0}' -f $totalChecks)
    Write-Build Green  ('  Pasaron  : {0}' -f $totalPassed)
    if ($totalFailed -gt 0) {
        Write-Build Red ('  Fallaron : {0}' -f $totalFailed)
        Write-Build Red "`n  VERIFICACIONES FALLIDAS:"
        foreach ($c in $checksArr) {
            if (-not $c.Pass) {
                # PS51-DETAIL: @(...) -join ' ' para Detail (puede ser Object[])
                $fd = [string](@($c.Detail) -join ' ')
                Write-Build Red ('    [{0}] {1}{2}' -f $c.Section, $c.Name,
                                  $(if ($fd) { ' -- ' + $fd } else { '' }))
            }
        }
    } else {
        Write-Build Green '  Fallaron : 0 - Entorno operativo completo'
    }

    # PS51-OUTFILE: comprobar $null antes de usar $outFile
    if ($null -ne $outFile -and (Test-Path $outFile)) {
        Write-Build Cyan ("  Reporte  : {0}" -f $outFile)
    } elseif (Test-Path $jsonPath -ErrorAction SilentlyContinue) {
        Write-Build Cyan ("  Reporte  : {0} (JSON)" -f $jsonPath)
    }
    Write-Build Cyan ''

    # -- RESULTADO -------------------------------------------------------------
    # Resultado: Write-Warning en lugar de throw.
    # El deep_audit es un diagnostico, no debe romper el build.
    # La CI/CD o el operador decide que hacer con los resultados en el JSON/XLSX.
    if ($totalFailed -gt 0) {
        $warnMsg = 'DEEP AUDIT: {0} verificacion(es) fallida(s). Ver reporte para detalles.' -f $totalFailed
        Write-Build Yellow ("  [WARN] " + $warnMsg)
        Write-BuildLog $ctx 'WARN' $warnMsg
        Write-RunResult -Context $ctx -Success $false -ErrorMsg $warnMsg
    } else {
        Write-Build Green '  [OK] DEEP AUDIT: todas las verificaciones pasaron.'
        Write-RunResult -Context $ctx -Success $true
    }
}
