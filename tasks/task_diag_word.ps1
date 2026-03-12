#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_word.ps1
# @Description : Diagnostico completo de Word COM: instancia, documento,
#                escritura de parrafos, guardado .docx, exportacion a PDF,
#                liberacion y verificacion de zombis.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Prueba ciclo completo Word COM (instancia, escritura, .docx, PDF, liberacion)
task diag_word {
    $ctx = New-TaskContext `
        -TaskName 'diag_word' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de Word COM'

    $errores = 0
    $wd      = $null
    $doc     = $null

    # ---- 0. Limpiar zombis previos -----------------------------------------
    $zombisPrev = Remove-ZombieCom
    if ($zombisPrev -gt 0) {
        Write-Build Yellow "  [PRE]  Zombis previos eliminados: $zombisPrev"
        Start-Sleep -Seconds 2
    }

    # ---- 1. Localizar ejecutable Word -------------------------------------
    $exeWord = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($exeWord) {
        Write-Build Green "  [INST] WinWord.exe       : $exeWord"
    } else {
        Write-Build Yellow "  [INST] WinWord.exe       : no encontrado en rutas conocidas"
    }

    # ---- 2. Comprobar servidor COM -----------------------------------------
    Write-Build Cyan  "  [COM]  Comprobando disponibilidad del servidor COM (hasta 20s)..."
    $comOk = Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 20
    if ($comOk) {
        Write-Build Green "  [COM]  Servidor COM       : disponible"
    } else {
        Write-Build Red   "  [COM]  Servidor COM       : NO responde"
        Write-Build Red   "         Abrir Word manualmente, cerrar dialogos pendientes y reintentar"
        Write-BuildLog $ctx 'ERROR' 'Servidor COM de Word no disponible'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'COM no disponible'
        throw 'Word COM no disponible - diagnostico abortado'
    }

    # ---- 3. Ciclo completo COM ---------------------------------------------
    $selection = $null
    $charsCol  = $null
    try {
        # 3a. Instancia silenciosa
        Write-Build Cyan  "  [COM]  Instanciando Word..."
        $wd = New-WordApp -Context $ctx -TimeoutSec 30
        if ($null -eq $wd) {
            throw 'New-WordApp devolvio null'
        }
        Write-Build Green "  [COM]  Instancia          : OK"

        try {
            $ver = $wd.Version
            Write-Build Green "  [COM]  Version Word       : $ver"
        } catch {
            Write-Build Yellow "  [COM]  Version Word       : no disponible"
        }

        # 3b. Nuevo documento
        Write-Build Cyan  "  [DOC]  Creando documento..."
        $doc = New-WordDocument -Context $ctx -WordApp $wd
        Write-Build Green "  [DOC]  Documento          : OK"

        # 3c. Escribir contenido
        Write-Build Cyan  "  [DATA] Escribiendo contenido..."
        $lineas = @(
            "AutoBuild Diagnostico Word",
            "RunId : $($ctx.RunId)",
            "Fecha : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
            "Motor : Invoke-Build portable PS 5.1"
        )

        $selection = Get-WordSelection -WordApp $wd
        foreach ($linea in $lineas) {
            $selection.TypeText($linea)
            $selection.TypeParagraph()
        }
        Release-ComObject $selection ; $selection = $null

        # Verificar que el documento tiene contenido
        $charCount = Get-WordCharCount -Document $doc
        if ($charCount -gt 10) {
            Write-Build Green "  [DATA] Contenido          : OK ($charCount caracteres)"
        } else {
            Write-Build Red   "  [DATA] Contenido          : FAIL (solo $charCount caracteres)"
            $errores++
        }

        # 3d. Guardar como .docx
        $docxFile = Join-Path $ctx.Paths.Output "diag_word_$($ctx.RunId).docx"
        Write-Build Cyan  "  [SAVE] Guardando .docx..."
        Save-WordDocument -Context $ctx -Document $doc -Path $docxFile

        if (Test-Path $docxFile) {
            $size = (Get-Item $docxFile).Length
            Write-Build Green "  [SAVE] .docx guardado     : OK ($size bytes)"
        } else {
            Write-Build Red   "  [SAVE] .docx guardado     : FAIL (archivo no encontrado)"
            $errores++
        }

        # 3e. Exportar como PDF
        $pdfFile = Join-Path $ctx.Paths.Output "diag_word_$($ctx.RunId).pdf"
        Write-Build Cyan  "  [PDF]  Exportando a PDF via Word..."
        try {
            Export-WordToPdf -Context $ctx -Document $doc -Path $pdfFile
            if (Test-Path $pdfFile) {
                $sizePdf = (Get-Item $pdfFile).Length
                Write-Build Green "  [PDF]  PDF exportado      : OK ($sizePdf bytes)"
            } else {
                Write-Build Red   "  [PDF]  PDF exportado      : FAIL (archivo no encontrado)"
                $errores++
            }
        } catch {
            Write-Build Yellow "  [PDF]  PDF exportado      : WARN - $_"
        }

        Write-RunResult -Context $ctx -Success ($errores -eq 0)

    } catch {
        $errores++
        Write-Build Red "  [ERR]  Excepcion: $_"
        Write-BuildLog $ctx 'ERROR' "Excepcion en diag_word: $_"
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
        throw

    } finally {
        if ($null -ne $selection) { Release-ComObject $selection; $selection = $null }
        if ($null -ne $charsCol)  { Release-ComObject $charsCol;  $charsCol  = $null }

        Close-WordDocument -Document $doc -Save $false
        Close-WordApp      -WordApp $wd
        $doc = $null
        $wd  = $null

        $zombis = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($zombis -eq 0) {
            Write-Build Green "  [GC]   Procesos zombi     : ninguno OK"
        } else {
            Write-Build Yellow "  [GC]   Procesos zombi     : $zombis sin ventana detectados"
            Write-Build Yellow "         Ejecutar: .\Run.ps1 limpiar_com"
        }
    }

    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - Word COM funciona correctamente"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en el ciclo COM de Word"
        throw "diag_word detecto $errores error(es)"
    }

    Write-BuildLog $ctx 'INFO' "diag_word completado. Errores=$errores"
}
