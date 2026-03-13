#Requires -Version 5.1
# =============================================================================
# lib/WordHelper.ps1
# Helpers Word COM para PS 5.1.
# Misma regla que ExcelHelper: todo COM object intermedio se captura
# y libera. Las cadenas inline generan fugas de ref-count.
# =============================================================================
# Solo ASCII. PS 5.1.

Set-StrictMode -Version Latest

function New-WordApp {
    <#
    .SYNOPSIS
        Crea una instancia silenciosa de Word con timeout de seguridad.
        Retorna $null si Word no esta disponible.
    #>
    param(
        [hashtable]$Context,
        [int]$TimeoutSec = 30
    )

    $wd = Invoke-ComWithTimeout -Context $Context -ProgId 'Word.Application' `
        -TimeoutSec $TimeoutSec -Label 'Word'

    if ($null -eq $wd) { return $null }

    try {
        $wd.Visible        = $false
        $wd.DisplayAlerts  = 0    # wdAlertsNone
        $wd.ScreenUpdating = $false
    } catch {
        Write-BuildLog $Context 'WARN' "Advertencia configurando Word: $_"
    }

    return $wd
}

function New-WordDocument {
    <#
    .SYNOPSIS
        Crea un nuevo documento vacio liberando la coleccion Documents.
    #>
    param(
        [hashtable]$Context,
        $WordApp
    )

    if ($null -eq $WordApp) { throw 'WordApp es nulo' }
    $docs = $null
    try {
        $docs = $WordApp.Documents
        $doc  = $docs.Add()
        Write-BuildLog $Context 'DEBUG' 'Documento Word creado'
        return $doc
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al crear documento Word: $_"
        throw
    } finally {
        if ($null -ne $docs) { Invoke-ReleaseComObject $docs; $docs = $null }
    }
}

function Open-WordDocument {
    <#
    .SYNOPSIS
        Abre un documento existente liberando la coleccion Documents.
    #>
    param(
        [hashtable]$Context,
        $WordApp,
        [string]$Path,
        [bool]$ReadOnly = $false
    )

    if ($null -eq $WordApp)     { throw 'WordApp es nulo' }
    if (-not (Test-Path $Path)) { throw "Archivo no encontrado: $Path" }

    $docs = $null
    try {
        $docs = $WordApp.Documents
        $doc  = $docs.Open($Path, $false, $ReadOnly)
        Write-BuildLog $Context 'DEBUG' "Documento Word abierto: $Path"
        return $doc
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al abrir documento Word $Path : $_"
        throw
    } finally {
        if ($null -ne $docs) { Invoke-ReleaseComObject $docs; $docs = $null }
    }
}

function Get-WordSelection {
    <#
    .SYNOPSIS
        Obtiene el objeto Selection de Word.
        El llamante es responsable de liberarlo con Release-ComObject.
    #>
    param($WordApp)

    if ($null -eq $WordApp) { throw 'WordApp es nulo' }
    return $WordApp.Selection
}

function Add-WordParagraph {
    <#
    .SYNOPSIS
        Escribe texto y un salto de parrafo usando Selection.
        Acepta el objeto Selection ya creado para evitar obtenerlo
        multiples veces (cada llamada a .Selection crea un COM object).
    #>
    param(
        [hashtable]$Context,
        $Selection,
        [string]$Text,
        [string]$StyleName = ''
    )

    if ($null -eq $Selection) { throw 'Selection es nulo' }
    try {
        if ($StyleName -ne '') {
            try { $Selection.Style = $StyleName } catch {}
        }
        $Selection.TypeText($Text)
        $Selection.TypeParagraph()
    } catch {
        Write-BuildLog $Context 'WARN' "Add-WordParagraph warn: $_"
    }
}

function Get-WordCharCount {
    <#
    .SYNOPSIS
        Devuelve el numero de caracteres del documento liberando
        la coleccion Characters inmediatamente.
    #>
    param($Document)

    $chars = $null
    try {
        $chars = $Document.Characters
        return $chars.Count
    } finally {
        if ($null -ne $chars) { Invoke-ReleaseComObject $chars; $chars = $null }
    }
}

function Save-WordDocument {
    <#
    .SYNOPSIS
        Guarda un documento en la ruta indicada.
        Formato 12 = wdFormatXMLDocument (.docx).
    #>
    param(
        [hashtable]$Context,
        $Document,
        [string]$Path,
        [int]$Format = 12
    )

    if ($null -eq $Document) { throw 'Document es nulo' }
    try {
        $Document.SaveAs2($Path, $Format)
        Write-BuildLog $Context 'INFO' "Documento Word guardado: $Path"
    } catch {
        try {
            $Document.SaveAs([ref]$Path, [ref]$Format)
        } catch {
            Write-BuildLog $Context 'ERROR' "Error al guardar documento Word $Path : $_"
            throw
        }
    }
}

function Export-WordToPdf {
    <#
    .SYNOPSIS
        Exporta el documento a PDF via ExportAsFixedFormat (no necesita impresora).
        wdExportFormatPDF = 17.
    #>
    param(
        [hashtable]$Context,
        $Document,
        [string]$Path
    )

    if ($null -eq $Document) { throw 'Document es nulo' }
    try {
        $Document.ExportAsFixedFormat($Path, 17)
        Write-BuildLog $Context 'INFO' "Documento exportado a PDF: $Path"
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al exportar PDF $Path : $_"
        throw
    }
}

function Close-WordDocument {
    <#
    .SYNOPSIS
        Cierra un documento, drena su referencia COM y dispara un ciclo GC.
    #>
    param(
        $Document,
        [bool]$Save = $false
    )

    if ($null -eq $Document) { return }
    $saveFlag = if ($Save) { -1 } else { 0 }    # wdSaveChanges / wdDoNotSaveChanges
    try { $Document.Close($saveFlag) } catch {}
    Invoke-ReleaseComObject $Document

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Close-WordApp {
    <#
    .SYNOPSIS
        Cierra Word, drena referencias COM, espera confirmacion del OS.
        Si el proceso sigue vivo al agotar el timeout lo elimina con Kill().
        Garantia: no quedan procesos WINWORD headless tras esta llamada.
    #>
    param(
        $WordApp,
        [int]$WaitSec = 15
    )

    if ($null -eq $WordApp) { return }
    try { $WordApp.Quit(0) } catch {}
    Invoke-ReleaseComObject $WordApp

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    $deadline = [datetime]::Now.AddSeconds($WaitSec)
    while ([datetime]::Now -lt $deadline) {
        $vivos = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero })
        if ($vivos.Count -eq 0) { return }
        Start-Sleep -Milliseconds 400
    }

    # Timeout agotado: forzar eliminacion de cualquier proceso headless restante
    @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }) | ForEach-Object {
        try { $_.Kill() } catch {}
    }
    Start-Sleep -Milliseconds 500
}
