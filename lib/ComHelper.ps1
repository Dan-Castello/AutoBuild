#Requires -Version 5.1
# lib/ComHelper.ps1
# Ayudas para COM seguro en PS 5.1.
# Incorpora leccion aprendida BUG-COM-FREEZE-01:
# New-Object -ComObject puede bloquearse si hay un dialogo modal pendiente.
# Solucion: verificar disponibilidad del servidor COM con Job + deadline.

Set-StrictMode -Version Latest

function Test-ComAvailable {
    <#
    .SYNOPSIS
        Comprueba si un servidor COM responde antes de intentar instanciarlo.
    .OUTPUTS
        $true si el servidor COM esta disponible, $false si no.
    #>
    param(
        [string]$ProgId,
        [int]$TimeoutSec = 20
    )

    $job = $null
    try {
        $job = Start-Job -ScriptBlock {
            param($pid_)
            try {
                $obj = New-Object -ComObject $pid_ -ErrorAction Stop
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
                return $true
            } catch {
                return $false
            }
        } -ArgumentList $ProgId

        $completed = Wait-Job -Job $job -Timeout $TimeoutSec
        if ($null -eq $completed) {
            return $false
        }
        $result = Receive-Job -Job $job
        return [bool]$result
    } catch {
        return $false
    } finally {
        if ($null -ne $job) {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-ComWithTimeout {
    <#
    .SYNOPSIS
        Instancia un objeto COM solo si el servidor responde antes del timeout.
        Devuelve $null si el servidor no responde.
    .NOTES
        BUG-COM-FREEZE-01: el objeto COM real se crea en el hilo principal
        porque los COM objects no son marshallables entre Jobs.
        El Job solo sirve para detectar si hay dialogo modal bloqueante.
    #>
    param(
        [hashtable]$Context,
        [string]$ProgId,
        [int]$TimeoutSec = 30,
        [string]$Label = ''
    )

    if (-not $Label) { $Label = $ProgId }
    Write-BuildLog $Context 'DEBUG' "Comprobando disponibilidad de $Label"

    if (-not (Test-ComAvailable -ProgId $ProgId -TimeoutSec $TimeoutSec)) {
        Write-BuildLog $Context 'WARN' "$Label no disponible en ${TimeoutSec}s. Resolver dialogo pendiente y reintentar."
        return $null
    }

    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        Write-BuildLog $Context 'DEBUG' "$Label instanciado correctamente"
        return $obj
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al instanciar $Label : $_"
        return $null
    }
}

function Release-ComObject {
    <#
    .SYNOPSIS
        Libera un objeto COM drenando todas sus referencias.
        Una sola llamada a ReleaseComObject decrementa en 1 el ref-count.
        Si hay multiples referencias vivas el proceso no termina.
        Este loop drena hasta que el ref-count llegue a cero o negativo.
    #>
    param($ComObject)

    if ($null -eq $ComObject) { return }
    try {
        $remaining = 1
        while ($remaining -gt 0) {
            $remaining = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
    } catch {
        # Ignorar - el objeto ya fue liberado
    }
}

function Invoke-ComCleanup {
    <#
    .SYNOPSIS
        Ejecuta la secuencia completa de limpieza COM.
        Llamar en el bloque finally de cualquier tarea que use COM.
    #>
    param(
        $Document,
        $Application,
        [hashtable]$Context
    )

    # 1. Cerrar documento
    if ($null -ne $Document) {
        try { $Document.Close($false) } catch {}
        Release-ComObject $Document
        $Document = $null
    }

    # 2. Cerrar aplicacion
    if ($null -ne $Application) {
        try { $Application.Quit() } catch {}
        Release-ComObject $Application
        $Application = $null
    }

    # 3. GC
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($null -ne $Context) {
        Write-BuildLog $Context 'DEBUG' 'COM liberado correctamente'
    }
}

function Remove-ZombieCom {
    <#
    .SYNOPSIS
        Elimina procesos Office sin ventana visible (zombis headless).
        Cubre Excel, Word, PowerPoint y Outlook.
    .OUTPUTS
        Numero de procesos eliminados.
    #>
    $count = 0
    foreach ($procName in @('EXCEL', 'WINWORD', 'POWERPNT', 'OUTLOOK')) {
        @(Get-Process -Name $procName -ErrorAction SilentlyContinue) | ForEach-Object {
            if ($_.MainWindowHandle -eq [IntPtr]::Zero) {
                try {
                    $_.Kill()
                    $count++
                } catch {}
            }
        }
    }
    return $count
}
