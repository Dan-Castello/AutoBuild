#Requires -Version 5.1
<#
.SYNOPSIS
    Verifica que todos los archivos .ps1 del proyecto sean ASCII puro.
    BUG-ASCII-01: PS 5.1 falla silenciosamente con caracteres no-ASCII.
.EXAMPLE
    .\tools\Test-Ascii.ps1
    .\tools\Test-Ascii.ps1 -Fix   # Reemplaza caracteres comunes automaticamente
#>
param([switch]$Fix)

Set-StrictMode -Version Latest

$Root    = Split-Path $PSScriptRoot -Parent
$found   = 0
$fixed   = 0

$replacements = @{
    [char]0x2014 = '--'     # em-dash
    [char]0x2013 = '-'      # en-dash
    [char]0x2018 = "'"      # comilla izquierda
    [char]0x2019 = "'"      # comilla derecha
    [char]0x201C = '"'      # comilla doble izquierda
    [char]0x201D = '"'      # comilla doble derecha
    [char]0x00E1 = 'a'      # a con acento
    [char]0x00E9 = 'e'      # e con acento
    [char]0x00ED = 'i'      # i con acento
    [char]0x00F3 = 'o'      # o con acento
    [char]0x00FA = 'u'      # u con acento
    [char]0x00F1 = 'n'      # enie
    [char]0x00C1 = 'A'
    [char]0x00C9 = 'E'
    [char]0x00CD = 'I'
    [char]0x00D3 = 'O'
    [char]0x00DA = 'U'
    [char]0x00D1 = 'N'
}

Get-ChildItem -Path $Root -Recurse -Filter '*.ps1' |
Where-Object { $_.FullName -notlike '*\tools\InvokeBuild\*' } |
ForEach-Object {
    $file  = $_
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $bad   = $bytes | Where-Object { $_ -gt 127 }

    if ($bad) {
        $found++
        Write-Host "No-ASCII: $($file.Name)" -ForegroundColor Red

        if ($Fix) {
            $text = Get-Content $file.FullName -Raw -Encoding UTF8
            foreach ($kv in $replacements.GetEnumerator()) {
                $text = $text -replace [regex]::Escape($kv.Key), $kv.Value
            }
            # Verificar si quedan no-ASCII
            $remainingBad = [System.Text.Encoding]::ASCII.GetBytes(
                [System.Text.Encoding]::ASCII.GetString(
                    [System.Text.Encoding]::UTF8.GetBytes($text)
                )
            ) | Where-Object { $_ -eq 63 }  # 63 = '?' substituto

            [System.IO.File]::WriteAllText($file.FullName, $text, [System.Text.Encoding]::ASCII)
            Write-Host "  -> Corregido (verificar manualmente)" -ForegroundColor Yellow
            $fixed++
        }
    }
}

Write-Host ''
if ($found -eq 0) {
    Write-Host 'Todos los archivos .ps1 son ASCII puro. OK' -ForegroundColor Green
} else {
    Write-Host "$found archivo(s) con caracteres no-ASCII encontrados." -ForegroundColor Red
    if ($Fix) {
        Write-Host "$fixed archivo(s) corregidos (verificar manualmente)." -ForegroundColor Yellow
    } else {
        Write-Host 'Ejecutar con -Fix para intentar correccion automatica.' -ForegroundColor Yellow
    }
}
