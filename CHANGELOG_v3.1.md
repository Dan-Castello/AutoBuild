# AutoBuild v3.3 — Changelog

**Date:** 2026-03-13
**Base:** AutoBuild_v3_2_FIXED
**Type:** Refactoring mayor — eliminación de todas las tareas de diagnóstico previas
         y creación de tarea única `deep_audit` validada para el entorno de producción.

---

## Entorno objetivo

| Componente       | Versión                         |
|------------------|---------------------------------|
| Office           | 16 (Excel, Word)                |
| SAP GUI          | 800 Final — 8000.1.11.1161      |
| SAP Build        | 2263391 Patch 11                |
| PowerShell       | 5.1.22621.6345 (Desktop)        |
| Windows Build    | 10.0.22621.6345                 |

---

## Cambios v3.3

### Eliminadas (9 tareas)
- `task_diag_excel.ps1`
- `task_diag_full_audit.ps1`
- `task_diag_multi_com.ps1`
- `task_diag_csv.ps1`
- `task_diag_engine.ps1`
- `task_diag_notify.ps1`
- `task_diag_security.ps1`
- `task_diag_stress.ps1`
- `task_diag_word.ps1`

### Creada (1 tarea)
- `task_deep_audit.ps1` — Diagnóstico profundo y auditoría unificada

---

## Restricciones PS 5.1 aplicadas en `task_deep_audit.ps1`

### PS51-ARRAY — Eliminar `+=` sobre `[object[]]`
**Error original:** `op_Addition` en `[System.Object[]]`
**Causa:** El operador `+=` sobre arrays en PS 5.1 recrea el array completo en cada
iteración (O(n²)). Bajo `Set-StrictMode -Version Latest`, si el tipo resuelto es
`[System.Object[]]` en lugar de `[string[]]`, lanza `op_Addition`.

**Patrón eliminado:**
```powershell
$arr += @{ ... }   # PROHIBIDO en PS 5.1
```
**Patrón aplicado en toda la tarea:**
```powershell
$list = [System.Collections.Generic.List[hashtable]]::new()
$list.Add(@{ ... })
# ...
$arr = $list.ToArray()   # sólo al pasar a Write-ExcelRange o ConvertTo-Json
```
La **única** `+=` que permanece es `$orphanCount += $orphans.Count` sobre `[int]`,
que es segura y nunca lanza `op_Addition`.

---

### PS51-SORT — `Hashtable.Keys` sin `@()` wrapper
**Error original:** `op_Addition` durante `Sort-Object` sobre `IDictionaryKeyCollection`
**Afectaba:** `Write-ExcelRange` en `ExcelHelper.ps1` (corregido en v3.2) y cualquier
tarea que iterara `$ht.Keys | Sort-Object` sin materializar primero.

**Patrón aplicado:**
```powershell
$sortedKeys = @($dataArr[0].Keys) | Sort-Object   # @() materializa antes del sort
```
La nueva tarea no itera `.Keys` directamente; usa `[ordered]@{}` en todas las
definiciones de hashtable para garantizar orden de inserción sin `Sort-Object`.

---

### PS51-STRICTMODE — Variables sin inicializar (`VariableIsUndefined`)
**Error original:** `VariableIsUndefined` en `$outFile` (línea 279) y `$total`
**Causa:** `Set-StrictMode -Version Latest` lanza si una excepción temprana
salta el bloque de asignación, dejando variables referenciadas en SUMMARY sin valor.

**Patrón aplicado:** Todas las variables usadas fuera del `try` se inicializan
ANTES del primer `try`:
```powershell
$outFile             = $null    # PS51-OUTFILE
$outFile             = $null    # PS51-OUTFILE
$totalChecks         = 0        # PS51-STRICTMODE
$totalPassed         = 0
$script:checksFailed = 0
$xlsxPath            = $null
$csvPath             = $null
```

---

### PS51-DETAIL — `ConvertFrom-Json` produce `[Object[]]` en campo Detail
**Error original:** `op_Addition` durante concatenación `' -- ' + [Object[]]`
**Causa:** `ConvertFrom-Json` en PS 5.1 puede deserializar un campo `Detail` como
`[Object[]]` cuando el JSON fue serializado con `-Depth` insuficiente o el valor
original era un array.

**Patrón aplicado en `Add-Check` y en todo acceso a `$c.Detail`:**
```powershell
$safeDetail = [string](@($Detail) -join ' ')
# y en displays:
$fd = [string](@($c.Detail) -join ' ')
```
`@(...) -join ' '` aplana `$null`, `[string]`, `[Object[]]` y `PSCustomObject`
a `[string]` en todos los casos.

---

### OFFICE16-COM-ADD — `Workbooks.Add()` devuelve `$null` sin lanzar excepción
**Error original:** Fallos encadenados tras `New-ExcelWorkbook` cuando Office 16
retorna `$null` de `Workbooks.Add()` sin lanzar excepción (DDE deshabilitado,
macro policy Restricted, diálogo de activación pendiente, primer arranque).

**`New-ExcelWorkbook` en `ExcelHelper.ps1`** ya tiene guarda interna.
**Segunda capa en `deep_audit`:**
```powershell
try {
    $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
    # Guarda externa: verificar null post-llamada antes de cualquier operación
    if ($null -eq $wb) {
        Write-BuildLog $ctx 'ERROR' 'New-ExcelWorkbook retornó $null ...'
    }
    $wbCreated = ($null -ne $wb)
} catch { $wb = $null }
Add-Check 'Excel' 'Workbook creado' $wbCreated $(...)

if ($null -ne $wb) {
    # TODAS las operaciones de sheet/write/save dentro de este bloque
} else {
    # Registrar todas las verificaciones dependientes como FAIL con motivo claro
}
```

---

## Estructura de la tarea `deep_audit`

| Sección | Verificaciones |
|---------|---------------|
| [1/6] Entorno y Runtime | PS version, edición, OS build, CLR, arquitectura, disco, heap, admin, StrictMode |
| [2/6] Seguridad y Config | engine.config.json, secciones obligatorias, logLevel, maxRetries, modo seguridad |
| [3/6] Excel COM (Office 16) | Disponibilidad, PID, Workbook (null-safe), escritura masiva 2D, UsedRange, XLSX, CSV, read-back |
| [4/6] Word COM (Office 16) | Disponibilidad, instancia, documento, párrafos, SaveAs docx |
| [5/6] SAP GUI 800 | DLL scripting, proceso activo, registro COM, config client 800, SapHelper |
| [6/6] Limpieza | Remove-ZombieCom, procesos huérfanos, GC heap post-cleanup |

**Salidas:**
- `deep_audit_<stamp>.json` — siempre, sin dependencia COM
- `deep_audit_<stamp>.xlsx` — datos de diagnóstico (Sheet 1: datos)
- `deep_audit_<stamp>.csv` — export CSV de los datos
- `deep_audit_<stamp>.docx` — documento Word (si Word disponible)
- `deep_audit_REPORT_<stamp>.xlsx` — reporte consolidado (Resumen / Checks / Fallos)

**Parámetros:**
```
RowCount         (default: 200)   — filas en prueba de escritura masiva Excel
OpenReport       (default: false) — abrir XLSX reporte al finalizar
StopOnComFailure (default: false) — omitir secciones Word/SAP si Excel no disponible
```

---

# AutoBuild v3.2 — Changelog (histórico)

**Date:** 2026-03-13
**Fixes:** FIX-WRITE-RANGE-KEYS, FIX-EXCEL-OP-ADDITION, FIX-EXCEL-UNINIT,
           FIX-FULL-AUDIT-DETAIL-V2, FIX-MULTI-COM-OP-ADDITION

---

# AutoBuild v3.1 — Changelog (histórico)

**Date:** 2026-03-13
**Fixes:** FIX-EXCEL-NULL, Config.ps1 hardening, SMTP optional, engine diagnostics

