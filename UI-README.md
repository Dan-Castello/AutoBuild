# AutoBuild Automation Interface

Production-ready, secure, and auditable WPF GUI for AutoBuild.
Fully integrated with Run.ps1 and the Invoke-Build engine.

---

## Requirements

| Requirement       | Minimum          | Notes                              |
|-------------------|------------------|------------------------------------|
| PowerShell        | 5.1 Desktop      | Core edition has limited COM       |
| Windows           | 10 / 11          | .NET Framework 4.x required        |
| .NET Framework    | 4.6+             | For WPF assemblies                 |
| AutoBuild Engine  | v1.4.x+          | Run.ps1 must be present            |

---

## Installation

1. Copy all files from this folder into your AutoBuild root directory
   (same folder that contains `Run.ps1`, `engine.config.json`, and the `tasks/` folder).

2. Unblock scripts if downloaded from the internet:

   ```powershell
   Get-ChildItem -Path . -Filter *.ps1 | Unblock-File
   ```

3. Set execution policy (if needed):

   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```

---

## Launch

**Option 1 — Double-click:**
```
Launch-AutoBuildUI.bat
```

**Option 2 — PowerShell (recommended):**
```powershell
.\Start-AutoBuildUI.ps1                          # Operator role (default)
.\Start-AutoBuildUI.ps1 -Role Developer          # Developer role
.\Start-AutoBuildUI.ps1 -Role Admin              # Admin role
.\Start-AutoBuildUI.ps1 -Role Admin -EnginePath "C:\path\to\AutoBuild"
```

**Option 3 — Direct:**
```powershell
powershell.exe -STA -File AutoBuild.UI.ps1 -Role Admin
```

> WPF requires STA (single-threaded apartment) mode.
> `Start-AutoBuildUI.ps1` handles this automatically.

---

## Role-Based Access Control

| Feature                    | Operator | Developer | Admin |
|----------------------------|:--------:|:---------:|:-----:|
| Browse task catalog        |    YES   |    YES    |  YES  |
| Execute tasks              |    YES   |    YES    |  YES  |
| View execution history     |    YES   |    YES    |  YES  |
| View artifacts             |    YES   |    YES    |  YES  |
| View metrics dashboard     |    YES   |    YES    |  YES  |
| Run environment diagnostics|    YES   |    YES    |  YES  |
| Create / edit task files   |     NO   |    YES    |  YES  |
| Edit engine.config.json    |     NO   |     NO    |  YES  |
| Delete artifacts           |     NO   |     NO    |  YES  |
| Manage checkpoints         |     NO   |     NO    |  YES  |
| View audit log             |     NO   |     NO    |  YES  |

---

## Files

| File                        | Purpose                                                  |
|-----------------------------|----------------------------------------------------------|
| `AutoBuild.UI.ps1`          | Main WPF application (all UI logic, RBAC, data access)  |
| `Start-AutoBuildUI.ps1`     | Smart launcher (handles STA mode, validation)            |
| `Launch-AutoBuildUI.bat`    | Windows batch launcher for double-click use              |
| `Invoke-UITask.ps1`         | Thin wrapper around Run.ps1 for UI-initiated executions  |
| `Invoke-RetentionCleanup.ps1` | Artifact retention policy enforcement script           |

---

## Pages

### Task Catalog
Dynamically scans `tasks/task_*.ps1` and displays name, category, description,
version, last run time, and last status. Supports search and category filtering.
Double-click or "Execute Selected" to jump to the execution panel.

### Execute Task
Select a task, fill in the dynamically generated parameter form,
choose execution mode (normal, WhatIf, checkpoint, resume),
and launch via Run.ps1. Live console output streams in real time.

### Live Monitor
Shows recent runs and live log tail from `logs/registry.jsonl`.
Auto-refreshes every 5 seconds.

### Execution History
Browse all run summaries grouped by RunId. Filter by task name and status.
Click a row to view detailed log entries for that run. Export to CSV.

### Checkpoint Manager
Lists `logs/checkpoint_*.clixml` files. Resume, delete, or inspect checkpoint state.
Requires Admin role to delete.

### Artifact Repository
Browses `output/` and `reports/` directories. Open, save-as, or delete artifacts.
Delete requires Admin role and is audit-logged.

### Metrics & Observability
KPI cards: total runs, success rate, average duration, top task.
Task frequency table with ASCII bar chart. Recent error log.

### Environment Diagnostics
Validates: PS version, engine files, folder permissions, COM availability
(Excel, Word, SAP GUI), and Invoke-Build tool presence.

### Configuration
JSON editor for `engine.config.json` with inline validation and guidance.
Admin role required. All saves are audit-logged.

### Create New Task
Calls `New-Task.ps1` with name validation and live file preview.
Developer or Admin role required.

### Audit Log
Read-only view of `logs/audit.jsonl`. Every execution, config change,
artifact deletion, and UI action is recorded. Admin role required. Export to CSV.

---

## Audit Log Events

| Action             | Trigger                                      |
|--------------------|----------------------------------------------|
| `UI_OPEN`          | Application launched                         |
| `UI_CLOSE`         | Application closed                           |
| `EXECUTE_TASK`     | Task execution started                       |
| `CANCEL_TASK`      | Running task cancelled                       |
| `EDIT_CONFIG`      | engine.config.json saved                     |
| `CREATE_TASK`      | New task file created                        |
| `DELETE_ARTIFACT`  | Artifact deleted                             |
| `DOWNLOAD_ARTIFACT`| Artifact saved to local path                 |
| `OPEN_ARTIFACT`    | Artifact opened in default application       |
| `DELETE_CHECKPOINT`| Checkpoint file deleted                      |
| `RESUME_CHECKPOINT`| Checkpoint resume initiated                  |
| `RUN_DIAGNOSTICS`  | Environment diagnostics executed             |
| `EXPORT_HISTORY`   | Execution history exported to CSV            |
| `EXPORT_AUDIT`     | Audit log exported to CSV                    |

---

## Artifact Retention

Run retention cleanup manually or schedule it:

```powershell
# Preview (no deletion)
.\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild" -WhatIf

# Execute
.\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild"
```

Retention period is read from `engine.config.json` (`reports.retentionDays`).
Default: 30 days.

---

## COM Safety Notes

All COM interactions in AutoBuild tasks use the project COM helpers:
- `ComHelper.ps1` — zombie process detection and cleanup
- `ExcelHelper.ps1` — `New-ExcelApp`, `Close-ExcelApp`, `Close-ExcelWorkbook`
- `SapHelper.ps1` — `Get-SapSession`, `Invoke-SapTransaction`
- `WordHelper.ps1` — `New-WordApp`, `Close-WordDocument`

The UI never bypasses these helpers. All task execution goes through `Run.ps1`.
Never implement task logic in the UI layer.

---

## Troubleshooting

**"WPF assemblies not available"**
> Install .NET Framework 4.6 or higher (included in Windows 10/11 by default).

**"Run.ps1 not found"**
> Copy the UI files to the same folder as `Run.ps1`, or use `-EnginePath`.

**Window does not appear / crashes on launch**
> Run via `Start-AutoBuildUI.ps1`. Direct execution of `AutoBuild.UI.ps1`
> may fail without STA mode if the calling shell is not STA.

**Execution policy error**
> `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

---

## Architecture

```
AutoBuild.UI.ps1
    |
    +-- RBAC (Test-Permission)
    +-- Audit (Write-AuditLog -> logs/audit.jsonl)
    +-- Task Parser (Get-TaskMetadata / Get-AllTasks)
    +-- Log Reader (Get-ExecutionHistory / Get-RunSummaries)
    +-- Metrics (Get-Metrics)
    +-- Diagnostics (Get-DiagnosticReport)
    +-- Config (Get-Config / Save-Config)
    +-- Job Manager (Start-TaskExecution / Get-JobOutput)
    |       |
    |       +-- powershell.exe -STA -File Run.ps1 [args]
    |               |
    |               +-- tools\InvokeBuild\Invoke-Build.ps1
    |                       |
    |                       +-- engine\Main.build.ps1
    |                               |
    |                               +-- tasks\task_*.ps1
    |                               +-- lib\*.ps1
    |
    +-- WPF Window (XAML)
            +-- 11 pages (Catalog, Execute, Monitor, History,
                           Checkpoints, Artifacts, Metrics,
                           Diagnostics, Config, NewTask, Audit)
```

The UI layer never imports engine libraries directly.
Engine state is read only through `logs/registry.jsonl` and file system scanning.
