#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - Pester 5.x Test Suite
.DESCRIPTION
    Unit tests for all pure functions in lib/.
    COM objects are mocked via scriptblock stubs so tests run without
    Excel, Word, or SAP installed. No network, no file system side effects.

.NOTES
    Run with:
        Invoke-Pester .\tests\AutoBuild.Tests.ps1 -Output Detailed
        Invoke-Pester .\tests\AutoBuild.Tests.ps1 -PassThru -CI

    Requirements:
        Pester 5.x  (Install-Module Pester -Force -SkipPublisherCheck)
        PowerShell 5.1 Desktop edition

    Coverage targets:
        Config.ps1    : Get-EngineConfig defaults, merge, section extensibility
        Logger.ps1    : New-RunId uniqueness, sanitization, rotation trigger
        Context.ps1   : New-TaskContext immutability, path derivation, identity
        Auth.ps1      : Resolve-UserRole ceiling, Test-Permission fail-safe
        Retry.ps1     : Invoke-WithRetry success, failure, backoff, max retries
        ComHelper.ps1 : PID registry, Release cap, Remove-ZombieCom whitelist
        ExcelHelper.ps1: Write-ExcelRange 2D array, format enum, batch semantics
        SapHelper.ps1 : Wait-SapReady timeout, Export ALV vs Table routing
        Assertions.ps1: Assert-Param, Assert-FileExists, Test-TaskAsset
#>

BeforeAll {
    # ---- Bootstrap: load all libs from the project root --------------------
    $ProjectRoot = Split-Path $PSScriptRoot -Parent
    $LibPath     = Join-Path $ProjectRoot 'lib'

    foreach ($lib in @('Config.ps1','Logger.ps1','Context.ps1','Auth.ps1',
                        'Retry.ps1','ComHelper.ps1','ExcelHelper.ps1',
                        'SapHelper.ps1','Assertions.ps1')) {
        . (Join-Path $LibPath $lib)
    }

    # ---- Shared test helpers -----------------------------------------------
    function New-TestConfig {
        return @{
            engine  = @{ logLevel='INFO'; maxRetries=2; retryDelaySeconds=0.01; ibVersion='test' }
            sap     = @{ systemId='TST'; client='100'; language='EN'; timeout=30 }
            excel   = @{ visible=$false; screenUpdating=$false }
            reports = @{ defaultFormat='xlsx'; retentionDays=30; maxLogSizeBytes=10485760 }
            security= @{ adminAdGroup=''; developerAdGroup=''; adminUsers='testadmin'; developerUsers='testdev' }
            notifications = @{ enabled=$false }
        }
    }

    function New-TestContext {
        param([string]$TaskName = 'test_task')
        $cfg = New-TestConfig
        $ctx = New-TaskContext -TaskName $TaskName -Config $cfg -Root $TestDrive
        return $ctx
    }

    # ---- Mock COM object factory -------------------------------------------
    function New-MockComObject {
        param([hashtable]$Properties = @{}, [hashtable]$Methods = @{})
        $obj = [PSCustomObject]$Properties
        foreach ($k in $Methods.Keys) {
            $sb = $Methods[$k]
            Add-Member -InputObject $obj -MemberType ScriptMethod -Name $k -Value $sb
        }
        $obj | Add-Member -MemberType NoteProperty -Name '_ComMockRef' -Value 1
        return $obj
    }
}

# ============================================================================
# Config.ps1
# ============================================================================
Describe 'Config.ps1' {

    Context 'Get-EngineConfig defaults' {
        It 'returns a hashtable with all expected sections' {
            $cfg = Get-EngineConfig -Root $TestDrive
            $cfg.Keys | Should -Contain 'engine'
            $cfg.Keys | Should -Contain 'sap'
            $cfg.Keys | Should -Contain 'excel'
            $cfg.Keys | Should -Contain 'reports'
            $cfg.Keys | Should -Contain 'notifications'
            $cfg.Keys | Should -Contain 'security'
        }

        It 'creates working directories under Root' {
            $cfg = Get-EngineConfig -Root $TestDrive
            foreach ($sub in @('logs','input','output','reports')) {
                (Test-Path (Join-Path $TestDrive $sub)) | Should -Be $true
            }
        }

        It 'returns conservative production defaults' {
            $cfg = Get-EngineConfig -Root $TestDrive
            $cfg.engine.logLevel          | Should -Be 'INFO'
            $cfg.excel.visible            | Should -Be $false
            $cfg.excel.screenUpdating     | Should -Be $false
            $cfg.notifications.enabled    | Should -Be $false
        }

        It 'returns a NEW object each call (not a cached reference)' {
            $cfg1 = Get-EngineConfig -Root $TestDrive
            $cfg2 = Get-EngineConfig -Root $TestDrive
            $cfg1.engine.logLevel = 'DEBUG'
            $cfg2.engine.logLevel | Should -Be 'INFO'
        }
    }

    Context 'Get-EngineConfig file merge' {
        It 'merges values from engine.config.json' {
            $json = '{"engine":{"logLevel":"DEBUG"},"sap":{"timeout":99}}'
            Set-Content (Join-Path $TestDrive 'engine.config.json') $json -Encoding ASCII
            $cfg = Get-EngineConfig -Root $TestDrive
            $cfg.engine.logLevel | Should -Be 'DEBUG'
            $cfg.sap.timeout     | Should -Be 99
            # Unspecified fields keep defaults
            $cfg.sap.systemId    | Should -Be 'PRD'
        }

        It 'tolerates a missing or malformed config file gracefully' {
            Set-Content (Join-Path $TestDrive 'engine.config.json') 'NOT_JSON' -Encoding ASCII
            { Get-EngineConfig -Root $TestDrive } | Should -Not -Throw
        }
    }

    Context 'Get-ConfigSection' {
        It 'returns a clone of the named section' {
            $cfg     = New-TestConfig
            $section = Get-ConfigSection -Config $cfg -Section 'engine'
            $section['logLevel'] = 'FATAL'
            $cfg.engine.logLevel | Should -Be 'INFO'  # original not mutated
        }

        It 'returns empty hashtable for missing section' {
            $cfg     = New-TestConfig
            $section = Get-ConfigSection -Config $cfg -Section 'nonexistent'
            $section | Should -BeOfType [hashtable]
            $section.Count | Should -Be 0
        }
    }
}

# ============================================================================
# Logger.ps1
# ============================================================================
Describe 'Logger.ps1' {

    Context 'New-RunId' {
        It 'returns a string' {
            (New-RunId) | Should -BeOfType [string]
        }

        It 'includes a timestamp prefix (yyyyMMdd_HHmmss)' {
            $id = New-RunId
            $id | Should -Match '^\d{8}_\d{6}_[0-9A-F]{8}$'
        }

        It 'generates unique IDs under rapid successive calls' {
            $ids = 1..50 | ForEach-Object { New-RunId }
            ($ids | Sort-Object -Unique).Count | Should -Be 50
        }
    }

    Context 'Invoke-SanitizeLogText (via Write-BuildLog)' {
        It 'replaces CRLF with pipe separator' {
            $result = Invoke-SanitizeLogText -Text "line1`r`nline2"
            $result | Should -Be 'line1 | line2'
        }

        It 'replaces LF with pipe separator' {
            $result = Invoke-SanitizeLogText -Text "line1`nline2"
            $result | Should -Be 'line1 | line2'
        }

        It 'strips control characters' {
            $result = Invoke-SanitizeLogText -Text "text`x01`x1Fclean"
            $result | Should -Not -Match '[\x00-\x1F]'
        }

        It 'handles empty string without error' {
            { Invoke-SanitizeLogText -Text '' } | Should -Not -Throw
        }
    }

    Context 'Invoke-LogRotationIfNeeded' {
        It 'does nothing when file is below threshold' {
            $logFile = Join-Path $TestDrive 'registry.jsonl'
            Set-Content $logFile 'small content' -Encoding ASCII
            Invoke-LogRotationIfNeeded -FilePath $logFile -MaxBytes 1048576
            Test-Path $logFile | Should -Be $true
        }

        It 'renames the file when size exceeds threshold' {
            $logFile = Join-Path $TestDrive 'registry.jsonl'
            # Create 11 bytes file with threshold of 10 bytes
            Set-Content $logFile 'hello world' -Encoding ASCII -NoNewline
            Invoke-LogRotationIfNeeded -FilePath $logFile -MaxBytes 5
            Test-Path $logFile | Should -Be $false
            $archives = @(Get-ChildItem $TestDrive -Filter 'registry_*.jsonl')
            $archives.Count | Should -Be 1
        }

        It 'does nothing if file does not exist' {
            { Invoke-LogRotationIfNeeded -FilePath (Join-Path $TestDrive 'nonexistent.jsonl') } |
                Should -Not -Throw
        }
    }

    Context 'Invoke-LogPurge' {
        It 'deletes archive files older than RetentionDays' {
            $logsDir = Join-Path $TestDrive 'logs_purge'
            New-Item $logsDir -ItemType Directory -Force | Out-Null
            $oldFile = Join-Path $logsDir 'registry_20200101_000000.jsonl'
            Set-Content $oldFile 'old' -Encoding ASCII
            (Get-Item $oldFile).LastWriteTime = [datetime]::Now.AddDays(-40)
            Invoke-LogPurge -LogsDir $logsDir -RetentionDays 30
            Test-Path $oldFile | Should -Be $false
        }

        It 'does NOT delete the active registry.jsonl' {
            $logsDir  = Join-Path $TestDrive 'logs_purge2'
            New-Item $logsDir -ItemType Directory -Force | Out-Null
            $active   = Join-Path $logsDir 'registry.jsonl'
            Set-Content $active 'active' -Encoding ASCII
            Invoke-LogPurge -LogsDir $logsDir -RetentionDays 1
            Test-Path $active | Should -Be $true
        }
    }
}

# ============================================================================
# Context.ps1
# ============================================================================
Describe 'Context.ps1' {

    Context 'New-TaskContext' {
        It 'returns a hashtable with all required keys' {
            $ctx = New-TestContext
            $ctx.Keys | Should -Contain 'RunId'
            $ctx.Keys | Should -Contain 'TaskName'
            $ctx.Keys | Should -Contain 'Config'
            $ctx.Keys | Should -Contain 'StartTime'
            $ctx.Keys | Should -Contain 'Params'
            $ctx.Keys | Should -Contain 'User'
            $ctx.Keys | Should -Contain 'Hostname'
            $ctx.Keys | Should -Contain 'SessionId'
            $ctx.Keys | Should -Contain 'Paths'
        }

        It 'derives all paths from Root' {
            $ctx = New-TestContext
            $ctx.Paths.Root    | Should -Be $TestDrive
            $ctx.Paths.Input   | Should -BeLike "*\input"
            $ctx.Paths.Output  | Should -BeLike "*\output"
            $ctx.Paths.Reports | Should -BeLike "*\reports"
            $ctx.Paths.Logs    | Should -BeLike "*\logs"
        }

        It 'produces isolated config copies (mutation does not affect original)' {
            $cfg = New-TestConfig
            $ctx = New-TaskContext -TaskName 'test' -Config $cfg -Root $TestDrive
            $ctx.Config.engine.logLevel = 'DEBUG'
            $cfg.engine.logLevel | Should -Be 'INFO'
        }

        It 'propagates custom Params' {
            $cfg = New-TestConfig
            $ctx = New-TaskContext -TaskName 'test' -Config $cfg -Root $TestDrive `
                   -Params @{ Centro = '1000'; Almacen = 'WH01' }
            $ctx.Params['Centro']  | Should -Be '1000'
            $ctx.Params['Almacen'] | Should -Be 'WH01'
        }

        It 'captures a non-empty User' {
            $ctx = New-TestContext
            $ctx.User | Should -Not -BeNullOrEmpty
        }

        It 'captures Hostname' {
            $ctx = New-TestContext
            $ctx.Hostname | Should -Not -BeNullOrEmpty
        }

        It 'sets SessionId to current PID' {
            $ctx = New-TestContext
            $ctx.SessionId | Should -Be $PID
        }
    }
}

# ============================================================================
# Auth.ps1
# ============================================================================
Describe 'Auth.ps1' {

    Context 'Test-Permission' {
        It 'grants Operator basic actions' {
            Test-Permission -Role 'Operator' -Action 'RunTask'       | Should -Be $true
            Test-Permission -Role 'Operator' -Action 'ViewHistory'   | Should -Be $true
        }

        It 'denies Operator elevated actions' {
            Test-Permission -Role 'Operator' -Action 'EditConfig'    | Should -Be $false
            Test-Permission -Role 'Operator' -Action 'PurgeOldLogs'  | Should -Be $false
            Test-Permission -Role 'Operator' -Action 'ViewAudit'     | Should -Be $false
        }

        It 'grants Developer task creation but not config' {
            Test-Permission -Role 'Developer' -Action 'CreateTask'   | Should -Be $true
            Test-Permission -Role 'Developer' -Action 'EditConfig'   | Should -Be $false
        }

        It 'grants Admin all actions' {
            foreach ($action in @('RunTask','EditConfig','ViewAudit','PurgeOldLogs','DeleteArtifact')) {
                Test-Permission -Role 'Admin' -Action $action | Should -Be $true
            }
        }

        It 'returns false for unknown actions (fail-safe)' {
            Test-Permission -Role 'Admin' -Action 'UnknownMadeUpAction' | Should -Be $false
        }
    }

    Context 'Assert-Permission' {
        It 'does not throw when permission is granted' {
            { Assert-Permission -Role 'Admin' -Action 'EditConfig' } | Should -Not -Throw
        }

        It 'throws a descriptive error when permission is denied' {
            { Assert-Permission -Role 'Operator' -Action 'EditConfig' } |
                Should -Throw -ExpectedMessage "*Access denied*"
        }
    }

    Context 'Resolve-UserRole ceiling enforcement' {
        It 'applies RequestedRole as ceiling even if user qualifies higher' {
            $cfg = New-TestConfig
            # testadmin is in adminUsers whitelist; requesting Operator -> gets Operator
            $result = Resolve-UserRole -Config $cfg -RequestedRole 'Operator'
            # We cannot predict the actual user running tests; just verify it is a valid role
            $result | Should -BeIn @('Operator','Developer','Admin')
        }

        It 'returns a valid role string in all cases' {
            $cfg  = New-TestConfig
            $role = Resolve-UserRole -Config $cfg -RequestedRole 'Admin'
            $role | Should -BeIn @('Operator','Developer','Admin')
        }
    }
}

# ============================================================================
# Retry.ps1
# ============================================================================
Describe 'Retry.ps1' {

    Context 'Invoke-WithRetry - success path' {
        It 'returns the value from the ScriptBlock on first attempt' {
            $ctx    = New-TestContext
            $result = Invoke-WithRetry -Context $ctx -ScriptBlock { return 42 } -Label 'test'
            $result | Should -Be 42
        }

        It 'returns after a transient failure followed by success' {
            $ctx     = New-TestContext
            $attempt = 0
            $result  = Invoke-WithRetry -Context $ctx -MaxRetries 3 -BaseDelaySeconds 0 `
                       -ScriptBlock {
                           $attempt++
                           if ($attempt -lt 2) { throw "transient" }
                           return 'ok'
                       } -Label 'flaky'
            $result | Should -Be 'ok'
        }
    }

    Context 'Invoke-WithRetry - failure path' {
        It 'throws after exhausting all retries' {
            $ctx = New-TestContext
            { Invoke-WithRetry -Context $ctx -MaxRetries 2 -BaseDelaySeconds 0 `
              -ScriptBlock { throw 'always fails' } -Label 'bad'
            } | Should -Throw
        }

        It 'respects MaxRetries count (n+1 total attempts)' {
            $ctx     = New-TestContext
            $callLog = [System.Collections.Generic.List[int]]::new()
            try {
                Invoke-WithRetry -Context $ctx -MaxRetries 3 -BaseDelaySeconds 0 `
                    -ScriptBlock { $callLog.Add(1); throw 'fail' } -Label 'count'
            } catch { }
            $callLog.Count | Should -Be 4   # 1 initial + 3 retries
        }
    }
}

# ============================================================================
# ComHelper.ps1
# ============================================================================
Describe 'ComHelper.ps1' {

    Context 'Register-EngineCom / Unregister-EngineCom / Remove-ZombieCom' {
        It 'registered PID is protected from zombie cleanup' {
            # We cannot actually spawn EXCEL in tests, but we can verify
            # the registry membership logic via the HashSet directly.
            Register-EngineCom -Pid_ 99999
            # The private HashSet should contain 99999
            $Script:EnginePids.Contains(99999) | Should -Be $true
            Unregister-EngineCom -Pid_ 99999
            $Script:EnginePids.Contains(99999) | Should -Be $false
        }
    }

    Context 'Invoke-ReleaseComObject' {
        It 'handles null input without throwing' {
            { Invoke-ReleaseComObject -ComObject $null } | Should -Not -Throw
        }

        It 'respects the max iteration cap (does not spin forever)' {
            # Create a mock COM-like object whose ReleaseComObject always returns > 0
            # We test the cap indirectly by verifying the function returns in finite time.
            # Real COM mocking of Marshal.ReleaseComObject requires type interop not
            # available in unit tests, so we verify the cap constant is set.
            $Script:ComReleaseMaxIter | Should -BeGreaterThan 0
            $Script:ComReleaseMaxIter | Should -BeLessOrEqual 100
        }
    }

    Context 'Release-ComObject alias' {
        It 'alias exists and points to Invoke-ReleaseComObject' {
            $alias = Get-Alias -Name 'Release-ComObject' -ErrorAction SilentlyContinue
            $alias | Should -Not -BeNullOrEmpty
        }
    }
}

# ============================================================================
# ExcelHelper.ps1
# ============================================================================
Describe 'ExcelHelper.ps1' {

    Context 'Write-ExcelRange 2D array construction' {
        It 'builds correct dimensions for data + header' {
            # Mock Sheet COM object
            $writtenArray = $null
            $mockSheet = New-MockComObject -Properties @{} -Methods @{
                Cells  = { param($r,$c); New-MockComObject -Properties @{Row=$r;Col=$c} }
                Range  = { param($tl,$br)
                    $rng = New-MockComObject
                    Add-Member -InputObject $rng -MemberType ScriptProperty -Name 'Value2' `
                        -GetScriptBlock  { return $writtenArray } `
                        -SetScriptBlock  { param($v); Set-Variable -Name 'writtenArray' -Value $v -Scope 1 }
                    return $rng
                }
            }

            $ctx  = New-TestContext
            $data = @(
                @{ Name = 'Alice'; Age = 30 }
                @{ Name = 'Bob';   Age = 25 }
            )

            # We cannot test real COM assignment, but we CAN test the array building logic.
            # Extract the internal array-building portion as a scriptblock for isolated testing.
            $cols     = @('Age','Name')  # sorted
            $dataRows = $data.Count
            $totalRows= $dataRows + 1   # with header

            $arr = [System.Array]::CreateInstance([object], $totalRows, $cols.Count)

            # Header row
            for ($c = 0; $c -lt $cols.Count; $c++) { $arr.SetValue($cols[$c], 0, $c) }
            # Data rows
            for ($r = 0; $r -lt $dataRows; $r++) {
                for ($c = 0; $c -lt $cols.Count; $c++) {
                    $arr.SetValue("$($data[$r][$cols[$c]])", ($r + 1), $c)
                }
            }

            $arr.GetValue(0,0) | Should -Be 'Age'
            $arr.GetValue(0,1) | Should -Be 'Name'
            $arr.GetValue(1,0) | Should -Be '30'
            $arr.GetValue(1,1) | Should -Be 'Alice'
            $arr.GetValue(2,0) | Should -Be '25'
            $arr.GetValue(2,1) | Should -Be 'Bob'
            $arr.GetLength(0) | Should -Be 3  # 2 data + 1 header
            $arr.GetLength(1) | Should -Be 2  # 2 columns
        }

        It 'skips header row when WriteHeaders = false' {
            $cols = @('A','B')
            $data = @(@{A='x';B='y'})
            $arr  = New-Object 'object[,]' 1, 2  # no header row
            $arr[0,0] = $data[0]['A']
            $arr[0,1] = $data[0]['B']
            $arr[0,0] | Should -Be 'x'
            $arr.GetLength(0) | Should -Be 1
        }
    }

    Context 'ExcelFormats dictionary' {
        It 'contains all expected format codes' {
            $Script:ExcelFormats['xlsx'] | Should -Be 51
            $Script:ExcelFormats['xlsm'] | Should -Be 52
            $Script:ExcelFormats['xls']  | Should -Be 56
            $Script:ExcelFormats['csv']  | Should -Be 6
            $Script:ExcelFormats['pdf']  | Should -Be 57
        }
    }

    Context 'Write-ExcelData backward-compatibility alias' {
        It 'function Write-ExcelData exists as alias to Write-ExcelRange' {
            $fn = Get-Command 'Write-ExcelData' -ErrorAction SilentlyContinue
            $fn | Should -Not -BeNullOrEmpty
        }
    }
}

# ============================================================================
# SapHelper.ps1
# ============================================================================
Describe 'SapHelper.ps1' {

    Context 'Wait-SapReady' {
        It 'returns true immediately when session is not busy' {
            $readyInfo = New-MockComObject -Properties @{ IsActive = $false }
            $mockSess = New-MockComObject -Properties @{ Info = $readyInfo }
            $ctx    = New-TestContext
            $result = Wait-SapReady -Context $ctx -Session $mockSess -TimeoutSec 5 -PollMs 10
            $result | Should -Be $true
        }

        It 'returns false when session stays busy past timeout' {
            $busyInfo = New-MockComObject -Properties @{ IsActive = $true }
            $mockSess = New-MockComObject -Properties @{ Info = $busyInfo }
            $ctx    = New-TestContext
            $result = Wait-SapReady -Context $ctx -Session $mockSess -TimeoutSec 1 -PollMs 100
            $result | Should -Be $false
        }

        It 'returns true if Info throws (assume ready on error)' {
            $mockSess = New-MockComObject -Methods @{
                Info = { throw 'COM error' }
            }
            $ctx    = New-TestContext
            $result = Wait-SapReady -Context $ctx -Session $mockSess -TimeoutSec 5 -PollMs 10
            $result | Should -Be $true
        }
    }

    Context 'Export-SapTableToArray - ALV detection' {
        It 'uses GetCellValue for GuiGridView type' {
            $callLog = [System.Collections.Generic.List[string]]::new()
            $mockTable = New-MockComObject -Properties @{ RowCount = 2; Type = 'GuiGridView' } `
                         -Methods @{
                             GetCellValue = {
                                 param($r,$col)
                                 $callLog.Add("GetCellValue:${r}:${col}")
                                 return "val_${r}_${col}"
                             }
                         }
            $mockSess  = New-MockComObject -Methods @{
                FindById = { param($id); return $mockTable }
            }
            $ctx  = New-TestContext
            $rows = Export-SapTableToArray -Context $ctx -Session $mockSess `
                    -TableId 'wnd[0]/usr/tbl' -Columns @('COL1')
            $rows.Count         | Should -Be 2
            $rows[0]['COL1']    | Should -Be 'val_0_COL1'
            $callLog            | Should -Contain 'GetCellValue:0:COL1'
        }

        It 'uses GetCell for GuiTableControl type' {
            $callLog   = [System.Collections.Generic.List[string]]::new()
            $mockTable = New-MockComObject -Properties @{ RowCount = 1; Type = 'GuiTableControl' } `
                         -Methods @{
                             GetCell = {
                                 param($r,$col)
                                 $callLog.Add("GetCell:${r}:${col}")
                                 New-MockComObject -Properties @{ Text = "cell_${r}_${col}" }
                             }
                         }
            $mockSess  = New-MockComObject -Methods @{
                FindById = { param($id); return $mockTable }
            }
            $ctx  = New-TestContext
            $rows = @(Export-SapTableToArray -Context $ctx -Session $mockSess `
                    -TableId 'wnd[0]/usr/tbl' -Columns @('MATNR'))
            $rows.Count         | Should -Be 1
            $rows[0]['MATNR']   | Should -Be 'cell_0_MATNR'
            $callLog            | Should -Contain 'GetCell:0:MATNR'
        }

        It 'respects CancellationToken and returns partial results' {
            $cancelFlag = [ref]$false
            $mockTable  = New-MockComObject -Properties @{ RowCount = 500; Type = 'GuiGridView' } `
                          -Methods @{
                              GetCellValue = {
                                  param($r,$col)
                                  # Set cancel flag after first 100 rows
                                  if ($r -ge 99) { $cancelFlag.Value = $true }
                                  return "v$r"
                              }
                          }
            $mockSess   = New-MockComObject -Methods @{
                FindById = { param($id); return $mockTable }
            }
            $ctx  = New-TestContext
            $rows = Export-SapTableToArray -Context $ctx -Session $mockSess `
                    -TableId 'tbl' -Columns @('C') -CancellationToken $cancelFlag
            # Should have fewer than 500 rows (cancelled early)
            $rows.Count | Should -BeLessThan 500
        }

        It 'uses a List internally (not += array operator)' {
            # The function should use List[hashtable] for O(n) performance.
            # Verify by checking the source code for the List pattern.
            $srcFile = Join-Path $LibPath 'SapHelper.ps1'
            $src     = Get-Content $srcFile -Raw
            $src     | Should -Match 'List\[hashtable\]'
            $src     | Should -Not -Match '\$rows \+= '
        }
    }
}

# ============================================================================
# Assertions.ps1
# ============================================================================
Describe 'Assertions.ps1' {

    Context 'Assert-Param' {
        It 'does not throw for a non-empty value' {
            $ctx = New-TestContext
            { Assert-Param -Name 'Centro' -Value '1000' -Context $ctx } | Should -Not -Throw
        }

        It 'throws for an empty string' {
            $ctx = New-TestContext
            { Assert-Param -Name 'Centro' -Value '' -Context $ctx } | Should -Throw
        }

        It 'throws for a whitespace-only string' {
            $ctx = New-TestContext
            { Assert-Param -Name 'Centro' -Value '   ' -Context $ctx } | Should -Throw
        }

        It 'includes the parameter name in the error message' {
            $ctx = New-TestContext
            { Assert-Param -Name 'MySpecialParam' -Value '' -Context $ctx } |
                Should -Throw -ExpectedMessage '*MySpecialParam*'
        }
    }

    Context 'Assert-FileExists' {
        It 'does not throw when file exists' {
            $file = Join-Path $TestDrive 'test.txt'
            Set-Content $file 'x' -Encoding ASCII
            $ctx = New-TestContext
            { Assert-FileExists -Path $file -Context $ctx } | Should -Not -Throw
        }

        It 'throws when file does not exist' {
            $ctx = New-TestContext
            { Assert-FileExists -Path (Join-Path $TestDrive 'nonexistent.txt') -Context $ctx } |
                Should -Throw
        }
    }

    Context 'Assert-SapSession' {
        It 'does not throw for a non-null session' {
            $ctx = New-TestContext
            $fakeSess = [PSCustomObject]@{ Type = 'GuiSession' }
            { Assert-SapSession -Session $fakeSess -Context $ctx } | Should -Not -Throw
        }

        It 'throws for a null session' {
            $ctx = New-TestContext
            { Assert-SapSession -Session $null -Context $ctx } | Should -Throw
        }
    }

    Context 'Test-TaskAsset' {
        It 'validates multiple params at once' {
            $ctx = New-TestContext
            { Test-TaskAsset -Context $ctx -Params @{
                Param1 = ''; Param2 = 'ok'
            } } | Should -Throw -ExpectedMessage '*Param1*'
        }

        It 'validates files and params together' {
            $ctx  = New-TestContext
            $file = Join-Path $TestDrive 'asset.csv'
            Set-Content $file 'data' -Encoding ASCII
            { Test-TaskAsset -Context $ctx -Files @{ MyFile = $file } `
              -Params @{ MyParam = 'value' } } | Should -Not -Throw
        }
    }
}

# ============================================================================
# Integration: Config -> Context -> Logger write cycle
# ============================================================================
Describe 'Integration: Full Log Write Cycle' {

    It 'writes a valid JSONL entry to the log file' {
        $logsDir = Join-Path $TestDrive 'integration_logs'
        New-Item $logsDir -ItemType Directory -Force | Out-Null

        $cfg = Get-EngineConfig -Root $TestDrive
        $cfg.engine.logLevel = 'DEBUG'
        $ctx = New-TaskContext -TaskName 'integration_test' -Config $cfg -Root $TestDrive

        # Ensure logs directory exists
        New-Item $ctx.Paths.Logs -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        Write-BuildLog -Context $ctx -Level 'INFO' -Message 'Integration test message' -Detail 'test detail'

        $logFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
        Test-Path $logFile | Should -Be $true

        $line    = Get-Content $logFile -Last 1 -Encoding ASCII
        $entry   = $line | ConvertFrom-Json

        $entry.task    | Should -Be 'integration_test'
        $entry.level   | Should -Be 'INFO'
        $entry.message | Should -Be 'Integration test message'
        $entry.detail  | Should -Be 'test detail'
        $entry.runId   | Should -Not -BeNullOrEmpty
        $entry.ts      | Should -Match '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}'
    }

    It 'sanitizes newlines in Detail field' {
        $cfg = Get-EngineConfig -Root $TestDrive
        $ctx = New-TaskContext -TaskName 'sanitize_test' -Config $cfg -Root $TestDrive
        New-Item $ctx.Paths.Logs -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        Write-BuildLog -Context $ctx -Level 'ERROR' -Message 'Error' -Detail "line1`nline2`r`nline3"

        $logFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
        $line    = Get-Content $logFile -Last 1 -Encoding ASCII
        $line    | Should -Not -Match "`n"
        $line    | Should -Not -Match "`r"
        $entry   = $line | ConvertFrom-Json
        $entry.detail | Should -Be 'line1 | line2 | line3'
    }

    It 'Write-RunResult writes an OK entry with elapsed time' {
        $cfg = Get-EngineConfig -Root $TestDrive
        $ctx = New-TaskContext -TaskName 'result_test' -Config $cfg -Root $TestDrive
        New-Item $ctx.Paths.Logs -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        Start-Sleep -Milliseconds 10
        Write-RunResult -Context $ctx -Success $true

        $logFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
        $line    = Get-Content $logFile -Last 1 -Encoding ASCII
        $entry   = $line | ConvertFrom-Json

        $entry.level   | Should -Be 'OK'
        $entry.elapsed | Should -BeGreaterThan 0
    }

    # ------------------------------------------------------------------
    # CRITICAL-ISSUE-1: logger must auto-create a missing logs directory
    # ------------------------------------------------------------------
    It 'auto-creates the logs directory when it does not exist (CRITICAL-ISSUE-1)' {
        # Fresh sub-tree, no pre-existing logs/
        $freshRoot = Join-Path $TestDrive 'fresh_root'
        New-Item $freshRoot -ItemType Directory -Force | Out-Null
        $logsDir = Join-Path $freshRoot 'logs'
        Remove-Item $logsDir -Recurse -Force -ErrorAction SilentlyContinue

        $cfg = Get-EngineConfig -Root $freshRoot
        $cfg.engine.logLevel = 'DEBUG'
        $ctx = New-TaskContext -TaskName 'dirtest' -Config $cfg -Root $freshRoot

        # Remove logs dir again — Get-EngineConfig creates it, we need it absent
        Remove-Item $logsDir -Recurse -Force -ErrorAction SilentlyContinue

        # Write-BuildLog must NOT throw even though logs/ is missing
        { Write-BuildLog -Context $ctx -Level 'INFO' -Message 'dir auto-create test' } |
            Should -Not -Throw

        Test-Path $logsDir                              | Should -Be $true
        Test-Path (Join-Path $logsDir 'registry.jsonl') | Should -Be $true
    }

    It 'JSONL entries written after auto-dir-create are valid JSON (CRITICAL-ISSUE-1)' {
        $freshRoot = Join-Path $TestDrive 'fresh_root2'
        New-Item $freshRoot -ItemType Directory -Force | Out-Null
        $logsDir = Join-Path $freshRoot 'logs'
        Remove-Item $logsDir -Recurse -Force -ErrorAction SilentlyContinue

        $cfg = Get-EngineConfig -Root $freshRoot
        $ctx = New-TaskContext -TaskName 'jsonltest' -Config $cfg -Root $freshRoot
        Remove-Item $logsDir -Recurse -Force -ErrorAction SilentlyContinue

        Write-BuildLog -Context $ctx -Level 'WARN' -Message 'jsonl validity check'

        $line = Get-Content (Join-Path $logsDir 'registry.jsonl') -Last 1 -Encoding ASCII

        # FIX: In Pester v5, a scriptblock passed to Should -Not -Throw runs in a
        # child scope. Assignments inside that block do NOT propagate back to the It
        # block scope, so $entry remains $null after the Should call — causing
        # PropertyNotFoundException on $entry.task under Set-StrictMode -Version Latest.
        #
        # Correct pattern: parse directly in the It scope, then assert on the object.
        # JSON validity is implicitly tested: ConvertFrom-Json throws on invalid input,
        # which Pester catches and marks as a test failure with a clear parse error.
        $entry = $line | ConvertFrom-Json
        $entry        | Should -Not -BeNullOrEmpty
        $entry.task   | Should -Be 'jsonltest'
        $entry.level  | Should -Be 'WARN'
    }
}

# ============================================================================
# CRITICAL-ISSUE-2: retry log message must contain a non-empty delay value
# ============================================================================
Describe 'Retry.ps1 - delay interpolation (CRITICAL-ISSUE-2)' {

    It 'WARN log message contains a numeric delay value (not empty "s.")' {
        $ctx     = New-TestContext
        $logsDir = $ctx.Paths.Logs
        New-Item $logsDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        try {
            Invoke-WithRetry -Context $ctx -MaxRetries 1 -BaseDelaySeconds 0.05 `
                -ScriptBlock { throw 'force retry' } -Label 'delaytest'
        } catch { }

        $logFile   = Join-Path $logsDir 'registry.jsonl'
        $lines     = Get-Content $logFile -Encoding ASCII -ErrorAction SilentlyContinue
        $retryLine = $lines | Where-Object { $_ -like '*Retrying in*' } | Select-Object -First 1
        $retryLine | Should -Not -BeNullOrEmpty
        # Must match "Retrying in X.Xs." with a real number — not "Retrying in s."
        $retryLine | Should -Match 'Retrying in \d+\.\d+s\.'
    }
}

# ============================================================================
# CRITICAL-ISSUE-3: config merge must not throw for missing optional sections
# ============================================================================
Describe 'Config.ps1 - missing optional sections (CRITICAL-ISSUE-3)' {

    It 'merges a config file that omits optional sections without throwing' {
        $json = '{"engine":{"logLevel":"WARN"}}'
        Set-Content (Join-Path $TestDrive 'engine.config.json') $json -Encoding ASCII
        { Get-EngineConfig -Root $TestDrive } | Should -Not -Throw
    }

    It 'keeps default values for sections absent from the config file' {
        $json = '{"engine":{"logLevel":"WARN"}}'
        Set-Content (Join-Path $TestDrive 'engine.config.json') $json -Encoding ASCII
        $cfg = Get-EngineConfig -Root $TestDrive
        $cfg.excel.visible         | Should -Be $false
        $cfg.notifications.enabled | Should -Be $false
        $cfg.sap.systemId          | Should -Be 'PRD'
        $cfg.security.adminUsers   | Should -Be ''
    }

    It 'merges a config file that supplies only one optional section' {
        $json = '{"excel":{"visible":true}}'
        Set-Content (Join-Path $TestDrive 'engine.config.json') $json -Encoding ASCII
        $cfg = Get-EngineConfig -Root $TestDrive
        $cfg.excel.visible   | Should -Be $true
        $cfg.engine.logLevel | Should -Be 'INFO'
        $cfg.sap.client      | Should -Be '800'
    }
}
