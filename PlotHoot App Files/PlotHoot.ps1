param(
    [switch]$SelfTest,
    [switch]$UiSmokeTest,
    [switch]$NoUi,
    [switch]$RepairShortcuts,
    [string]$Database = ""
)

# PlotHoot
# Offline QA field data checker for BIA CFI Access databases.

$ErrorActionPreference = "Stop"

$script:AppName = "PlotHoot"
$script:AppVersion = "1.0.03"
$script:AppUserModelId = "USDI.BIA.PlotHoot"
$portableAppRoot = [Environment]::GetEnvironmentVariable("PLOTHOOT_APPROOT")
$script:AppRoot = if (-not [string]::IsNullOrWhiteSpace($portableAppRoot)) {
    $portableAppRoot
}
elseif (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Split-Path -Parent $PSCommandPath
}
else {
    (Get-Location).Path
}
$script:AssetRoot = Join-Path $script:AppRoot "Assets"
$script:DataRoot = Join-Path $script:AppRoot "Data"
$script:AppIconPath = Join-Path $script:AssetRoot "PlotHoot.ico"
$script:OwlImagePath = Join-Path $script:AssetRoot "PlotHoot-owl.png"
$script:SettingsPath = Join-Path $script:DataRoot "PlotHootSettings.json"
if (-not (Test-Path -LiteralPath $script:AppIconPath)) {
    $fallbackIcon = Get-ChildItem -LiteralPath $script:AssetRoot -Filter "*.ico" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $fallbackIcon) { $script:AppIconPath = $fallbackIcon.FullName }
}
if (-not (Test-Path -LiteralPath $script:OwlImagePath)) {
    $fallbackImage = Get-ChildItem -LiteralPath $script:AssetRoot -Filter "*.png" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $fallbackImage) { $script:OwlImagePath = $fallbackImage.FullName }
}

$script:DatabasePath = ""
$script:ConnectionString = ""
$script:ProjectName = ""
$script:PlotListMode = "Completed"
$script:InventoryTotalPlotCount = 0
$script:FieldCatalog = @{
    Plot = @()
    Tree = @()
    Regen = @()
}
$script:Plots = @()
$script:CurrentPlot = $null
$script:CurrentSessionId = ""
$script:QaSessions = New-Object System.Collections.Generic.List[object]
$script:Tolerances = @{}
$script:FieldOrder = @{}
$script:ExcludedFieldKeys = @{}
$script:DatabaseFieldPoints = @{}
$script:AutoAdvanceEditingControlHandles = @{}
$script:SuppressGridEvents = $false
$script:OverallRefreshTimer = $null
$script:TreeProgressRefreshTimer = $null
$script:RegenProgressRefreshTimer = $null
$script:DeferredSettingsSaveTimer = $null
$script:StemScoringRefreshTimer = $null
$script:MissedTreeRefreshTimer = $null
$script:AutoSaveNoticeHideTimer = $null
$script:TouchKeyboardEnabled = -not ($SelfTest -or $UiSmokeTest -or $NoUi)
$script:LastTouchKeyboardRequest = [datetime]::MinValue
$script:SplashEnabled = -not ($SelfTest -or $UiSmokeTest -or $NoUi -or $RepairShortcuts)
$script:SettingsDragSourceIndex = -1
$script:SettingsDragStartPoint = $null
$script:DataEntryDragSourceGrid = $null
$script:DataEntryDragSourceIndex = -1
$script:DataEntryDragStartPoint = $null
$script:SuppressStemScoringToggleEvent = $false
$script:SuppressMissedTreeCheckEvent = $false
$script:DefaultStemToleranceBands = @(
    [pscustomobject]@{ Key = "100plus"; Label = "100+ stems"; Min = 100.0; Max = $null; Tolerance = "" },
    [pscustomobject]@{ Key = "51to99"; Label = "51-99 stems"; Min = 51.0; Max = 99.0; Tolerance = "" },
    [pscustomobject]@{ Key = "30to50"; Label = "30-50 stems"; Min = 30.0; Max = 50.0; Tolerance = "" },
    [pscustomobject]@{ Key = "11to29"; Label = "11-29 stems"; Min = 11.0; Max = 29.0; Tolerance = "" },
    [pscustomobject]@{ Key = "2to10"; Label = "2-10 stems"; Min = 2.0; Max = 10.0; Tolerance = "" },
    [pscustomobject]@{ Key = "0to1"; Label = "0-1 stems"; Min = 0.0; Max = 1.0; Tolerance = "" }
)
$script:StemToleranceBands = @()
$script:DefaultUseStemCountPercentageForScoring = $false
$script:UseStemCountPercentageForScoring = $script:DefaultUseStemCountPercentageForScoring
$script:DefaultScorePassPercent = "90"
$script:ScorePassPercent = $script:DefaultScorePassPercent
$script:DefaultMaxPointLoss = ""
$script:MaxPointLoss = $script:DefaultMaxPointLoss
$script:DefaultPlotPointTotal = ""
$script:DefaultTreePointTotal = ""
$script:DefaultRegenPointTotal = ""
$script:DefaultExecutionPointTotal = ""
$script:DefaultInventoryQaTargetPercent = 10.0
$script:InventoryQaTargetPercent = $script:DefaultInventoryQaTargetPercent
$script:PlotPointTotal = $script:DefaultPlotPointTotal
$script:TreePointTotal = $script:DefaultTreePointTotal
$script:RegenPointTotal = $script:DefaultRegenPointTotal
$script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
$script:Ui = @{}

function Set-HooterLaunchShortcut {
    param([string]$ShortcutPath)

    $scriptPath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath) -and (Test-Path -LiteralPath $PSCommandPath)) {
        $PSCommandPath
    }
    else {
        Join-Path $script:AppRoot "PlotHoot.ps1"
    }
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        $fallbackScript = Get-ChildItem -LiteralPath $script:AppRoot -Filter "Plot*.ps1" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $fallbackScript) { $scriptPath = $fallbackScript.FullName }
    }
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        return [pscustomobject]@{ Success = $false; Path = $ShortcutPath; Message = "Could not find the PlotHoot app file." }
    }

    $systemRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
    if ([string]::IsNullOrWhiteSpace($systemRoot)) { $systemRoot = $env:SystemRoot }
    $psPath = Join-Path $systemRoot "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $psPath)) {
        $psPath = Join-Path $systemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    }
    if (-not (Test-Path -LiteralPath $psPath)) {
        return [pscustomobject]@{ Success = $false; Path = $ShortcutPath; Message = "Could not find Windows PowerShell." }
    }

    try {
        $shortcutFolder = Split-Path -Parent $ShortcutPath
        if (-not [string]::IsNullOrWhiteSpace($shortcutFolder) -and -not (Test-Path -LiteralPath $shortcutFolder)) {
            [void](New-Item -ItemType Directory -Force -Path $shortcutFolder)
        }

        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $psPath
        $scriptLiteral = $scriptPath.Replace("'", "''")
        $appRootLiteral = $script:AppRoot.Replace("'", "''")
        $shortcut.Arguments = ('-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command "Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $env:PLOTHOOT_APPROOT=''{0}''; $p=''{1}''; $s=[System.IO.File]::ReadAllText($p); . ([scriptblock]::Create($s)) @args"' -f $appRootLiteral, $scriptLiteral)
        $shortcut.WorkingDirectory = $script:AppRoot
        $iconPath = $script:AppIconPath
        if (-not (Test-Path -LiteralPath $iconPath)) {
            $fallbackIcon = Get-ChildItem -LiteralPath $script:AssetRoot -Filter "*.ico" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($null -ne $fallbackIcon) { $iconPath = $fallbackIcon.FullName }
        }
        $shortcut.IconLocation = if (Test-Path -LiteralPath $iconPath) { "$iconPath,0" } else { "$psPath,0" }
        $shortcut.Description = "PlotHoot offline field QA app"
        $shortcut.Save()
        return [pscustomobject]@{ Success = $true; Path = $ShortcutPath; Message = "Shortcut updated." }
    }
    catch {
        return [pscustomobject]@{ Success = $false; Path = $ShortcutPath; Message = $_.Exception.Message }
    }
}

function Repair-HooterLaunchShortcuts {
    param(
        [switch]$CreateDesktop,
        [switch]$Quiet
    )

    $results = @()
    $rootFolder = Split-Path -Parent $script:AppRoot
    if (-not [string]::IsNullOrWhiteSpace($rootFolder)) {
        $results += Set-HooterLaunchShortcut -ShortcutPath (Join-Path $rootFolder "PlotHoot.lnk")
    }

    if ($CreateDesktop) {
        $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
        if ([string]::IsNullOrWhiteSpace($desktop)) {
            $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
        }
        if (-not [string]::IsNullOrWhiteSpace($desktop)) {
            $results += Set-HooterLaunchShortcut -ShortcutPath (Join-Path $desktop "PlotHoot.lnk")
        }
    }

    if (-not $Quiet) {
        foreach ($result in $results) {
            $status = if ($result.Success) { "OK" } else { "FAILED" }
            "{0}: {1} - {2}" -f $status, $result.Path, $result.Message
        }
        return
    }

    return
}

function ConvertTo-HooterText {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToString("yyyy-MM-dd") }
    if ($Value -is [bool]) { if ($Value) { return "Yes" } else { return "No" } }
    $text = [string]$Value
    $text = $text.Replace([string][char]0xFEFF, "")
    $text = $text.Replace([string][char]0x200B, "")
    $text = $text.Replace([string][char]0x200C, "")
    $text = $text.Replace([string][char]0x200D, "")
    $text = $text.Replace([string][char]0x00A0, " ")
    $text = [regex]::Replace($text, "\s+", " ")
    return $text.Trim()
}

function ConvertTo-HooterBool {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $false }
    if ($Value -is [bool]) { return [bool]$Value }
    if ($Value -is [int32] -or $Value -is [int64] -or $Value -is [int16] -or $Value -is [byte]) { return ([int64]$Value -ne 0) }
    $text = (ConvertTo-HooterText $Value).ToLowerInvariant()
    return ($text -in @("1", "true", "yes", "y", "checked", "active", "on"))
}

function Get-HooterDatabaseDisplayName {
    param([string]$Path)

    $clean = ConvertTo-HooterText $Path
    if ([string]::IsNullOrWhiteSpace($clean)) { return "" }
    try {
        $name = [System.IO.Path]::GetFileName($clean)
        if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
    }
    catch {}
    return $clean
}

function Set-HooterDatabaseBoxPath {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [string]$Path
    )

    if ($null -eq $TextBox) { return }
    $clean = ConvertTo-HooterText $Path
    $TextBox.Tag = $clean
    $TextBox.Text = Get-HooterDatabaseDisplayName -Path $clean
}

function Get-HooterDatabaseBoxPath {
    $box = if ($script:Ui.ContainsKey("DatabaseBox")) { $script:Ui.DatabaseBox } else { $null }
    if ($null -eq $box) { return "" }
    $text = ConvertTo-HooterText $box.Text
    $tag = ConvertTo-HooterText $box.Tag
    if (-not [string]::IsNullOrWhiteSpace($tag)) {
        $tagName = Get-HooterDatabaseDisplayName -Path $tag
        if ([string]::IsNullOrWhiteSpace($text) -or
            $text.Equals($tag, [System.StringComparison]::OrdinalIgnoreCase) -or
            $text.Equals($tagName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $tag
        }
    }
    return $text
}

function Get-NameKey {
    param([object]$Value)

    $text = (ConvertTo-HooterText $Value).ToLowerInvariant()
    return ($text -replace "[^a-z0-9]", "")
}

function Quote-Name {
    param([string]$Name)
    return "[" + $Name.Replace("]", "]]") + "]"
}

function Sql-Text {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "Null" }
    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "Null" }
    return "'" + $text.Replace("'", "''") + "'"
}

function ConvertTo-HooterGuidText {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $match = [regex]::Match($text, "[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")
    if (-not $match.Success) { return "" }
    return $match.Value
}

function Sql-Guid {
    param([object]$Value)

    $guidText = ConvertTo-HooterGuidText $Value
    if ([string]::IsNullOrWhiteSpace($guidText)) {
        throw "Expected a database ID value in GUID format, but found '$(ConvertTo-HooterText $Value)'."
    }
    return "{guid {" + $guidText + "}}"
}

function Field-Text-Expression {
    param([string]$FieldName)
    return "(" + (Quote-Name $FieldName) + " & '')"
}

function Escape-Html {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode((ConvertTo-HooterText $Value))
}

function New-InsensitiveHashtable {
    return New-Object "System.Collections.Hashtable" ([System.StringComparer]::OrdinalIgnoreCase)
}

function Get-ProviderCandidates {
    param([string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq ".mdb") {
        return @(
            "Microsoft.Jet.OLEDB.4.0",
            "Microsoft.ACE.OLEDB.16.0",
            "Microsoft.ACE.OLEDB.15.0",
            "Microsoft.ACE.OLEDB.12.0"
        )
    }

    return @(
        "Microsoft.ACE.OLEDB.16.0",
        "Microsoft.ACE.OLEDB.15.0",
        "Microsoft.ACE.OLEDB.12.0",
        "Microsoft.Jet.OLEDB.4.0"
    )
}

function New-AccessConnectionString {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Database not found: $Path"
    }

    $lastError = ""
    foreach ($provider in (Get-ProviderCandidates -Path $Path)) {
        $connectionString = "Provider=$provider;Data Source=$Path;Mode=Read;Persist Security Info=False;"
        $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
        try {
            $connection.Open()
            $connection.Close()
            return $connectionString
        }
        catch {
            $lastError = $_.Exception.Message
        }
        finally {
            $connection.Dispose()
        }
    }

    $bitness = if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" }
    throw "Could not open the Access database from this $bitness process. Use the 32-bit PlotHoot launcher and make sure the matching Microsoft Access Database Engine is installed. Last driver error: $lastError"
}

function Open-AccessConnection {
    param([string]$ConnectionString)

    $connection = New-Object System.Data.OleDb.OleDbConnection($ConnectionString)
    $connection.Open()
    return $connection
}

function Get-DbTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
    $table = New-Object System.Data.DataTable
    try {
        [void]$adapter.Fill($table)
        return ,$table
    }
    finally {
        $adapter.Dispose()
        $command.Dispose()
    }
}

function Get-UserTables {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $schema = $Connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($row in $schema.Rows) {
        $name = [string]$row["TABLE_NAME"]
        $type = [string]$row["TABLE_TYPE"]
        if ($type -eq "TABLE" -and
            -not $name.StartsWith("MSys", [System.StringComparison]::OrdinalIgnoreCase) -and
            -not $name.StartsWith("~", [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$names.Add($name)
        }
    }
    return @($names.ToArray() | Sort-Object)
}

function Get-TableColumns {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $table = Get-DbTable -Connection $Connection -Sql "SELECT * FROM $(Quote-Name $TableName) WHERE 1=0"
    $columns = New-Object System.Collections.Generic.List[string]
    foreach ($column in $table.Columns) {
        [void]$columns.Add([string]$column.ColumnName)
    }
    return [string[]]$columns.ToArray()
}

function Get-HooterQueryColumnOrder {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$QueryNames
    )

    foreach ($queryName in @($QueryNames)) {
        if ([string]::IsNullOrWhiteSpace($queryName)) { continue }
        $order = @{}
        try {
            $table = Get-DbTable -Connection $Connection -Sql "SELECT * FROM $(Quote-Name $queryName) WHERE 1=0"
            for ($index = 0; $index -lt $table.Columns.Count; $index++) {
                $columnName = [string]$table.Columns[$index].ColumnName
                $key = Get-NameKey $columnName
                if (-not [string]::IsNullOrWhiteSpace($key) -and -not $order.ContainsKey($key)) {
                    $order[$key] = $index
                }
            }
        }
        catch {
            try {
                $schema = $Connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns, $null)
                $rows = @($schema.Rows | Where-Object {
                    ([string]$_["TABLE_NAME"]).Equals($queryName, [System.StringComparison]::OrdinalIgnoreCase)
                } | Sort-Object @{
                    Expression = {
                        $ordinal = 0
                        try { $ordinal = [int]$_["ORDINAL_POSITION"] } catch { $ordinal = 0 }
                        $ordinal
                    }
                })
                for ($index = 0; $index -lt $rows.Count; $index++) {
                    $columnName = [string]$rows[$index]["COLUMN_NAME"]
                    $key = Get-NameKey $columnName
                    if (-not [string]::IsNullOrWhiteSpace($key) -and -not $order.ContainsKey($key)) {
                        $order[$key] = $index
                    }
                }
            }
            catch {
                $order = @{}
            }
        }
        if ($order.Count -gt 0) {
            return [pscustomobject]@{
                QueryName = $queryName
                Order = $order
            }
        }
    }

    return [pscustomobject]@{
        QueryName = ""
        Order = @{}
    }
}

function Get-HooterMeasurementQueryOrders {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    return @{
        Plot = Get-HooterQueryColumnOrder -Connection $Connection -QueryNames @(
            "GetPlotMeasurementsForPeriod",
            "GetPlotMeasuremetsForPeriod",
            "getplotmeasuremetsforperiod",
            "getplotmeasurementsforperiod"
        )
        Tree = Get-HooterQueryColumnOrder -Connection $Connection -QueryNames @(
            "GetTreeMeasurementsForPeriodByKey",
            "GetTreeMeasurementForPeriodByKey",
            "gettreemeasurementsforperiodbykey"
        )
        Regen = Get-HooterQueryColumnOrder -Connection $Connection -QueryNames @(
            "GetRegenMeasurements",
            "getregenmeasurements"
        )
    }
}

function Get-HooterFieldQueryOrder {
    param(
        [object]$Field,
        [hashtable]$QueryOrders
    )

    if ($null -eq $Field -or $null -eq $QueryOrders) { return 100000 }
    $group = ConvertTo-HooterText $Field.Group
    if ([string]::IsNullOrWhiteSpace($group) -or -not $QueryOrders.ContainsKey($group)) { return 100000 }
    $queryOrder = $QueryOrders[$group]
    if ($null -eq $queryOrder -or $null -eq $queryOrder.Order -or $queryOrder.Order.Count -eq 0) { return 100000 }

    foreach ($key in @(Get-HooterFieldQueryCandidateKeys -Field $Field)) {
        if (-not [string]::IsNullOrWhiteSpace($key) -and $queryOrder.Order.ContainsKey($key)) {
            return [int]$queryOrder.Order[$key]
        }
    }

    foreach ($key in @(Get-HooterFieldQueryCandidateKeys -Field $Field)) {
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        foreach ($queryKey in @($queryOrder.Order.Keys)) {
            if ([string]::IsNullOrWhiteSpace($queryKey) -or $queryKey.Length -lt 4) { continue }
            if ($key.EndsWith($queryKey, [System.StringComparison]::OrdinalIgnoreCase) -or $queryKey.EndsWith($key, [System.StringComparison]::OrdinalIgnoreCase)) {
                return [int]$queryOrder.Order[$queryKey]
            }
        }
    }

    return 100000
}

function Get-HooterFieldQueryCandidateKeys {
    param([object]$Field)

    $keys = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $addKey = {
        param([string]$Key)
        $clean = Get-NameKey $Key
        if ([string]::IsNullOrWhiteSpace($clean) -or $seen.ContainsKey($clean)) { return }
        $seen[$clean] = $true
        [void]$keys.Add($clean)
    }

    $fieldName = ConvertTo-HooterText $Field.FieldName
    $label = ConvertTo-HooterText $Field.Label
    $tableName = ConvertTo-HooterText $Field.TableName
    $group = ConvertTo-HooterText $Field.Group
    foreach ($name in @($fieldName, $label, ("{0}.{1}" -f $tableName, $fieldName))) {
        & $addKey $name
        $nameKey = Get-NameKey $name
        foreach ($prefix in @((Get-NameKey $group), (Get-NameKey $tableName), "plot", "tree", "regen", "measurement", "measurements")) {
            if (-not [string]::IsNullOrWhiteSpace($prefix) -and $nameKey.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase) -and $nameKey.Length -gt $prefix.Length) {
                & $addKey $nameKey.Substring($prefix.Length)
            }
        }
        foreach ($suffix in @((Get-NameKey $group), "plot", "tree", "regen")) {
            if (-not [string]::IsNullOrWhiteSpace($suffix) -and $nameKey.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase) -and $nameKey.Length -gt $suffix.Length) {
                & $addKey $nameKey.Substring(0, $nameKey.Length - $suffix.Length)
            }
        }
    }
    return @($keys.ToArray())
}

function Get-ActualTableName {
    param(
        [string[]]$Tables,
        [string]$Name
    )

    foreach ($table in @($Tables)) {
        if ($table.Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $table }
    }
    return ""
}

function Get-ActualColumnName {
    param(
        [string[]]$Columns,
        [string[]]$Names
    )

    foreach ($name in @($Names)) {
        foreach ($column in @($Columns)) {
            if ($column.Equals($name, [System.StringComparison]::OrdinalIgnoreCase)) { return $column }
        }
    }
    return ""
}

function Test-TableExists {
    param(
        [string[]]$Tables,
        [string]$Name
    )

    return -not [string]::IsNullOrWhiteSpace((Get-ActualTableName -Tables $Tables -Name $Name))
}

function Test-ColumnExists {
    param(
        [string[]]$Columns,
        [string]$Name
    )

    return -not [string]::IsNullOrWhiteSpace((Get-ActualColumnName -Columns $Columns -Names @($Name)))
}

function Get-DataRowValue {
    param(
        [System.Data.DataRow]$Row,
        [string[]]$ColumnNames
    )

    if ($null -eq $Row) { return $null }
    foreach ($name in @($ColumnNames)) {
        foreach ($column in $Row.Table.Columns) {
            if ($column.ColumnName.Equals($name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $Row[$column.ColumnName]
            }
        }
    }
    return $null
}

function Get-DataRowText {
    param(
        [System.Data.DataRow]$Row,
        [string[]]$ColumnNames
    )

    return ConvertTo-HooterText (Get-DataRowValue -Row $Row -ColumnNames $ColumnNames)
}

function Normalize-HooterProjectName {
    param(
        [object]$Value,
        [string]$FallbackPath = ""
    )

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text) -and -not [string]::IsNullOrWhiteSpace($FallbackPath)) {
        try { $text = [System.IO.Path]::GetFileNameWithoutExtension($FallbackPath) } catch { $text = $FallbackPath }
    }
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    try {
        if ([System.IO.Path]::HasExtension($text)) {
            $extension = [System.IO.Path]::GetExtension($text)
            if ($extension -match "^\.(mdb|accdb)$") {
                $text = [System.IO.Path]::GetFileNameWithoutExtension($text)
            }
        }
    }
    catch {}

    $text = $text -replace "_+", " "
    do {
        $before = $text
        $text = [regex]::Replace($text, "(?i)\s*\([^)]*(?:\d{4}|v(?:ersion)?\.?\s*\d|database|db)[^)]*\)\s*$", "")
        $text = [regex]::Replace($text, "(?i)\s*[-_ ]+(?:v(?:ersion)?\.?\s*\d+(?:\.\d+)*)\s*$", "")
        $text = [regex]::Replace($text, "\s*[-_ ]+(?:\d{4}[-_ ]?\d{1,2}[-_ ]?\d{1,2}|\d{1,2}[-_ ]\d{1,2}[-_ ]\d{2,4}|\d{8})\s*$", "")
        $text = [regex]::Replace($text, "\s*[-_ ]+(?:19\d{2}|20\d{2})\s*$", "")
        $text = [regex]::Replace($text, "(?i)\s*[-_ ]+(?:database|db|master|copy|backup)\s*$", "")
        $text = $text.Trim(" ", "-", "_", ".")
    } while ($text -ne $before -and -not [string]::IsNullOrWhiteSpace($text))

    $text = [regex]::Replace($text, "\s+", " ").Trim()
    return $text
}

function Get-HooterProjectNameFromDatabase {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$DatabasePath = ""
    )

    $tables = @(Get-UserTables -Connection $Connection)
    $directTables = @(
        "Projects",
        "Project",
        "ProjectInfo",
        "ProjectInformation",
        "ProjectSettings",
        "ProjectSetup",
        "ProjectHeader",
        "ProjectHeaders",
        "ProjectPeriods",
        "ProjectPeriod",
        "DatabaseInfo",
        "DatabaseInformation"
    )
    $nameColumns = @("ProjectName", "Project_Name", "ProjectTitle", "Project", "Name", "Title", "Description")

    foreach ($tableName in $directTables) {
        $actualTable = Get-ActualTableName -Tables $tables -Name $tableName
        if ([string]::IsNullOrWhiteSpace($actualTable)) { continue }
        $columns = @(Get-TableColumns -Connection $Connection -TableName $actualTable)
        foreach ($columnName in $nameColumns) {
            $actualColumn = Get-ActualColumnName -Columns $columns -Names @($columnName)
            if ([string]::IsNullOrWhiteSpace($actualColumn)) { continue }
            try {
                $data = Get-DbTable -Connection $Connection -Sql "SELECT TOP 1 * FROM $(Quote-Name $actualTable) WHERE $(Quote-Name $actualColumn) Is Not Null"
                foreach ($row in $data.Rows) {
                    $candidate = Normalize-HooterProjectName -Value (Get-DataRowText -Row $row -ColumnNames @($actualColumn))
                    if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
                }
            }
            catch {}
        }
    }

    $keyTables = @("Settings", "AppSettings", "ApplicationSettings", "DatabaseSettings", "ProjectSettings", "ProjectInfo")
    foreach ($tableName in $keyTables) {
        $actualTable = Get-ActualTableName -Tables $tables -Name $tableName
        if ([string]::IsNullOrWhiteSpace($actualTable)) { continue }
        $columns = @(Get-TableColumns -Connection $Connection -TableName $actualTable)
        $keyColumn = Get-ActualColumnName -Columns $columns -Names @("Key", "SettingKey", "SettingName", "PropertyName", "Name", "Item")
        $valueColumn = Get-ActualColumnName -Columns $columns -Names @("Value", "SettingValue", "PropertyValue", "TextValue", "Description")
        if ([string]::IsNullOrWhiteSpace($keyColumn) -or [string]::IsNullOrWhiteSpace($valueColumn)) { continue }
        try {
            $data = Get-DbTable -Connection $Connection -Sql "SELECT * FROM $(Quote-Name $actualTable)"
            foreach ($row in $data.Rows) {
                $key = Get-NameKey (Get-DataRowText -Row $row -ColumnNames @($keyColumn))
                if ($key -in @("projectname", "projecttitle", "project")) {
                    $candidate = Normalize-HooterProjectName -Value (Get-DataRowText -Row $row -ColumnNames @($valueColumn))
                    if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
                }
            }
        }
        catch {}
    }

    return Normalize-HooterProjectName -FallbackPath $DatabasePath
}

function Get-HooterCurrentProjectName {
    $name = Normalize-HooterProjectName -Value $script:ProjectName
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = Normalize-HooterProjectName -FallbackPath $script:DatabasePath
    }
    return $name
}

function Test-HooterSessionMatchesCurrentInventory {
    param([object]$Session)

    if ($null -eq $Session) { return $false }
    $currentDatabase = ConvertTo-HooterText $script:DatabasePath
    $sessionDatabase = ConvertTo-HooterText $Session.Database
    if (-not [string]::IsNullOrWhiteSpace($currentDatabase) -and -not [string]::IsNullOrWhiteSpace($sessionDatabase)) {
        if ($sessionDatabase.Equals($currentDatabase, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }

    $currentProject = Get-HooterCurrentProjectName
    $sessionProject = Normalize-HooterProjectName -Value $Session.ProjectName
    if (-not [string]::IsNullOrWhiteSpace($currentProject) -and -not [string]::IsNullOrWhiteSpace($sessionProject)) {
        if ($sessionProject.Equals($currentProject, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }

    return [string]::IsNullOrWhiteSpace($currentDatabase)
}

function Get-HooterInventoryQaTargetPercent {
    $targetPercent = 0.0
    if (-not (ConvertTo-HooterNumber -Value $script:InventoryQaTargetPercent -Number ([ref]$targetPercent))) {
        $targetPercent = $script:DefaultInventoryQaTargetPercent
    }
    if ($targetPercent -lt $script:DefaultInventoryQaTargetPercent) {
        $targetPercent = $script:DefaultInventoryQaTargetPercent
    }
    return [Math]::Max(0.0, $targetPercent)
}

function Get-HooterInventoryQaProgress {
    $plotKeys = @{}
    foreach ($session in @($script:QaSessions.ToArray())) {
        if (-not (Test-HooterSessionMatchesCurrentInventory -Session $session)) { continue }
        $plotNumber = ConvertTo-HooterText $session.PlotNumber
        if ([string]::IsNullOrWhiteSpace($plotNumber)) { continue }
        $plotKeys[(Get-NameKey $plotNumber)] = $true
    }

    $totalPlots = 0
    if ($script:InventoryTotalPlotCount -gt 0) {
        $totalPlots = [int]$script:InventoryTotalPlotCount
    }
    else {
        $totalPlots = @($script:Plots).Count
    }
    $checkedPlots = $plotKeys.Count
    $targetPercent = Get-HooterInventoryQaTargetPercent
    $requiredPlots = if ($totalPlots -gt 0) { [int][Math]::Ceiling($totalPlots * ($targetPercent / 100.0)) } else { 0 }
    $completionPercent = if ($totalPlots -gt 0) { [Math]::Round(($checkedPlots / [double]$totalPlots) * 100.0, 1) } else { 0.0 }
    $remainingPlots = [Math]::Max(0, $requiredPlots - $checkedPlots)
    return [pscustomobject]@{
        TotalPlots = $totalPlots
        CheckedPlots = $checkedPlots
        TargetPercent = $targetPercent
        RequiredPlots = $requiredPlots
        CompletionPercent = $completionPercent
        RemainingPlots = $remainingPlots
        TargetMet = ($requiredPlots -gt 0 -and $checkedPlots -ge $requiredPlots)
    }
}

function Update-HooterInventoryProgress {
    if (-not ($script:Ui.ContainsKey("InventoryProgressLabel") -or $script:Ui.ContainsKey("InventoryProgressBar"))) { return }

    $progress = Get-HooterInventoryQaProgress
    $targetText = Format-HooterScoreNumber $progress.TargetPercent
    if ($progress.TotalPlots -gt 0) {
        $plotWord = if ($progress.RequiredPlots -eq 1) { "plot" } else { "plots" }
        $statusText = if ($progress.TargetMet) { "target met" } else { "$($progress.RemainingPlots) more plot(s) needed" }
        $text = "BIA inventory QA progress: $($progress.CheckedPlots)/$($progress.TotalPlots) plots saved ($($progress.CompletionPercent)% complete). $targetText% target = $($progress.RequiredPlots) $plotWord; $statusText."
    }
    else {
        $text = "BIA inventory QA progress: connect to an inventory database to calculate the $targetText% plot target."
    }

    if ($script:Ui.ContainsKey("InventoryProgressLabel")) {
        $script:Ui.InventoryProgressLabel.Text = $text
        $script:Ui.InventoryProgressLabel.ForeColor = if ($progress.TargetMet) { [System.Drawing.Color]::FromArgb(31, 112, 74) } else { [System.Drawing.Color]::FromArgb(128, 92, 22) }
    }
    if ($script:Ui.ContainsKey("InventoryProgressBar")) {
        $script:Ui.InventoryProgressBar.Maximum = 1000
        $value = if ($progress.RequiredPlots -gt 0) { [int][Math]::Min(1000, [Math]::Round(($progress.CheckedPlots / [double]$progress.RequiredPlots) * 1000.0)) } else { 0 }
        $script:Ui.InventoryProgressBar.Value = [Math]::Max(0, $value)
    }
}

function Update-HooterProjectDisplay {
    $name = Get-HooterCurrentProjectName
    if ($script:Ui.ContainsKey("ProjectLabel")) {
        $script:Ui.ProjectLabel.Text = if ([string]::IsNullOrWhiteSpace($name)) { "Project: Not connected" } else { "Project: $name" }
    }
    if ($script:Ui.ContainsKey("Form")) {
        $script:Ui.Form.Text = if ([string]::IsNullOrWhiteSpace($name)) { "$script:AppName $script:AppVersion" } else { "$script:AppName - $name" }
    }
}

function Get-HooterTableRole {
    param([string]$TableName)

    $key = Get-NameKey $TableName
    if ($key -in @("plots", "plotmeasurements", "plotcustommeasurements")) { return "Plot" }
    if ($key -in @("trees", "treemeasurements", "treecustommeasurements")) { return "Tree" }
    if ($key -in @("regenmeasurements", "regencustommeasurements", "regenerationmeasurements", "regencounts")) { return "Regen" }
    return ""
}

function Get-HooterFieldKey {
    param(
        [string]$Group,
        [string]$TableName,
        [string]$FieldName
    )

    return "$Group|$TableName|$FieldName"
}

function Test-HooterHiddenField {
    param(
        [string]$TableName,
        [string]$FieldName
    )

    $tableKey = Get-NameKey $TableName
    if ($tableKey -match "calculation|summary|assignment") { return $true }

    $fieldKey = Get-NameKey $FieldName
    if ([string]::IsNullOrWhiteSpace($fieldKey)) { return $true }
    if ($fieldKey.StartsWith("calc")) { return $true }

    $hidden = @(
        "id",
        "projectid",
        "reservationid",
        "typeid",
        "tableid",
        "datatypeid",
        "categorytypeid",
        "plotid",
        "treeid",
        "measurementid",
        "retrievalid",
        "projectperiodid",
        "calculationid",
        "summaryid",
        "volumeid",
        "plotkey",
        "treekey",
        "plotmeaskey",
        "treemeaskey",
        "regenmeaskey",
        "created",
        "updated",
        "upduser",
        "isdeleted",
        "gmp",
        "fieldnamesampleremove",
        "lastgrowingseason",
        "peracreexpansion",
        "flccommercial",
        "managementunit",
        "mgmtunit",
        "managementunitcode",
        "managementunitid",
        "plotremarks",
        "plotremark",
        "treeremarks",
        "treeremark",
        "regenremarks",
        "regenremark",
        "remarks",
        "remark",
        "comments",
        "comment"
    )
    return ($hidden -contains $fieldKey)
}

function Normalize-HooterToleranceMode {
    param([object]$Mode)

    $modeText = ConvertTo-HooterText $Mode
    if ([string]::IsNullOrWhiteSpace($modeText)) { return "Exact" }
    if ($modeText.Equals("Absolute", [System.StringComparison]::OrdinalIgnoreCase)) { return "Range" }
    if ($modeText.Equals("Range", [System.StringComparison]::OrdinalIgnoreCase)) { return "Range" }
    foreach ($mode in @("Exact", "Percent", "Class", "StemCount", "StemPercent", "PassFail", "Filled", "GoodFairPoor")) {
        if ($modeText.Equals($mode, [System.StringComparison]::OrdinalIgnoreCase)) { return $mode }
    }
    return $modeText
}

function Test-HooterFieldExcluded {
    param([object]$Field)

    $fieldKey = if ($Field -is [string]) { ConvertTo-HooterText $Field } else { ConvertTo-HooterText $Field.FieldKey }
    return (-not [string]::IsNullOrWhiteSpace($fieldKey) -and $script:ExcludedFieldKeys.ContainsKey($fieldKey))
}

function Test-HooterHiddenFieldKey {
    param([string]$FieldKey)

    $parts = (ConvertTo-HooterText $FieldKey) -split "\|", 3
    if ($parts.Count -lt 3) { return $false }
    return (Test-HooterHiddenField -TableName $parts[1] -FieldName $parts[2])
}

function Test-HooterFieldUseCanBeChanged {
    param([object]$Field)

    $fieldKey = if ($Field -is [string]) { ConvertTo-HooterText $Field } else { ConvertTo-HooterText $Field.FieldKey }
    if ([string]::IsNullOrWhiteSpace($fieldKey)) { return $false }
    if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) { return $false }
    if ($fieldKey.Equals((ConvertTo-HooterText (Get-HooterTreeFoundField).FieldKey), [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    return $true
}

function Get-HooterToleranceUnitHint {
    param(
        [object]$Field = $null,
        [object]$Tolerance = $null
    )

    $mode = if ($null -ne $Tolerance) { Normalize-HooterToleranceMode $Tolerance.Mode } else { "" }
    if ($mode -eq "Percent") { return "%" }
    if ($mode -eq "Class") { return "class" }
    if ($mode -eq "StemPercent") { return "allowance/step" }
    if ($mode -eq "StemCount") { return "stems" }
    if ($mode -eq "GoodFairPoor") { return "point loss" }
    if ($mode -eq "Exact" -or $mode -eq "Filled" -or $mode -eq "PassFail") { return "" }

    $nameKey = ""
    if ($null -ne $Field) {
        $nameKey = Get-NameKey ("$(ConvertTo-HooterText $Field.FieldName) $(ConvertTo-HooterText $Field.Label)")
    }
    elseif ($null -ne $Tolerance) {
        $nameKey = Get-NameKey ("$(ConvertTo-HooterText $Tolerance.FieldName) $(ConvertTo-HooterText $Tolerance.Label)")
    }

    if ($nameKey -match "utm|easting|northing|elevation|distance|azimuth") { return "ft" }
    if ($nameKey -match "dbh|diameter|increment") { return "in" }
    if ($nameKey -match "height") { return "ft" }
    if ($nameKey -match "age") { return "years" }
    if ($nameKey -match "stem|count") { return "stems" }
    return ""
}

function Test-HooterFilledOnlyField {
    param(
        [string]$Group,
        [string]$FieldName,
        [string]$Label = ""
    )

    $keys = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($candidate in @($FieldName, $Label, "$FieldName $Label")) {
        $candidateKey = Get-NameKey $candidate
        if (-not [string]::IsNullOrWhiteSpace($candidateKey) -and -not $seen.ContainsKey($candidateKey)) {
            $seen[$candidateKey] = $true
            [void]$keys.Add($candidateKey)
        }
    }
    foreach ($nameKey in @($keys.ToArray())) {
        if ($nameKey -match "^(crew|crewname|crewid|crewleader|crewlead|fieldcrew|fieldcrewname|cruiser|cruisername|cruiserid)$") { return $true }
        if ($nameKey -match "^(measurementdate|measurementday|measurementdt|measuredate|measureday|measuredat|measdate|measday|measdt|datemeasured|daymeasured|fielddate|fieldday)$") { return $true }
        if ($nameKey -match "(measurement|measure|meas).*(date|day|dt)$") { return $true }
        if ($Group -eq "Plot" -and $nameKey -match "^(plotremarks|plotremark|plotcomments|plotcomment|plotnotes|plotnote|remarks|remark|comments|comment|notes|note)$") { return $true }
        if ($Group -eq "Plot" -and $nameKey -match "^(plotnumber|plotno|plotnum|plot)$") { return $true }
        if ($Group -eq "Tree" -and $nameKey -match "^(treenumber|treeno|treenum)$") { return $true }
    }
    return $false
}

function Test-HooterFilledOnlyFieldKey {
    param([string]$FieldKey)

    $parts = (ConvertTo-HooterText $FieldKey) -split "\|", 3
    if ($parts.Count -lt 3) { return $false }
    return (Test-HooterFilledOnlyField -Group $parts[0] -FieldName $parts[2])
}

function Test-HooterRegenStemCountFieldKey {
    param([string]$FieldKey)

    $parts = (ConvertTo-HooterText $FieldKey) -split "\|", 3
    if ($parts.Count -lt 3) { return $false }
    return (Test-HooterRegenStemCountField -Group $parts[0] -FieldName $parts[2])
}

function Test-HooterHiddenTreeCheckField {
    param([object]$Field)

    if ($null -eq $Field) { return $false }
    $group = ConvertTo-HooterText $Field.Group
    $fieldKey = Get-NameKey ("$(ConvertTo-HooterText $Field.FieldName) $(ConvertTo-HooterText $Field.Label)")
    if ($group -in @("Tree", "Regen") -and $fieldKey -match "plotnumber|plotno|plotnum") { return $true }
    if ($group -eq "Tree" -and $fieldKey -match "realdbh") { return $true }
    return $false
}

function Test-HooterRegenStemCountField {
    param(
        [string]$Group,
        [string]$FieldName,
        [string]$Label = ""
    )

    if (-not (ConvertTo-HooterText $Group).Equals("Regen", [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    $key = Get-NameKey ("$FieldName $Label")
    if ($key -match "stem.*count|count.*stem|stemcnt|stemnumber|numberofstems|totalstems|stemtotal|regenstem") { return $true }
    return ($key -match "^(count|totalcount|regencount|stems)$")
}

function Copy-HooterStemToleranceBand {
    param([object]$Band)

    $min = 0.0
    if ($null -ne $Band.PSObject.Properties["Min"]) {
        [void](ConvertTo-HooterNumber -Value $Band.Min -Number ([ref]$min))
    }

    $max = $null
    if ($null -ne $Band.PSObject.Properties["Max"]) {
        $maxNumber = 0.0
        if (ConvertTo-HooterNumber -Value $Band.Max -Number ([ref]$maxNumber)) {
            $max = $maxNumber
        }
    }

    $tolerance = ""
    if ($null -ne $Band.PSObject.Properties["Tolerance"]) {
        $toleranceNumber = 0.0
        if (ConvertTo-HooterNumber -Value $Band.Tolerance -Number ([ref]$toleranceNumber)) {
            $tolerance = [Math]::Max(0.0, $toleranceNumber)
        }
    }

    return [pscustomobject]@{
        Key = ConvertTo-HooterText $Band.Key
        Label = ConvertTo-HooterText $Band.Label
        Min = [Math]::Max(0.0, $min)
        Max = $max
        Tolerance = $tolerance
    }
}

function Reset-HooterStemToleranceBands {
    $script:StemToleranceBands = @($script:DefaultStemToleranceBands | ForEach-Object { Copy-HooterStemToleranceBand -Band $_ })
}

function Get-HooterStemToleranceBands {
    if (@($script:StemToleranceBands).Count -eq 0) {
        Reset-HooterStemToleranceBands
    }
    return @($script:StemToleranceBands | Sort-Object -Property @{ Expression = { [double]$_.Min }; Descending = $true })
}

function Set-HooterStemToleranceBand {
    param(
        [string]$Key,
        [object]$Tolerance
    )

    $keyText = ConvertTo-HooterText $Key
    if ($keyText.Contains("|")) {
        $keyText = ConvertTo-HooterText (@($keyText -split "\|")[-1])
    }
    if ([string]::IsNullOrWhiteSpace($keyText)) { return $false }

    $toleranceText = ConvertTo-HooterText $Tolerance
    $toleranceValue = ""
    if (-not [string]::IsNullOrWhiteSpace($toleranceText)) {
        $toleranceNumber = 0.0
        if (-not (ConvertTo-HooterNumber -Value $toleranceText -Number ([ref]$toleranceNumber))) { return $false }
        $toleranceValue = [Math]::Max(0.0, $toleranceNumber)
    }

    foreach ($band in @(Get-HooterStemToleranceBands)) {
        if ($keyText.Equals((ConvertTo-HooterText $band.Key), [System.StringComparison]::OrdinalIgnoreCase) -or
            $keyText.Equals((ConvertTo-HooterText $band.Label), [System.StringComparison]::OrdinalIgnoreCase)) {
            $band.Tolerance = $toleranceValue
            return $true
        }
    }
    return $false
}

function Get-HooterRegenStemToleranceBand {
    param([double]$QaCount)

    $count = [Math]::Max(0.0, [Math]::Abs($QaCount))
    foreach ($band in @(Get-HooterStemToleranceBands)) {
        $maxOk = $true
        if ($null -ne $band.Max) {
            $maxOk = ($count -le [double]$band.Max)
        }
        if ($count -ge [double]$band.Min -and $maxOk) {
            return [pscustomobject]@{
                Label = ConvertTo-HooterText $band.Label
                Tolerance = $band.Tolerance
            }
        }
    }
    return [pscustomobject]@{ Label = "0-1 stems"; Tolerance = "" }
}

function Get-HooterStemPercentSettings {
    param([object]$Tolerance)

    $allowance = 10.0
    $step = 5.0
    $value = if ($null -ne $Tolerance -and $null -ne $Tolerance.PSObject.Properties["Value"]) { ConvertTo-HooterText $Tolerance.Value } else { "" }
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        $pairMatches = [regex]::Matches($value, "(?i)\b(allowance|allowed|allow|tolerance|tol|step|increment|every|per)\s*[:=]\s*([+-]?\d+(?:\.\d+)?)")
        foreach ($match in $pairMatches) {
            $number = 0.0
            if (-not (ConvertTo-HooterNumber -Value $match.Groups[2].Value -Number ([ref]$number))) { continue }
            $name = $match.Groups[1].Value.ToLowerInvariant()
            if ($name -in @("allowance", "allowed", "allow", "tolerance", "tol")) {
                $allowance = [Math]::Max(0.0, $number)
            }
            else {
                $step = [Math]::Max(0.000001, $number)
            }
        }
        if ($pairMatches.Count -eq 0) {
            $numberMatches = [regex]::Matches($value, "[+-]?\d+(?:\.\d+)?")
            if ($numberMatches.Count -ge 1) {
                $number = 0.0
                if (ConvertTo-HooterNumber -Value $numberMatches[0].Value -Number ([ref]$number)) {
                    $allowance = [Math]::Max(0.0, $number)
                }
            }
            if ($numberMatches.Count -ge 2) {
                $number = 0.0
                if (ConvertTo-HooterNumber -Value $numberMatches[1].Value -Number ([ref]$number)) {
                    $step = [Math]::Max(0.000001, $number)
                }
            }
        }
    }

    return [pscustomobject]@{
        AllowancePercent = $allowance
        StepPercent = $step
    }
}

function Get-HooterStemPercentRuleText {
    param([object]$Tolerance)

    $settings = Get-HooterStemPercentSettings -Tolerance $Tolerance
    $points = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $Tolerance)
    return "Stem percent: $(Format-HooterScoreNumber $settings.AllowancePercent)% allowed, $points pt per $(Format-HooterScoreNumber $settings.StepPercent)% over"
}

function Get-HooterStemPercentScoreResult {
    param(
        [object]$CrewValue,
        [object]$QaValue,
        [object]$Tolerance
    )

    $crewNumber = 0.0
    $qaNumber = 0.0
    if (-not (ConvertTo-HooterNumber -Value $CrewValue -Number ([ref]$crewNumber))) { return $null }
    if (-not (ConvertTo-HooterNumber -Value $QaValue -Number ([ref]$qaNumber))) { return $null }

    $settings = Get-HooterStemPercentSettings -Tolerance $Tolerance
    $pointsPerStep = Get-HooterPointValue -Tolerance $Tolerance
    $percentError = 0.0
    if ([Math]::Abs($qaNumber) -lt 0.000001) {
        $percentError = if ([Math]::Abs($crewNumber) -lt 0.000001) { 0.0 } else { 100.0 }
    }
    else {
        $percentError = ([Math]::Abs($crewNumber - $qaNumber) / [Math]::Abs($qaNumber)) * 100.0
    }
    $overPercent = [Math]::Max(0.0, $percentError - $settings.AllowancePercent)
    $stepCount = if ($overPercent -le 0.000001) { 0 } else { [int][Math]::Ceiling($overPercent / $settings.StepPercent) }
    $loss = $stepCount * $pointsPerStep
    $detail = if ($stepCount -eq 0) {
        "Stem percent error $(Format-HooterScoreNumber $percentError)% within $(Format-HooterScoreNumber $settings.AllowancePercent)% allowance"
    }
    else {
        "Stem percent error $(Format-HooterScoreNumber $percentError)%; $(Format-HooterScoreNumber $overPercent)% over allowance; $stepCount step(s) x $(Format-HooterScoreNumber $pointsPerStep)"
    }

    return [pscustomobject]@{
        Status = if ($loss -gt 0) { "Fail" } else { "Pass" }
        Detail = $detail
        PercentError = $percentError
        OverPercent = $overPercent
        StepCount = $stepCount
        PointLoss = $loss
        PointsPerStep = $pointsPerStep
        PossiblePoints = [Math]::Max($pointsPerStep, $loss)
    }
}

function Test-HooterToleranceIsRegenStemCount {
    param(
        [string]$FieldKey,
        [object]$Tolerance = $null
    )

    if (Test-HooterRegenStemCountFieldKey -FieldKey $FieldKey) { return $true }
    if ($null -eq $Tolerance) { return $false }
    return (Test-HooterRegenStemCountField `
        -Group (ConvertTo-HooterText $Tolerance.Group) `
        -FieldName (ConvertTo-HooterText $Tolerance.FieldName) `
        -Label (ConvertTo-HooterText $Tolerance.Label))
}

function Apply-HooterStemScoringPreferenceToTolerance {
    param(
        [object]$Tolerance,
        [string]$FieldKey = ""
    )

    if ($null -eq $Tolerance) { return $Tolerance }
    $fieldKeyText = ConvertTo-HooterText $FieldKey
    if ([string]::IsNullOrWhiteSpace($fieldKeyText) -and $null -ne $Tolerance.PSObject.Properties["FieldKey"]) {
        $fieldKeyText = ConvertTo-HooterText $Tolerance.FieldKey
    }
    if (-not (Test-HooterToleranceIsRegenStemCount -FieldKey $fieldKeyText -Tolerance $Tolerance)) { return $Tolerance }

    $oldMode = ConvertTo-HooterText $Tolerance.Mode
    if ($script:UseStemCountPercentageForScoring) {
        $Tolerance.Mode = "StemPercent"
        if (-not $oldMode.Equals("StemPercent", [System.StringComparison]::OrdinalIgnoreCase) -or
            [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $Tolerance.Value))) {
            $Tolerance.Value = "Allowance=10; Step=5"
        }
    }
    else {
        $Tolerance.Mode = "StemCount"
        $Tolerance.Value = ""
    }
    return $Tolerance
}

function Get-DefaultToleranceForField {
    param(
        [string]$Group,
        [string]$TableName,
        [string]$FieldName,
        [string]$Label = "",
        [object]$QAPoints = ""
    )

    $key = Get-NameKey ("$FieldName $Label")
    $mode = "Exact"
    $value = ""
    $pointValue = Normalize-HooterQaPointValue $QAPoints
    if ([string]::IsNullOrWhiteSpace($pointValue)) { $pointValue = "1" }
    $criticalFail = $false

    if (Test-HooterFilledOnlyField -Group $Group -FieldName $FieldName -Label $Label) {
        $mode = "Filled"
        $value = ""
    }
    elseif (Test-HooterRegenStemCountField -Group $Group -FieldName $FieldName -Label $Label) {
        $mode = "StemCount"
        $value = ""
    }
    elseif ((ConvertTo-HooterText $Group).Equals("Regen", [System.StringComparison]::OrdinalIgnoreCase) -and $key -match "^(idbh|regenidbh)$|regen.*idbh|idbh.*regen") {
        $mode = "Exact"
        $value = ""
    }
    elseif ($key -match "utm.*(east|north)|easting|northing") {
        $mode = "Range"
        $value = "30"
    }
    elseif ($key -match "siteindex.*dbh|sitedbh") {
        $mode = "Range"
        $value = "0.10"
    }
    elseif ($key -match "siteindex.*height|siteheight") {
        $mode = "Range"
        $value = "2"
    }
    elseif ($key -match "elevation") {
        $mode = "Range"
        $value = "100"
    }
    elseif ($key -match "dbh|diameter") {
        $mode = "Range"
        $value = "0.20"
    }
    elseif ($key -match "slopepercent|standage|stockabilitypercent|bdftdefect") {
        $mode = "Percent"
        $value = "5"
    }
    elseif ($key -match "crownratio") {
        $mode = "Percent"
        $value = "10"
    }
    elseif ($key -match "height|sitetotalheight") {
        $mode = "Percent"
        $value = "5"
    }
    elseif ($key -match "crownclass|covertype|densityclass|sizeclass|slopeposition|severity|snagclass|defect|class") {
        $mode = "Class"
        $value = "1"
    }
    elseif ($key -match "treefound|treepresent|treemissing|missingtree|passfail|pass|fail") {
        $mode = "PassFail"
        $value = ""
    }

    if ($mode -eq "PassFail") { $criticalFail = $true }

    return [pscustomobject]@{
        FieldKey = Get-HooterFieldKey -Group $Group -TableName $TableName -FieldName $FieldName
        Group = $Group
        TableName = $TableName
        FieldName = $FieldName
        Label = if ([string]::IsNullOrWhiteSpace($Label)) { $FieldName } else { $Label }
        Mode = $mode
        Value = $value
        PointValue = $pointValue
        CriticalFail = $criticalFail
    }
}

function Get-HooterTolerance {
    param([object]$Field)

    $fieldKey = if ($Field -is [string]) { $Field } else { $Field.FieldKey }
    $databasePointValue = Get-HooterDatabaseFieldPointValue -FieldKey $fieldKey -Field $Field
    $isExecutionField = Test-HooterExecutionFieldKey -FieldKey $fieldKey
    if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and $script:Tolerances.ContainsKey($fieldKey)) {
        $stored = $script:Tolerances[$fieldKey]
        $stored.Mode = Normalize-HooterToleranceMode $stored.Mode
        if ($isExecutionField) {
            $parts = $fieldKey -split "\|", 3
            $defaultExecution = if ($parts.Count -eq 3) { Get-HooterDefaultExecutionTolerance -ItemKey $parts[2] } else { $null }
            if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.Mode)) -or (ConvertTo-HooterText $stored.Mode) -eq "Exact") {
                $stored.Mode = "GoodFairPoor"
            }
            if ($null -ne $defaultExecution) {
                if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.Value))) { $stored.Value = ConvertTo-HooterText $defaultExecution.Value }
                if ($null -eq $stored.PSObject.Properties["PointValue"] -or [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.PointValue))) {
                    Set-HooterTolerancePointValue -Tolerance $stored -PointValue $defaultExecution.PointValue
                }
                if ($null -eq $stored.PSObject.Properties["CriticalFail"]) {
                    $stored | Add-Member -MemberType NoteProperty -Name "CriticalFail" -Value (Get-HooterCriticalFail -Tolerance $defaultExecution)
                }
            }
            return $stored
        }
        if (((ConvertTo-HooterText $stored.Mode) -eq "Exact" -or [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.Mode))) -and
            ((Test-HooterFilledOnlyFieldKey -FieldKey $fieldKey) -or (Test-HooterFilledOnlyField -Group (ConvertTo-HooterText $stored.Group) -FieldName (ConvertTo-HooterText $stored.FieldName) -Label (ConvertTo-HooterText $stored.Label)))) {
            $stored.Mode = "Filled"
            $stored.Value = ""
        }
        if (((ConvertTo-HooterText $stored.Mode) -eq "Exact" -or [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.Mode))) -and
            ((Test-HooterRegenStemCountFieldKey -FieldKey $fieldKey) -or (Test-HooterRegenStemCountField -Group (ConvertTo-HooterText $stored.Group) -FieldName (ConvertTo-HooterText $stored.FieldName) -Label (ConvertTo-HooterText $stored.Label)))) {
            $stored.Mode = "StemCount"
            $stored.Value = ""
        }
        $storedNameKey = Get-NameKey ("$(ConvertTo-HooterText $stored.FieldName) $(ConvertTo-HooterText $stored.Label)")
        if ($storedNameKey -match "utm.*(east|north)|easting|northing") {
            if ((ConvertTo-HooterText $stored.Mode) -eq "Exact" -or [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.Value))) {
                $stored.Mode = "Range"
                $stored.Value = "30"
            }
        }
        if ((ConvertTo-HooterText $stored.Group).Equals("Regen", [System.StringComparison]::OrdinalIgnoreCase) -and $storedNameKey -match "^(idbh|regenidbh)$|regen.*idbh|idbh.*regen") {
            $stored.Mode = "Exact"
            $stored.Value = ""
        }
        [void](Apply-HooterStemScoringPreferenceToTolerance -Tolerance $stored -FieldKey $fieldKey)
        if (-not [string]::IsNullOrWhiteSpace($databasePointValue) -and [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stored.PointValue))) {
            Set-HooterTolerancePointValue -Tolerance $stored -PointValue $databasePointValue
        }
        return $stored
    }
    if ($Field -is [string]) {
        $parts = $Field -split "\|", 3
        if ($parts.Count -eq 3) {
            if (Test-HooterExecutionFieldKey -FieldKey $Field) {
                return Get-HooterDefaultExecutionTolerance -ItemKey $parts[2]
            }
            $defaultTolerance = Get-DefaultToleranceForField -Group $parts[0] -TableName $parts[1] -FieldName $parts[2] -QAPoints $databasePointValue
            return (Apply-HooterStemScoringPreferenceToTolerance -Tolerance $defaultTolerance -FieldKey $Field)
        }
    }
    if ($isExecutionField) {
        return Get-HooterDefaultExecutionTolerance -ItemKey $Field.FieldName
    }
    $defaultFieldTolerance = Get-DefaultToleranceForField -Group $Field.Group -TableName $Field.TableName -FieldName $Field.FieldName -Label $Field.Label -QAPoints $databasePointValue
    return (Apply-HooterStemScoringPreferenceToTolerance -Tolerance $defaultFieldTolerance -FieldKey $fieldKey)
}

function Get-HooterTreeFoundField {
    return [pscustomobject]@{
        Group = "Tree"
        TableName = "PlotHoot QA"
        FieldName = "TreeFound"
        Label = "Crew-recorded tree found by QA"
        FieldKey = Get-HooterFieldKey -Group "Tree" -TableName "PlotHoot QA" -FieldName "TreeFound"
    }
}

function Get-HooterExecutionBaseItems {
    return @(
        [pscustomobject]@{ ItemKey = "StartingPointDistanceAzimuth"; Item = "Starting Point distance and azimuth provided/correct"; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 0 }
        [pscustomobject]@{ ItemKey = "PlotReferenceTreeSelection"; Item = "Plot Reference Tree(s) Selection"; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 1 }
        [pscustomobject]@{ ItemKey = "PlotReferenceTreesMonumented"; Item = "Plot Reference Tree(s) clearly monumented (tags, nailing)"; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 2 }
        [pscustomobject]@{ ItemKey = "PlotReferenceTreeDistanceAzimuth"; Item = "Plot Reference Tree(s) distance and azimuth provided/correct"; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 3 }
        [pscustomobject]@{ ItemKey = "SiteTreesMonumented"; Item = "Site Tree(s) clearly monumented (tags, nailing)"; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 4 }
        [pscustomobject]@{ ItemKey = "GpsCoordinatesAccuracy"; Item = "GPS Coordinates 30' Accuracy (UTMs N/E/Zone)"; FairLoss = 0; PoorLoss = 0; CriticalOnPoor = $true; DefaultOrder = 5 }
    )
}

function Get-HooterExecutionFieldKey {
    param([string]$ItemKey)
    return (Get-HooterFieldKey -Group "Execution" -TableName "TableD" -FieldName $ItemKey)
}

function Test-HooterExecutionFieldKey {
    param([string]$FieldKey)

    $parts = (ConvertTo-HooterText $FieldKey) -split "\|", 3
    return ($parts.Count -eq 3 -and
        $parts[0].Equals("Execution", [System.StringComparison]::OrdinalIgnoreCase) -and
        $parts[1].Equals("TableD", [System.StringComparison]::OrdinalIgnoreCase))
}

function Get-HooterExecutionBaseItem {
    param([string]$ItemKey)

    $cleanKey = ConvertTo-HooterText $ItemKey
    foreach ($item in Get-HooterExecutionBaseItems) {
        if ((ConvertTo-HooterText $item.ItemKey).Equals($cleanKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
    }
    return $null
}

function Test-HooterObsoleteExecutionFieldKey {
    param([string]$FieldKey)

    if (-not (Test-HooterExecutionFieldKey -FieldKey $FieldKey)) { return $false }
    $parts = (ConvertTo-HooterText $FieldKey) -split "\|", 3
    if ($parts.Count -ne 3) { return $false }
    return ($null -eq (Get-HooterExecutionBaseItem -ItemKey $parts[2]))
}

function Format-HooterExecutionToleranceValue {
    param(
        [object]$FairLoss,
        [object]$PoorLoss
    )

    return "Good=0; Fair=$(Format-HooterScoreNumber $FairLoss); Poor=$(Format-HooterScoreNumber $PoorLoss)"
}

function Get-HooterDefaultExecutionTolerance {
    param([string]$ItemKey)

    $base = Get-HooterExecutionBaseItem -ItemKey $ItemKey
    if ($null -eq $base) {
        $base = [pscustomobject]@{ ItemKey = $ItemKey; Item = $ItemKey; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 0 }
    }

    return [pscustomobject]@{
        FieldKey = Get-HooterExecutionFieldKey -ItemKey $base.ItemKey
        Group = "Execution"
        TableName = "TableD"
        FieldName = $base.ItemKey
        Label = $base.Item
        Mode = "GoodFairPoor"
        Value = Format-HooterExecutionToleranceValue -FairLoss $base.FairLoss -PoorLoss $base.PoorLoss
        PointValue = Format-HooterScoreNumber $base.PoorLoss
        CriticalFail = [bool]$base.CriticalOnPoor
    }
}

function Get-HooterExecutionSettingFields {
    $fields = New-Object System.Collections.Generic.List[object]
    foreach ($item in Get-HooterExecutionBaseItems) {
        [void]$fields.Add([pscustomobject]@{
            Group = "Execution"
            TableName = "TableD"
            FieldName = $item.ItemKey
            Label = $item.Item
            FieldKey = Get-HooterExecutionFieldKey -ItemKey $item.ItemKey
            QAPoints = Format-HooterScoreNumber $item.PoorLoss
            DefaultOrder = [int]$item.DefaultOrder
        })
    }
    return @(Sort-HooterFieldsBySavedOrder -Fields @($fields.ToArray()))
}

function Get-HooterAllScoredFields {
    $fields = New-Object System.Collections.Generic.List[object]
    foreach ($field in @($script:FieldCatalog.Plot)) { [void]$fields.Add($field) }
    foreach ($field in @(Get-HooterTreeEntryFields)) { [void]$fields.Add($field) }
    foreach ($field in @($script:FieldCatalog.Regen)) { [void]$fields.Add($field) }
    foreach ($field in @(Get-HooterExecutionSettingFields)) { [void]$fields.Add($field) }
    return @($fields.ToArray())
}

function Get-HooterSavedFieldOrder {
    param([object]$Field)

    $fieldKey = if ($Field -is [string]) { ConvertTo-HooterText $Field } else { ConvertTo-HooterText $Field.FieldKey }
    if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and $script:FieldOrder.ContainsKey($fieldKey)) {
        $order = 0
        if ([int]::TryParse((ConvertTo-HooterText $script:FieldOrder[$fieldKey]), [ref]$order)) {
            return $order
        }
    }
    return 100000
}

function Get-HooterDefaultFieldOrder {
    param(
        [object]$Field,
        [int]$Index
    )

    if ($null -ne $Field -and $null -ne $Field.PSObject.Properties["DefaultOrder"]) {
        $order = 0
        if ([int]::TryParse((ConvertTo-HooterText $Field.DefaultOrder), [ref]$order)) {
            return $order
        }
    }
    return $Index
}

function Sort-HooterFieldsBySavedOrder {
    param([object[]]$Fields)

    $items = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($field in @($Fields)) {
        [void]$items.Add([pscustomobject]@{
            Field = $field
            SavedOrder = Get-HooterSavedFieldOrder -Field $field
            DefaultOrder = Get-HooterDefaultFieldOrder -Field $field -Index $index
        })
        $index++
    }
    return @($items.ToArray() | Sort-Object SavedOrder, DefaultOrder | ForEach-Object { $_.Field })
}

function Apply-HooterSavedFieldOrderToCatalog {
    if ($null -eq $script:FieldCatalog) { return }
    $script:FieldCatalog.Plot = @(Sort-HooterFieldsBySavedOrder -Fields @($script:FieldCatalog.Plot))
    $script:FieldCatalog.Tree = @(Sort-HooterFieldsBySavedOrder -Fields @($script:FieldCatalog.Tree))
    $script:FieldCatalog.Regen = @(Sort-HooterFieldsBySavedOrder -Fields @($script:FieldCatalog.Regen))
}

function Get-HooterTreeEntryFields {
    $fields = @((Get-HooterTreeFoundField)) + @($script:FieldCatalog.Tree)
    return @(Sort-HooterFieldsBySavedOrder -Fields $fields)
}

function ConvertTo-HooterNumber {
    param(
        [object]$Value,
        [ref]$Number
    )

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    $clean = $text -replace "[,%]", ""
    $parsed = 0.0
    if ([double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed) -or [double]::TryParse($clean, [ref]$parsed)) {
        $Number.Value = $parsed
        return $true
    }
    return $false
}

function Format-HooterScoreNumber {
    param([object]$Value)

    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $Value -Number ([ref]$number))) { return "0" }
    if ([Math]::Abs($number - [Math]::Round($number)) -lt 0.000001) {
        return ([Math]::Round($number)).ToString("0", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $number.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Normalize-HooterQaPointValue {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $text -Number ([ref]$number))) { return "" }
    return (Format-HooterScoreNumber ([Math]::Max(0.0, $number)))
}

function Get-HooterFieldQaPoints {
    param([object]$Field)

    if ($null -eq $Field -or $Field -is [string]) { return "" }
    foreach ($propertyName in @("QAPoints", "QApoints", "QaPoints", "QAPoint", "QAPts")) {
        if ($null -ne $Field.PSObject.Properties[$propertyName]) {
            $points = Normalize-HooterQaPointValue $Field.$propertyName
            if (-not [string]::IsNullOrWhiteSpace($points)) { return $points }
        }
    }
    return ""
}

function Get-HooterDatabaseFieldPointValue {
    param(
        [string]$FieldKey,
        [object]$Field = $null
    )

    $fieldPoints = Get-HooterFieldQaPoints -Field $Field
    if (-not [string]::IsNullOrWhiteSpace($fieldPoints)) { return $fieldPoints }
    $key = ConvertTo-HooterText $FieldKey
    if (-not [string]::IsNullOrWhiteSpace($key) -and $script:DatabaseFieldPoints.ContainsKey($key)) {
        return ConvertTo-HooterText $script:DatabaseFieldPoints[$key]
    }
    return ""
}

function Set-HooterTolerancePointValue {
    param(
        [object]$Tolerance,
        [object]$PointValue
    )

    if ($null -eq $Tolerance) { return }
    $points = Normalize-HooterQaPointValue $PointValue
    if ([string]::IsNullOrWhiteSpace($points)) { return }
    if ($null -eq $Tolerance.PSObject.Properties["PointValue"]) {
        $Tolerance | Add-Member -MemberType NoteProperty -Name "PointValue" -Value $points
    }
    else {
        $Tolerance.PointValue = $points
    }
}

function Normalize-HooterScorePassPercent {
    param([object]$Value)

    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $Value -Number ([ref]$number))) {
        $number = 90.0
    }
    $number = [Math]::Max(0.0, [Math]::Min(100.0, $number))
    return (Format-HooterScoreNumber $number)
}

function Get-HooterScorePassPercent {
    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $script:ScorePassPercent -Number ([ref]$number))) {
        $number = 90.0
    }
    return [Math]::Max(0.0, [Math]::Min(100.0, $number))
}

function Normalize-HooterMaxPointLoss {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $text -Number ([ref]$number))) { return "" }
    $number = [Math]::Max(0.0, $number)
    return (Format-HooterScoreNumber $number)
}

function Test-HooterMaxPointLossConfigured {
    param([object]$Value = $script:MaxPointLoss)

    $number = 0.0
    return (ConvertTo-HooterNumber -Value $Value -Number ([ref]$number))
}

function Get-HooterMaxPointLoss {
    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $script:MaxPointLoss -Number ([ref]$number))) {
        return [double]::PositiveInfinity
    }
    return [Math]::Max(0.0, $number)
}

function Get-HooterMaxPointLossDisplay {
    param([object]$Value = $script:MaxPointLoss)

    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $Value -Number ([ref]$number))) { return "Not set" }
    return (Format-HooterScoreNumber $number)
}

function Get-HooterMaxPointLossSummaryText {
    param([object]$Value = $script:MaxPointLoss)

    if (Test-HooterMaxPointLossConfigured -Value $Value) {
        return "fail if > $(Get-HooterMaxPointLossDisplay -Value $Value)"
    }
    return "no total loss threshold"
}

function Normalize-HooterPointTotal {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $number = 0.0
    if (-not (ConvertTo-HooterNumber -Value $text -Number ([ref]$number))) { return "" }
    $number = [Math]::Max(0.0, $number)
    return (Format-HooterScoreNumber $number)
}

function Get-HooterTreeScoringCount {
    $maxCount = 0

    if ($null -ne $script:CurrentPlot) {
        $count = @($script:CurrentPlot.TreeRecords).Count
        if ($count -gt 0) { $maxCount = [Math]::Max($maxCount, $count) }
    }

    if ($script:Ui.ContainsKey("TreeRecordBox") -and $script:Ui.TreeRecordBox.Items.Count -gt 0) {
        $maxCount = [Math]::Max($maxCount, $script:Ui.TreeRecordBox.Items.Count)
    }

    if ($script:Ui.ContainsKey("TreeGrid") -and $script:Ui.TreeGrid.Columns.Contains("EntryNumber")) {
        $entries = @{}
        foreach ($row in $script:Ui.TreeGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $entryNumber = ConvertTo-HooterText $row.Cells["EntryNumber"].Value
            if (-not [string]::IsNullOrWhiteSpace($entryNumber)) {
                $entries[$entryNumber] = $true
            }
        }
        if ($entries.Count -gt 0) { $maxCount = [Math]::Max($maxCount, $entries.Count) }
    }

    return $maxCount
}

function Get-HooterTreeMaxLossSummary {
    $treeCount = Get-HooterTreeScoringCount
    $perTreeMax = 0.0
    $hasPerTreeMax = ((ConvertTo-HooterNumber -Value $script:TreePointTotal -Number ([ref]$perTreeMax)) -and $perTreeMax -gt 0)
    $totalMax = if ($hasPerTreeMax -and $treeCount -gt 0) { $perTreeMax * $treeCount } else { 0.0 }

    return [pscustomobject]@{
        PerTreeMax = $perTreeMax
        HasPerTreeMax = $hasPerTreeMax
        TreeCount = $treeCount
        TotalMax = $totalMax
        PerTreeText = if ($hasPerTreeMax) { Format-HooterScoreNumber $perTreeMax } else { "Auto" }
        TotalText = if ($hasPerTreeMax) { Format-HooterScoreNumber $totalMax } else { "Auto" }
    }
}

function Get-HooterSectionPointTotal {
    param(
        [string]$Scope,
        [double]$Fallback = 0.0
    )

    $setting = switch ($Scope) {
        "Plot" { $script:PlotPointTotal }
        "Tree" { $script:TreePointTotal }
        "Regen" { $script:RegenPointTotal }
        "Execution" { $script:ExecutionPointTotal }
        default { "" }
    }
    $number = 0.0
    if ((ConvertTo-HooterNumber -Value $setting -Number ([ref]$number)) -and $number -gt 0) {
        if ($Scope -eq "Tree") {
            $treeCount = Get-HooterTreeScoringCount
            if ($treeCount -gt 0) {
                return ($number * $treeCount)
            }
        }
        return $number
    }
    return [Math]::Max(0.0, $Fallback)
}

function Get-HooterPointValue {
    param([object]$Tolerance)

    $pointText = ""
    if ($null -ne $Tolerance -and $null -ne $Tolerance.PSObject.Properties["PointValue"]) {
        $pointText = ConvertTo-HooterText $Tolerance.PointValue
    }
    $points = 0.0
    if (-not (ConvertTo-HooterNumber -Value $pointText -Number ([ref]$points))) {
        $points = 1.0
    }
    return [Math]::Max(0.0, $points)
}

function Get-HooterCriticalFail {
    param([object]$Tolerance)

    if ($null -ne $Tolerance -and $null -ne $Tolerance.PSObject.Properties["CriticalFail"]) {
        $criticalText = ConvertTo-HooterText $Tolerance.CriticalFail
        if (-not [string]::IsNullOrWhiteSpace($criticalText)) {
            return (ConvertTo-HooterBool $Tolerance.CriticalFail)
        }
    }
    return ((ConvertTo-HooterText $Tolerance.Mode) -eq "PassFail")
}

function Get-HooterRowScoreResult {
    param(
        [string]$Status,
        [object]$Tolerance,
        [object]$CrewValue = "",
        [object]$QaValue = ""
    )

    $points = Get-HooterPointValue -Tolerance $Tolerance
    $critical = Get-HooterCriticalFail -Tolerance $Tolerance
    $earned = 0.0
    $lost = 0.0
    $criticalFailure = $false
    $display = ""

    switch ($Status) {
        "Pass" {
            $earned = $points
            $display = "Loss 0"
        }
        "Fail" {
            $lost = $points
            if ((ConvertTo-HooterText $Tolerance.Mode) -eq "StemPercent") {
                $stemPercentResult = Get-HooterStemPercentScoreResult -CrewValue $CrewValue -QaValue $QaValue -Tolerance $Tolerance
                if ($null -ne $stemPercentResult) {
                    $lost = [Math]::Max(0.0, [double]$stemPercentResult.PointLoss)
                    $points = [Math]::Max($points, [double]$stemPercentResult.PossiblePoints)
                }
            }
            if ($critical) {
                $criticalFailure = $true
                $display = "Critical fail"
            }
            else {
                $display = "Loss $(Format-HooterScoreNumber $lost)"
            }
        }
    }

    return [pscustomobject]@{
        PointValue = $points
        CriticalFail = $critical
        EarnedPoints = $earned
        PossiblePoints = $points
        LostPoints = $lost
        CriticalFailure = $criticalFailure
        Display = $display
    }
}

function Get-NormalizedQaText {
    param([object]$Value)
    $text = (ConvertTo-HooterText $Value).ToLowerInvariant()
    $text = $text -replace "\s+", " "
    return $text.Trim()
}

function ConvertTo-HooterCheckChoice {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    if ($Value -is [bool]) {
        if ([bool]$Value) { return "yes" } else { return "no" }
    }
    $text = Get-NormalizedQaText $Value
    $yesWords = @("1", "true", "yes", "y", "checked", "active", "on", "p", "pass", "passing", "filled", "present")
    $noWords = @("0", "false", "no", "n", "unchecked", "off", "f", "fail", "failing", "blank", "missing", "not filled", "notfilled")
    if ($text -in $yesWords) { return "yes" }
    if ($text -in $noWords) { return "no" }
    return $text
}

function Test-HooterDbhTenthsTolerance {
    param([object]$Tolerance)

    if ($null -eq $Tolerance) { return $false }
    $nameKey = Get-NameKey ("$(ConvertTo-HooterText $Tolerance.FieldName) $(ConvertTo-HooterText $Tolerance.Label)")
    return ($nameKey -match "idbh|integerdbh|dbh10|dbhx10|dbhtenths|diametertenths")
}

function Test-HooterPercentPointTolerance {
    param([object]$Tolerance)

    if ($null -eq $Tolerance) { return $false }
    $nameKey = Get-NameKey ("$(ConvertTo-HooterText $Tolerance.FieldName) $(ConvertTo-HooterText $Tolerance.Label)")
    return ($nameKey -match "percent|ratio|defect")
}

function ConvertTo-HooterComparableNumber {
    param(
        [double]$Number,
        [object]$OriginalValue,
        [object]$Tolerance
    )

    if (Test-HooterDbhTenthsTolerance -Tolerance $Tolerance) {
        $text = ConvertTo-HooterText $OriginalValue
        if ($text -match "^[+-]?\d+$") {
            return ($Number / 10.0)
        }
    }
    return $Number
}

function Test-HooterFieldPass {
    param(
        [object]$CrewValue,
        [object]$QaValue,
        [object]$Tolerance
    )

    $crewText = ConvertTo-HooterText $CrewValue
    $qaText = ConvertTo-HooterText $QaValue
    if ([string]::IsNullOrWhiteSpace($qaText)) {
        return [pscustomobject]@{ Status = ""; Detail = "Not checked" }
    }

    $mode = Normalize-HooterToleranceMode $Tolerance.Mode
    $limitText = ConvertTo-HooterText $Tolerance.Value
    if ($mode -eq "Filled") {
        $choice = ConvertTo-HooterCheckChoice $QaValue
        if ($choice -eq "yes") {
            if (-not [string]::IsNullOrWhiteSpace($crewText)) { return [pscustomobject]@{ Status = "Pass"; Detail = "Crew value was filled in" } }
            return [pscustomobject]@{ Status = "Fail"; Detail = "Crew value is blank" }
        }
        if ($choice -eq "no") {
            return [pscustomobject]@{ Status = "Fail"; Detail = "Marked not filled in" }
        }
        return [pscustomobject]@{ Status = ""; Detail = "Not checked" }
    }
    if ([string]::IsNullOrWhiteSpace($crewText)) {
        if ([string]::IsNullOrWhiteSpace($qaText)) { return [pscustomobject]@{ Status = ""; Detail = "Not checked" } }
        return [pscustomobject]@{ Status = "Fail"; Detail = "Crew value is blank" }
    }

    $crewNumber = 0.0
    $qaNumber = 0.0
    $hasCrewNumber = ConvertTo-HooterNumber -Value $crewText -Number ([ref]$crewNumber)
    $hasQaNumber = ConvertTo-HooterNumber -Value $qaText -Number ([ref]$qaNumber)

    switch ($mode) {
        { $_ -in @("Range", "Absolute") } {
            if ($hasCrewNumber -and $hasQaNumber) {
                $limit = 0.0
                [void](ConvertTo-HooterNumber -Value $limitText -Number ([ref]$limit))
                $rawCrewNumber = $crewNumber
                $rawQaNumber = $qaNumber
                $crewNumber = ConvertTo-HooterComparableNumber -Number $crewNumber -OriginalValue $crewText -Tolerance $Tolerance
                $qaNumber = ConvertTo-HooterComparableNumber -Number $qaNumber -OriginalValue $qaText -Tolerance $Tolerance
                $diff = [Math]::Abs($crewNumber - $qaNumber)
                $detail = if (Test-HooterDbhTenthsTolerance -Tolerance $Tolerance) {
                    "DBH difference $(Format-HooterScoreNumber $diff) in; raw values $(Format-HooterScoreNumber $rawCrewNumber) vs $(Format-HooterScoreNumber $rawQaNumber)"
                }
                else {
                    "Difference $([Math]::Round($diff, 4))"
                }
                if ($diff -le ($limit + 0.000001)) { return [pscustomobject]@{ Status = "Pass"; Detail = "$detail <= $limit" } }
                return [pscustomobject]@{ Status = "Fail"; Detail = "$detail > $limit" }
            }
        }
        "Percent" {
            if ($hasCrewNumber -and $hasQaNumber) {
                $percent = 0.0
                [void](ConvertTo-HooterNumber -Value $limitText -Number ([ref]$percent))
                $diff = [Math]::Abs($crewNumber - $qaNumber)
                if (Test-HooterPercentPointTolerance -Tolerance $Tolerance) {
                    $limit = [Math]::Abs($percent)
                    $detail = "Percentage-point difference $([Math]::Round($diff, 4))"
                    if ($diff -le ($limit + 0.000001)) { return [pscustomobject]@{ Status = "Pass"; Detail = "$detail <= $limit" } }
                    return [pscustomobject]@{ Status = "Fail"; Detail = "$detail > $limit" }
                }
                else {
                    $percentLimit = [Math]::Abs($percent)
                    $limit = [Math]::Abs($crewNumber) * ($percentLimit / 100.0)
                    if ([Math]::Abs($crewNumber) -lt 0.000001) { $limit = $percentLimit / 100.0 }
                    if ($diff -le ($limit + 0.000001)) { return [pscustomobject]@{ Status = "Pass"; Detail = "Difference $([Math]::Round($diff, 4)) within $percent%" } }
                    return [pscustomobject]@{ Status = "Fail"; Detail = "Difference $([Math]::Round($diff, 4)) exceeds $percent%" }
                }
            }
        }
        "Class" {
            if ($hasCrewNumber -and $hasQaNumber) {
                $limit = 0.0
                [void](ConvertTo-HooterNumber -Value $limitText -Number ([ref]$limit))
                $diff = [Math]::Abs($crewNumber - $qaNumber)
                if ($diff -le $limit) { return [pscustomobject]@{ Status = "Pass"; Detail = "Class difference $diff <= $limit" } }
                return [pscustomobject]@{ Status = "Fail"; Detail = "Class difference $diff > $limit" }
            }
        }
        "StemCount" {
            if ($hasCrewNumber -and $hasQaNumber) {
                $band = Get-HooterRegenStemToleranceBand -QaCount $qaNumber
                $limit = 0.0
                $hasStemTolerance = ConvertTo-HooterNumber -Value $band.Tolerance -Number ([ref]$limit)
                $diff = [Math]::Abs($crewNumber - $qaNumber)
                $detail = if ($hasStemTolerance) {
                    "Stem count difference $([Math]::Round($diff, 4)); QA count $([Math]::Round($qaNumber, 4)) falls in $($band.Label), which allows +/- $(Format-HooterScoreNumber $limit)"
                }
                else {
                    "Stem count difference $([Math]::Round($diff, 4)); QA count $([Math]::Round($qaNumber, 4)) falls in $($band.Label), no tolerance set so exact match is required"
                }
                if ($diff -le $limit) { return [pscustomobject]@{ Status = "Pass"; Detail = $detail } }
                return [pscustomobject]@{ Status = "Fail"; Detail = $detail }
            }
        }
        "StemPercent" {
            if ($hasCrewNumber -and $hasQaNumber) {
                $result = Get-HooterStemPercentScoreResult -CrewValue $crewNumber -QaValue $qaNumber -Tolerance $Tolerance
                if ($null -ne $result) {
                    return [pscustomobject]@{ Status = $result.Status; Detail = $result.Detail }
                }
            }
        }
        "PassFail" {
            $crewNorm = Get-NormalizedQaText $crewText
            $qaNorm = Get-NormalizedQaText $qaText
            $passWords = @{
                "p" = "pass"; "pass" = "pass"; "passing" = "pass"; "yes" = "pass"; "y" = "pass"; "true" = "pass"; "1" = "pass";
                "f" = "fail"; "fail" = "fail"; "failing" = "fail"; "no" = "fail"; "n" = "fail"; "false" = "fail"; "0" = "fail"
            }
            if ($passWords.ContainsKey($crewNorm)) { $crewNorm = $passWords[$crewNorm] }
            if ($passWords.ContainsKey($qaNorm)) { $qaNorm = $passWords[$qaNorm] }
            if ($crewNorm -eq $qaNorm) { return [pscustomobject]@{ Status = "Pass"; Detail = "Pass/fail matched" } }
            return [pscustomobject]@{ Status = "Fail"; Detail = "Pass/fail did not match" }
        }
    }

    if ((Get-NormalizedQaText $crewText) -eq (Get-NormalizedQaText $qaText)) {
        return [pscustomobject]@{ Status = "Pass"; Detail = "Exact match" }
    }
    return [pscustomobject]@{ Status = "Fail"; Detail = "Values do not match" }
}

function Get-HooterColumnTypeName {
    param([System.Data.DataColumn]$Column)

    if ($null -eq $Column -or $null -eq $Column.DataType) { return "Text" }
    switch ($Column.DataType.FullName) {
        "System.Guid" { return "Guid" }
        "System.Byte" { return "Number" }
        "System.Int16" { return "Number" }
        "System.Int32" { return "Number" }
        "System.Int64" { return "Number" }
        "System.Single" { return "Number" }
        "System.Double" { return "Number" }
        "System.Decimal" { return "Number" }
        "System.DateTime" { return "Date" }
        "System.Boolean" { return "Yes/No" }
        default { return "Text" }
    }
}

function Get-TableColumnInfo {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $table = Get-DbTable -Connection $Connection -Sql "SELECT * FROM $(Quote-Name $TableName) WHERE 1=0"
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($column in $table.Columns) {
        [void]$items.Add([pscustomobject]@{
            ColumnName = [string]$column.ColumnName
            DataType = Get-HooterColumnTypeName -Column $column
        })
    }
    return @($items.ToArray())
}

function Get-HooterColumnDataType {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    foreach ($column in (Get-TableColumnInfo -Connection $Connection -TableName $TableName)) {
        if ($column.ColumnName.Equals($FieldName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return ConvertTo-HooterText $column.DataType
        }
    }
    return ""
}

function New-HooterFieldEqualsPredicate {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName,
        [object]$Value
    )

    $dataType = Get-HooterColumnDataType -Connection $Connection -TableName $TableName -FieldName $FieldName
    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "$(Quote-Name $FieldName) Is Null"
    }
    if ($dataType.Equals("Guid", [System.StringComparison]::OrdinalIgnoreCase)) {
        return "$(Quote-Name $FieldName) = $(Sql-Guid $text)"
    }
    return "$(Field-Text-Expression $FieldName) = $(Sql-Text $text)"
}

function Get-HooterPhysicalFields {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$Tables
    )

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($tableName in @("Plots", "PlotMeasurements", "PlotCustomMeasurements", "Trees", "TreeMeasurements", "TreeCustomMeasurements", "RegenMeasurements", "RegenCustomMeasurements")) {
        $actualTable = Get-ActualTableName -Tables $Tables -Name $tableName
        if ([string]::IsNullOrWhiteSpace($actualTable)) { continue }
        $role = Get-HooterTableRole -TableName $actualTable
        if ([string]::IsNullOrWhiteSpace($role)) { continue }
        foreach ($column in (Get-TableColumnInfo -Connection $Connection -TableName $actualTable)) {
            if (Test-HooterHiddenField -TableName $actualTable -FieldName $column.ColumnName) { continue }
            $fieldKey = Get-HooterFieldKey -Group $role -TableName $actualTable -FieldName $column.ColumnName
            [void]$items.Add([pscustomobject]@{
                FieldKey = $fieldKey
                Group = $role
                TableName = $actualTable
                FieldName = $column.ColumnName
                Label = $column.ColumnName
                DataType = $column.DataType
                Category = ""
                QAPoints = ""
                Source = "Physical"
            })
        }
    }
    return @($items.ToArray())
}

function Get-HooterAppColumnFields {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$Tables
    )

    if (-not (Test-TableExists -Tables $Tables -Name "AppColumns")) { return @() }

    $columnTable = Get-DbTable -Connection $Connection -Sql "SELECT * FROM [AppColumns]"
    $tableLookup = @{}
    if (Test-TableExists -Tables $Tables -Name "AppTables") {
        $tableRows = Get-DbTable -Connection $Connection -Sql "SELECT * FROM [AppTables]"
        foreach ($tableRow in $tableRows.Rows) {
            $key = Get-DataRowText -Row $tableRow -ColumnNames @("ID", "TableID", "AppTableID")
            $name = Get-DataRowText -Row $tableRow -ColumnNames @("TableName", "AppTableName", "Name", "Description")
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not [string]::IsNullOrWhiteSpace($name) -and -not $tableLookup.ContainsKey($key)) {
                $tableLookup[$key] = $name
            }
        }
    }

    $typeLookup = @{}
    if (Test-TableExists -Tables $Tables -Name "AppColumnDataTypes") {
        $typeRows = Get-DbTable -Connection $Connection -Sql "SELECT * FROM [AppColumnDataTypes]"
        foreach ($typeRow in $typeRows.Rows) {
            $key = Get-DataRowText -Row $typeRow -ColumnNames @("ID", "DataTypeID", "AppColumnDataTypeID")
            $name = Get-DataRowText -Row $typeRow -ColumnNames @("TypeName", "DataType", "DataTypeName", "Name", "Description")
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not [string]::IsNullOrWhiteSpace($name) -and -not $typeLookup.ContainsKey($key)) {
                $typeLookup[$key] = $name
            }
        }
    }

    $categoryLookup = @{}
    if (Test-TableExists -Tables $Tables -Name "AppColumnCategoryTypes") {
        $categoryRows = Get-DbTable -Connection $Connection -Sql "SELECT * FROM [AppColumnCategoryTypes]"
        foreach ($categoryRow in $categoryRows.Rows) {
            $key = Get-DataRowText -Row $categoryRow -ColumnNames @("ID", "CategoryTypeID", "AppColumnCategoryTypeID")
            $name = Get-DataRowText -Row $categoryRow -ColumnNames @("CategoryName", "Category", "CategoryTypeName", "Name", "Description")
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not [string]::IsNullOrWhiteSpace($name) -and -not $categoryLookup.ContainsKey($key)) {
                $categoryLookup[$key] = $name
            }
        }
    }

    $items = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $appColumnRowOrder = 0
    foreach ($row in $columnTable.Rows) {
        $fieldName = Get-DataRowText -Row $row -ColumnNames @("ColumnName", "FieldName", "Name")
        $sourceOrder = $appColumnRowOrder
        $appColumnRowOrder++
        if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }

        $activeColumnNames = @("Active", "IsActive", "Collected", "QueryVisible", "ReportVisible")
        $hasActiveFlag = $false
        foreach ($activeColumn in $activeColumnNames) {
            foreach ($column in $row.Table.Columns) {
                if ($column.ColumnName.Equals($activeColumn, [System.StringComparison]::OrdinalIgnoreCase)) { $hasActiveFlag = $true }
            }
        }
        $active = if ($hasActiveFlag) {
            (ConvertTo-HooterBool (Get-DataRowValue -Row $row -ColumnNames @("Active", "IsActive", "Collected"))) -or
            (ConvertTo-HooterBool (Get-DataRowValue -Row $row -ColumnNames @("QueryVisible", "ReportVisible")))
        }
        else {
            $true
        }
        if (-not $active) { continue }

        $tableName = Get-DataRowText -Row $row -ColumnNames @("TableName", "AppTableName")
        if ([string]::IsNullOrWhiteSpace($tableName)) {
            $tableId = Get-DataRowText -Row $row -ColumnNames @("TableID", "AppTableID")
            if (-not [string]::IsNullOrWhiteSpace($tableId) -and $tableLookup.ContainsKey($tableId)) {
                $tableName = ConvertTo-HooterText $tableLookup[$tableId]
            }
        }
        if ([string]::IsNullOrWhiteSpace($tableName)) { continue }

        $actualTable = Get-ActualTableName -Tables $Tables -Name $tableName
        if ([string]::IsNullOrWhiteSpace($actualTable)) { continue }
        $role = Get-HooterTableRole -TableName $actualTable
        if ([string]::IsNullOrWhiteSpace($role)) { continue }
        if (Test-HooterHiddenField -TableName $actualTable -FieldName $fieldName) { continue }

        $columns = @(Get-TableColumns -Connection $Connection -TableName $actualTable)
        $actualField = Get-ActualColumnName -Columns $columns -Names @($fieldName)
        if ([string]::IsNullOrWhiteSpace($actualField)) { continue }

        $typeName = Get-DataRowText -Row $row -ColumnNames @("TypeName", "DataType", "DataTypeName")
        if ([string]::IsNullOrWhiteSpace($typeName)) {
            $typeId = Get-DataRowText -Row $row -ColumnNames @("DataTypeID", "AppColumnDataTypeID")
            if (-not [string]::IsNullOrWhiteSpace($typeId) -and $typeLookup.ContainsKey($typeId)) {
                $typeName = ConvertTo-HooterText $typeLookup[$typeId]
            }
        }
        if ([string]::IsNullOrWhiteSpace($typeName)) { $typeName = "Text" }

        $categoryName = Get-DataRowText -Row $row -ColumnNames @("CategoryName", "Category", "CategoryTypeName")
        if ([string]::IsNullOrWhiteSpace($categoryName)) {
            $categoryId = Get-DataRowText -Row $row -ColumnNames @("CategoryTypeID", "AppColumnCategoryTypeID")
            if (-not [string]::IsNullOrWhiteSpace($categoryId) -and $categoryLookup.ContainsKey($categoryId)) {
                $categoryName = ConvertTo-HooterText $categoryLookup[$categoryId]
            }
        }

        $label = Get-DataRowText -Row $row -ColumnNames @("FriendlyName", "ShortName", "Description")
        if ([string]::IsNullOrWhiteSpace($label)) { $label = $actualField }
        $qaPoints = Normalize-HooterQaPointValue (Get-DataRowText -Row $row -ColumnNames @("QApoints", "QAPoints", "QaPoints", "QAPoint", "QApoint", "QAPts", "QAPointValue"))
        $sourceOrderText = Get-DataRowText -Row $row -ColumnNames @("DisplayOrder", "FieldOrder", "ColumnOrder", "SortOrder", "Order", "OrderNumber", "Sequence", "SequenceNumber", "Position", "Ordinal", "Index")
        if (-not [string]::IsNullOrWhiteSpace($sourceOrderText)) {
            [void][int]::TryParse($sourceOrderText, [ref]$sourceOrder)
        }
        $fieldKey = Get-HooterFieldKey -Group $role -TableName $actualTable -FieldName $actualField
        if ($seen.ContainsKey($fieldKey)) { continue }
        $seen[$fieldKey] = $true
        [void]$items.Add([pscustomobject]@{
            FieldKey = $fieldKey
            Group = $role
            TableName = $actualTable
            FieldName = $actualField
            Label = $label
            DataType = $typeName
            Category = $categoryName
            QAPoints = $qaPoints
            SourceOrder = $sourceOrder
            Source = "AppColumns"
        })
    }
    return @($items.ToArray())
}

function Get-HooterFieldCatalog {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection)
    $fields = @(Get-HooterAppColumnFields -Connection $Connection -Tables $tables)
    if ($fields.Count -eq 0) {
        $fields = @(Get-HooterPhysicalFields -Connection $Connection -Tables $tables)
    }
    $fields = @($fields | Where-Object { -not (Test-HooterHiddenTreeCheckField -Field $_) })
    $queryOrders = Get-HooterMeasurementQueryOrders -Connection $Connection

    $tableOrder = @{
        "Plots" = 1
        "PlotMeasurements" = 2
        "PlotCustomMeasurements" = 3
        "Trees" = 1
        "TreeMeasurements" = 2
        "TreeCustomMeasurements" = 3
        "RegenMeasurements" = 1
        "RegenCustomMeasurements" = 2
    }
    $defaultSorted = @($fields | Sort-Object `
        @{ Expression = { $_.Group } }, `
        @{ Expression = { Get-HooterFieldQueryOrder -Field $_ -QueryOrders $queryOrders } }, `
        @{ Expression = { if ($null -ne $_.PSObject.Properties["SourceOrder"]) { [int]$_.SourceOrder } else { 100000 } } }, `
        @{ Expression = { if ($tableOrder.ContainsKey($_.TableName)) { $tableOrder[$_.TableName] } else { 99 } } }, `
        @{ Expression = { $_.FieldName } })
    $defaultOrder = 0
    foreach ($field in $defaultSorted) {
        if ($null -eq $field.PSObject.Properties["DefaultOrder"]) {
            $field | Add-Member -MemberType NoteProperty -Name "DefaultOrder" -Value $defaultOrder
        }
        else {
            $field.DefaultOrder = $defaultOrder
        }
        $defaultOrder++
    }
    $sorted = @(Sort-HooterFieldsBySavedOrder -Fields $defaultSorted)
    $script:DatabaseFieldPoints = @{}
    foreach ($field in $sorted) {
        $fieldKey = ConvertTo-HooterText $field.FieldKey
        $points = Get-HooterFieldQaPoints -Field $field
        if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and -not [string]::IsNullOrWhiteSpace($points)) {
            $script:DatabaseFieldPoints[$fieldKey] = $points
        }
    }
    return @{
        Plot = @($sorted | Where-Object { $_.Group -eq "Plot" })
        Tree = @($sorted | Where-Object { $_.Group -eq "Tree" })
        Regen = @($sorted | Where-Object { $_.Group -eq "Regen" })
    }
}

function Get-HooterCompletedPlotWhereClause {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$Tables,
        [string]$PlotTable,
        [string[]]$PlotColumns
    )

    $assignmentTable = ""
    foreach ($name in @("InventoryAssignmentPlots", "InventoryAssignmentPlot")) {
        $assignmentTable = Get-ActualTableName -Tables $Tables -Name $name
        if (-not [string]::IsNullOrWhiteSpace($assignmentTable)) { break }
    }
    if ([string]::IsNullOrWhiteSpace($assignmentTable)) {
        throw "Completed plot filtering needs an InventoryAssignmentPlots table. Use Load All Plots if this database does not have plot assignment rows."
    }

    $assignmentColumns = @(Get-TableColumns -Connection $Connection -TableName $assignmentTable)
    $statusColumn = Get-ActualColumnName -Columns $assignmentColumns -Names @("StatusID", "PlotStatusID", "PlotStatusCode", "StatusCode", "AssignmentStatusID", "PlotStatus")
    if ([string]::IsNullOrWhiteSpace($statusColumn)) {
        throw "Completed plot filtering needs a status field on $assignmentTable, such as StatusID. Use Load All Plots if this database does not track plot assignment status."
    }

    foreach ($name in @("PlotID", "PlotKey", "PlotNumber", "PlotNo", "Plot")) {
        $plotColumn = Get-ActualColumnName -Columns $PlotColumns -Names @($name)
        $assignmentPlotColumn = Get-ActualColumnName -Columns $assignmentColumns -Names @($name)
        if (-not [string]::IsNullOrWhiteSpace($plotColumn) -and -not [string]::IsNullOrWhiteSpace($assignmentPlotColumn)) {
            $subWhere = "$(Quote-Name $statusColumn) = 4"
            if (Test-ColumnExists -Columns $assignmentColumns -Name "IsDeleted") {
                $subWhere += " AND ([IsDeleted] = False OR [IsDeleted] Is Null)"
            }
            return "$(Quote-Name $plotColumn) IN (SELECT $(Quote-Name $assignmentPlotColumn) FROM $(Quote-Name $assignmentTable) WHERE $subWhere)"
        }
    }

    throw "Completed plot filtering needs a shared plot key between $PlotTable and $assignmentTable, such as PlotID. Use Load All Plots if this database does not link assignment rows to plots."
}

function Get-HooterPlotInventoryCount {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection)
    $plotTable = Get-ActualTableName -Tables $tables -Name "Plots"
    if ([string]::IsNullOrWhiteSpace($plotTable)) { throw "The selected database does not contain a Plots table." }
    $columns = @(Get-TableColumns -Connection $Connection -TableName $plotTable)
    $where = ""
    if (Test-ColumnExists -Columns $columns -Name "IsDeleted") {
        $where = " WHERE ([IsDeleted] = False OR [IsDeleted] Is Null)"
    }
    $rows = Get-DbTable -Connection $Connection -Sql "SELECT Count(*) AS PlotCount FROM $(Quote-Name $plotTable)$where"
    if ($rows.Rows.Count -eq 0) { return 0 }
    $count = 0
    [void][int]::TryParse((ConvertTo-HooterText (Get-DataRowValue -Row $rows.Rows[0] -ColumnNames @("PlotCount"))), [ref]$count)
    return $count
}

function Get-HooterPlots {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [switch]$CompletedOnly
    )

    $tables = @(Get-UserTables -Connection $Connection)
    $plotTable = Get-ActualTableName -Tables $tables -Name "Plots"
    if ([string]::IsNullOrWhiteSpace($plotTable)) { throw "The selected database does not contain a Plots table." }
    $columns = @(Get-TableColumns -Connection $Connection -TableName $plotTable)
    $plotNumberColumn = Get-ActualColumnName -Columns $columns -Names @("PlotNumber", "PlotNo", "Plot")
    if ([string]::IsNullOrWhiteSpace($plotNumberColumn)) { throw "The Plots table does not have a PlotNumber field." }

    $selectColumns = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("PlotID", "PlotNumber", "PlotLabel", "UTMEastingCoordinate", "UTMNorthingCoordinate", "UTMZone")) {
        $actual = Get-ActualColumnName -Columns $columns -Names @($name)
        if (-not [string]::IsNullOrWhiteSpace($actual)) { [void]$selectColumns.Add((Quote-Name $actual)) }
    }
    if ($selectColumns.Count -eq 0) { [void]$selectColumns.Add("*") }
    $whereParts = New-Object System.Collections.Generic.List[string]
    if (Test-ColumnExists -Columns $columns -Name "IsDeleted") {
        [void]$whereParts.Add("([IsDeleted] = False OR [IsDeleted] Is Null)")
    }
    if ($CompletedOnly) {
        [void]$whereParts.Add((Get-HooterCompletedPlotWhereClause -Connection $Connection -Tables $tables -PlotTable $plotTable -PlotColumns $columns))
    }
    $where = if ($whereParts.Count -gt 0) { " WHERE " + ([string]::Join(" AND ", [string[]]$whereParts.ToArray())) } else { "" }
    $sql = "SELECT $([string]::Join(', ', [string[]]$selectColumns.ToArray())) FROM $(Quote-Name $plotTable)$where ORDER BY $(Quote-Name $plotNumberColumn)"
    $rows = Get-DbTable -Connection $Connection -Sql $sql
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows.Rows) {
        $plotNumber = Get-DataRowText -Row $row -ColumnNames @("PlotNumber", "PlotNo", "Plot")
        if ([string]::IsNullOrWhiteSpace($plotNumber)) { continue }
        [void]$items.Add([pscustomobject]@{
            PlotNumber = $plotNumber
            PlotID = Get-DataRowText -Row $row -ColumnNames @("PlotID")
            PlotLabel = Get-DataRowText -Row $row -ColumnNames @("PlotLabel")
            UTMEasting = Get-DataRowText -Row $row -ColumnNames @("UTMEastingCoordinate", "UTMEasting", "Easting")
            UTMNorthing = Get-DataRowText -Row $row -ColumnNames @("UTMNorthingCoordinate", "UTMNorthing", "Northing")
            UTMZone = Get-DataRowText -Row $row -ColumnNames @("UTMZone", "Zone")
            Display = if ([string]::IsNullOrWhiteSpace((Get-DataRowText -Row $row -ColumnNames @("PlotLabel")))) { $plotNumber } else { "$plotNumber - $(Get-DataRowText -Row $row -ColumnNames @("PlotLabel"))" }
        })
    }
    return @($items.ToArray())
}

function Get-HooterRowsByWhere {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$Where = "",
        [string]$OrderBy = ""
    )

    $sql = "SELECT * FROM $(Quote-Name $TableName)"
    if (-not [string]::IsNullOrWhiteSpace($Where)) { $sql += " WHERE $Where" }
    if (-not [string]::IsNullOrWhiteSpace($OrderBy)) { $sql += " ORDER BY $OrderBy" }
    try {
        return Get-DbTable -Connection $Connection -Sql $sql
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($OrderBy)) { throw }
        $sql = "SELECT * FROM $(Quote-Name $TableName)"
        if (-not [string]::IsNullOrWhiteSpace($Where)) { $sql += " WHERE $Where" }
        return Get-DbTable -Connection $Connection -Sql $sql
    }
}

function Get-HooterFirstRowByWhere {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$Where = "",
        [string]$OrderBy = ""
    )

    $sql = "SELECT TOP 1 * FROM $(Quote-Name $TableName)"
    if (-not [string]::IsNullOrWhiteSpace($Where)) { $sql += " WHERE $Where" }
    if (-not [string]::IsNullOrWhiteSpace($OrderBy)) { $sql += " ORDER BY $OrderBy" }
    try {
        $table = Get-DbTable -Connection $Connection -Sql $sql
    }
    catch {
        $sql = "SELECT TOP 1 * FROM $(Quote-Name $TableName)"
        if (-not [string]::IsNullOrWhiteSpace($Where)) { $sql += " WHERE $Where" }
        $table = Get-DbTable -Connection $Connection -Sql $sql
    }
    if ($table.Rows.Count -gt 0) { return $table.Rows[0] }
    return $null
}

function Get-HooterMeasurementRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$IdColumn,
        [string]$IdValue
    )

    if ([string]::IsNullOrWhiteSpace($TableName) -or [string]::IsNullOrWhiteSpace($IdColumn) -or [string]::IsNullOrWhiteSpace($IdValue)) { return $null }
    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    if (-not (Test-ColumnExists -Columns $columns -Name $IdColumn)) { return $null }
    $where = New-HooterFieldEqualsPredicate -Connection $Connection -TableName $TableName -FieldName $IdColumn -Value $IdValue
    if (Test-ColumnExists -Columns $columns -Name "IsDeleted") {
        $where += " AND ([IsDeleted] = False OR [IsDeleted] Is Null)"
    }
    $orderParts = New-Object System.Collections.Generic.List[string]
    if (Test-ColumnExists -Columns $columns -Name "PeriodNumber") { [void]$orderParts.Add("[PeriodNumber] DESC") }
    if (Test-ColumnExists -Columns $columns -Name "Updated") { [void]$orderParts.Add("[Updated] DESC") }
    if (Test-ColumnExists -Columns $columns -Name "Created") { [void]$orderParts.Add("[Created] DESC") }
    return Get-HooterFirstRowByWhere -Connection $Connection -TableName $TableName -Where $where -OrderBy ([string]::Join(", ", [string[]]$orderParts.ToArray()))
}

function Get-HooterCustomRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [System.Data.DataRow]$MeasurementRow,
        [string[]]$KeyNames
    )

    if ([string]::IsNullOrWhiteSpace($TableName) -or $null -eq $MeasurementRow) { return $null }
    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    foreach ($keyName in @("MeasurementID") + @($KeyNames)) {
        $customColumn = Get-ActualColumnName -Columns $columns -Names @($keyName)
        if ([string]::IsNullOrWhiteSpace($customColumn)) { continue }
        $value = Get-DataRowText -Row $MeasurementRow -ColumnNames @($keyName)
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $where = New-HooterFieldEqualsPredicate -Connection $Connection -TableName $TableName -FieldName $customColumn -Value $value
        return Get-HooterFirstRowByWhere -Connection $Connection -TableName $TableName -Where $where
    }
    return $null
}

function New-HooterRecord {
    param(
        [string]$Display,
        [hashtable]$RowsByTable,
        [string]$RecordId = ""
    )

    return [pscustomobject]@{
        Display = $Display
        RecordId = $RecordId
        RowsByTable = $RowsByTable
    }
}

function Get-RecordFieldValue {
    param(
        [object]$Record,
        [object]$Field
    )

    if ($null -eq $Record -or $null -eq $Field) { return "" }
    $tableName = ConvertTo-HooterText $Field.TableName
    if ([string]::IsNullOrWhiteSpace($tableName)) { return "" }
    $row = $null
    if ($Record.PSObject.Properties["RowsByTable"] -and $Record.RowsByTable.ContainsKey($tableName)) {
        $row = $Record.RowsByTable[$tableName]
    }
    elseif ($Record -is [hashtable] -and $Record.ContainsKey($tableName)) {
        $row = $Record[$tableName]
    }
    if ($null -eq $row) { return "" }
    return ConvertTo-HooterText (Get-DataRowValue -Row $row -ColumnNames @($Field.FieldName))
}

function Get-HooterPlotData {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$PlotNumber
    )

    $tables = @(Get-UserTables -Connection $Connection)
    $plotTable = Get-ActualTableName -Tables $tables -Name "Plots"
    if ([string]::IsNullOrWhiteSpace($plotTable)) { throw "The database does not contain a Plots table." }
    $plotColumns = @(Get-TableColumns -Connection $Connection -TableName $plotTable)
    $plotNumberColumn = Get-ActualColumnName -Columns $plotColumns -Names @("PlotNumber", "PlotNo", "Plot")
    if ([string]::IsNullOrWhiteSpace($plotNumberColumn)) { throw "The Plots table does not have a PlotNumber field." }

    $where = "$(Field-Text-Expression $plotNumberColumn) = $(Sql-Text $PlotNumber)"
    if (Test-ColumnExists -Columns $plotColumns -Name "IsDeleted") {
        $where += " AND ([IsDeleted] = False OR [IsDeleted] Is Null)"
    }
    $plotRow = Get-HooterFirstRowByWhere -Connection $Connection -TableName $plotTable -Where $where
    if ($null -eq $plotRow) { throw "Plot $PlotNumber was not found in the selected database." }

    $plotId = Get-DataRowText -Row $plotRow -ColumnNames @("PlotID")
    $plotMeasTable = Get-ActualTableName -Tables $tables -Name "PlotMeasurements"
    $plotCustomTable = Get-ActualTableName -Tables $tables -Name "PlotCustomMeasurements"
    $plotMeasRow = if (-not [string]::IsNullOrWhiteSpace($plotMeasTable)) { Get-HooterMeasurementRow -Connection $Connection -TableName $plotMeasTable -IdColumn "PlotID" -IdValue $plotId } else { $null }
    $plotCustomRow = if (-not [string]::IsNullOrWhiteSpace($plotCustomTable)) { Get-HooterCustomRow -Connection $Connection -TableName $plotCustomTable -MeasurementRow $plotMeasRow -KeyNames @("PlotMeasKey") } else { $null }

    $plotRows = New-InsensitiveHashtable
    $plotRows[$plotTable] = $plotRow
    if ($null -ne $plotMeasRow) { $plotRows[$plotMeasTable] = $plotMeasRow }
    if ($null -ne $plotCustomRow) { $plotRows[$plotCustomTable] = $plotCustomRow }

    $treeRecords = New-Object System.Collections.Generic.List[object]
    $treeTable = Get-ActualTableName -Tables $tables -Name "Trees"
    $treeMeasTable = Get-ActualTableName -Tables $tables -Name "TreeMeasurements"
    $treeCustomTable = Get-ActualTableName -Tables $tables -Name "TreeCustomMeasurements"
    if (-not [string]::IsNullOrWhiteSpace($treeTable) -and -not [string]::IsNullOrWhiteSpace($plotId)) {
        $treeColumns = @(Get-TableColumns -Connection $Connection -TableName $treeTable)
        if (Test-ColumnExists -Columns $treeColumns -Name "PlotID") {
            $treeWhere = New-HooterFieldEqualsPredicate -Connection $Connection -TableName $treeTable -FieldName "PlotID" -Value $plotId
            if (Test-ColumnExists -Columns $treeColumns -Name "IsDeleted") {
                $treeWhere += " AND ([IsDeleted] = False OR [IsDeleted] Is Null)"
            }
            $treeOrder = if (Test-ColumnExists -Columns $treeColumns -Name "TreeNumber") { "[TreeNumber]" } else { "" }
            $treeRows = Get-HooterRowsByWhere -Connection $Connection -TableName $treeTable -Where $treeWhere -OrderBy $treeOrder
            foreach ($treeRow in $treeRows.Rows) {
                $treeId = Get-DataRowText -Row $treeRow -ColumnNames @("TreeID")
                $treeMeasRow = if (-not [string]::IsNullOrWhiteSpace($treeMeasTable)) { Get-HooterMeasurementRow -Connection $Connection -TableName $treeMeasTable -IdColumn "TreeID" -IdValue $treeId } else { $null }
                $treeCustomRow = if (-not [string]::IsNullOrWhiteSpace($treeCustomTable)) { Get-HooterCustomRow -Connection $Connection -TableName $treeCustomTable -MeasurementRow $treeMeasRow -KeyNames @("TreeMeasKey") } else { $null }
                $rowsByTable = New-InsensitiveHashtable
                $rowsByTable[$treeTable] = $treeRow
                if ($null -ne $treeMeasRow) { $rowsByTable[$treeMeasTable] = $treeMeasRow }
                if ($null -ne $treeCustomRow) { $rowsByTable[$treeCustomTable] = $treeCustomRow }
                $treeNumber = Get-DataRowText -Row $treeRow -ColumnNames @("TreeNumber", "TreeNo")
                $species = Get-DataRowText -Row $treeRow -ColumnNames @("SpeciesCode", "Species")
                $displayParts = @($treeNumber, $species) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                $display = if ($displayParts.Count -gt 0) { "Tree " + ([string]::Join(" / ", [string[]]$displayParts)) } else { "Tree record $($treeRecords.Count + 1)" }
                [void]$treeRecords.Add((New-HooterRecord -Display $display -RowsByTable $rowsByTable -RecordId $treeId))
            }
        }
    }

    $regenRecords = New-Object System.Collections.Generic.List[object]
    $regenTable = Get-ActualTableName -Tables $tables -Name "RegenMeasurements"
    $regenCustomTable = Get-ActualTableName -Tables $tables -Name "RegenCustomMeasurements"
    if (-not [string]::IsNullOrWhiteSpace($regenTable) -and -not [string]::IsNullOrWhiteSpace($plotId)) {
        $regenColumns = @(Get-TableColumns -Connection $Connection -TableName $regenTable)
        if (Test-ColumnExists -Columns $regenColumns -Name "PlotID") {
            $regenWhere = New-HooterFieldEqualsPredicate -Connection $Connection -TableName $regenTable -FieldName "PlotID" -Value $plotId
            if (Test-ColumnExists -Columns $regenColumns -Name "IsDeleted") {
                $regenWhere += " AND ([IsDeleted] = False OR [IsDeleted] Is Null)"
            }
            $orderParts = New-Object System.Collections.Generic.List[string]
            foreach ($orderName in @("MinorPlot", "SpeciesCode", "IDBH")) {
                if (Test-ColumnExists -Columns $regenColumns -Name $orderName) { [void]$orderParts.Add((Quote-Name $orderName)) }
            }
            $regenRows = Get-HooterRowsByWhere -Connection $Connection -TableName $regenTable -Where $regenWhere -OrderBy ([string]::Join(", ", [string[]]$orderParts.ToArray()))
            foreach ($regenRow in $regenRows.Rows) {
                $regenCustomRow = if (-not [string]::IsNullOrWhiteSpace($regenCustomTable)) { Get-HooterCustomRow -Connection $Connection -TableName $regenCustomTable -MeasurementRow $regenRow -KeyNames @("RegenMeasKey") } else { $null }
                $rowsByTable = New-InsensitiveHashtable
                $rowsByTable[$regenTable] = $regenRow
                if ($null -ne $regenCustomRow) { $rowsByTable[$regenCustomTable] = $regenCustomRow }
                $minorPlot = Get-DataRowText -Row $regenRow -ColumnNames @("MinorPlot")
                $species = Get-DataRowText -Row $regenRow -ColumnNames @("SpeciesCode", "Species")
                $idbh = Get-DataRowText -Row $regenRow -ColumnNames @("IDBH")
                $displayParts = @()
                if (-not [string]::IsNullOrWhiteSpace($minorPlot)) { $displayParts += "MP $minorPlot" }
                if (-not [string]::IsNullOrWhiteSpace($species)) { $displayParts += $species }
                if (-not [string]::IsNullOrWhiteSpace($idbh)) { $displayParts += "IDBH $idbh" }
                $display = if ($displayParts.Count -gt 0) { "Regen " + ([string]::Join(" / ", [string[]]$displayParts)) } else { "Regen record $($regenRecords.Count + 1)" }
                [void]$regenRecords.Add((New-HooterRecord -Display $display -RowsByTable $rowsByTable -RecordId (Get-DataRowText -Row $regenRow -ColumnNames @("MeasurementID", "RegenMeasKey"))))
            }
        }
    }

    return [pscustomobject]@{
        PlotNumber = Get-DataRowText -Row $plotRow -ColumnNames @("PlotNumber", "PlotNo", "Plot")
        PlotID = $plotId
        UTMNorthing = Get-DataRowText -Row $plotRow -ColumnNames @("UTMNorthingCoordinate", "UTMNorthing", "Northing")
        UTMEasting = Get-DataRowText -Row $plotRow -ColumnNames @("UTMEastingCoordinate", "UTMEasting", "Easting")
        UTMZone = Get-DataRowText -Row $plotRow -ColumnNames @("UTMZone", "Zone")
        RowsByTable = $plotRows
        TreeRecords = @($treeRecords.ToArray())
        RegenRecords = @($regenRecords.ToArray())
    }
}

function Set-HooterDataRoot {
    param([string]$Path)

    $clean = ConvertTo-HooterText $Path
    if ([string]::IsNullOrWhiteSpace($clean)) { return }
    $script:DataRoot = $clean
    $script:SettingsPath = Join-Path $script:DataRoot "PlotHootSettings.json"
    if (-not (Test-Path -LiteralPath $script:SettingsPath)) {
        $fallbackSettings = Get-ChildItem -LiteralPath $script:DataRoot -Filter "Plot*Settings.json" -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $fallbackSettings) { $script:SettingsPath = $fallbackSettings.FullName }
    }
}

function Test-HooterWritableFolder {
    param([string]$Path)

    $clean = ConvertTo-HooterText $Path
    if ([string]::IsNullOrWhiteSpace($clean)) { return $false }
    try {
        if (-not (Test-Path -LiteralPath $clean)) {
            [void](New-Item -ItemType Directory -Force -Path $clean)
        }
        $testPath = Join-Path $clean ("PlotHoot_WriteTest_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
        "ok" | Set-Content -LiteralPath $testPath -Encoding ASCII
        $canWrite = Test-Path -LiteralPath $testPath
        try { [System.IO.File]::Delete($testPath) } catch { Remove-Item -LiteralPath $testPath -Force -ErrorAction SilentlyContinue }
        return $canWrite
    }
    catch {
        return $false
    }
}

function Ensure-HooterDataRoot {
    $candidates = New-Object System.Collections.Generic.List[string]
    [void]$candidates.Add((Join-Path $script:AppRoot "Data"))
    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
        [void]$candidates.Add((Join-Path $localAppData "PlotHoot\Data"))
    }
    if (-not [string]::IsNullOrWhiteSpace($script:DataRoot)) {
        [void]$candidates.Add($script:DataRoot)
    }

    $seen = New-InsensitiveHashtable
    foreach ($candidate in @($candidates.ToArray())) {
        $clean = ConvertTo-HooterText $candidate
        if ([string]::IsNullOrWhiteSpace($clean) -or $seen.ContainsKey($clean)) { continue }
        $seen[$clean] = $true
        if (Test-HooterWritableFolder -Path $clean) {
            Set-HooterDataRoot -Path $clean
            return
        }
    }

    throw "PlotHoot could not find a writable setup folder for this Windows account."
}

function Load-HooterSettings {
    $script:Tolerances = @{}
    $script:FieldOrder = @{}
    $script:ExcludedFieldKeys = @{}
    Reset-HooterStemToleranceBands
    $script:ScorePassPercent = $script:DefaultScorePassPercent
    $script:MaxPointLoss = $script:DefaultMaxPointLoss
    $script:PlotPointTotal = $script:DefaultPlotPointTotal
    $script:TreePointTotal = $script:DefaultTreePointTotal
    $script:RegenPointTotal = $script:DefaultRegenPointTotal
    $script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
    $script:UseStemCountPercentageForScoring = $script:DefaultUseStemCountPercentageForScoring
    Ensure-HooterDataRoot
    if (-not (Test-Path -LiteralPath $script:SettingsPath)) { return }
    $obsoleteExecutionSettingFound = $false
    try {
        $settings = Get-Content -Raw -LiteralPath $script:SettingsPath | ConvertFrom-Json
        if ($null -ne $settings.PSObject.Properties["ScorePassPercent"]) {
            $script:ScorePassPercent = Normalize-HooterScorePassPercent $settings.ScorePassPercent
        }
        if ($null -ne $settings.PSObject.Properties["MaxPointLoss"]) {
            $script:MaxPointLoss = Normalize-HooterMaxPointLoss $settings.MaxPointLoss
        }
        if ($null -ne $settings.PSObject.Properties["PlotPointTotal"]) {
            $script:PlotPointTotal = Normalize-HooterPointTotal $settings.PlotPointTotal
        }
        if ($null -ne $settings.PSObject.Properties["TreePointTotal"]) {
            $script:TreePointTotal = Normalize-HooterPointTotal $settings.TreePointTotal
        }
        if ($null -ne $settings.PSObject.Properties["RegenPointTotal"]) {
            $script:RegenPointTotal = Normalize-HooterPointTotal $settings.RegenPointTotal
        }
        if ($null -ne $settings.PSObject.Properties["ExecutionPointTotal"]) {
            $script:ExecutionPointTotal = Normalize-HooterPointTotal $settings.ExecutionPointTotal
        }
        if ($null -ne $settings.PSObject.Properties["UseStemCountPercentageForScoring"]) {
            $script:UseStemCountPercentageForScoring = ConvertTo-HooterBool $settings.UseStemCountPercentageForScoring
        }
        if ($null -ne $settings.PSObject.Properties["StemToleranceBands"]) {
            foreach ($band in @($settings.StemToleranceBands)) {
                [void](Set-HooterStemToleranceBand -Key $band.Key -Tolerance $band.Tolerance)
            }
        }
        if ($null -ne $settings.PSObject.Properties["FieldOrder"]) {
            foreach ($item in @($settings.FieldOrder)) {
                $fieldKey = ConvertTo-HooterText $item.FieldKey
                if (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey) {
                    $obsoleteExecutionSettingFound = $true
                    continue
                }
                $order = 0
                if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and [int]::TryParse((ConvertTo-HooterText $item.Order), [ref]$order)) {
                    $script:FieldOrder[$fieldKey] = $order
                }
            }
        }
        if ($null -ne $settings.PSObject.Properties["ExcludedFields"]) {
            foreach ($item in @($settings.ExcludedFields)) {
                $fieldKey = ConvertTo-HooterText $(if ($null -ne $item.PSObject.Properties["FieldKey"]) { $item.FieldKey } else { $item })
                if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and -not (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey)) {
                    $script:ExcludedFieldKeys[$fieldKey] = $true
                }
            }
        }
        foreach ($item in @($settings.Tolerances)) {
            $fieldKey = ConvertTo-HooterText $item.FieldKey
            if ([string]::IsNullOrWhiteSpace($fieldKey)) { continue }
            if (Test-HooterHiddenFieldKey -FieldKey $fieldKey) { continue }
            if (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey) {
                $obsoleteExecutionSettingFound = $true
                continue
            }
            $default = Get-HooterTolerance -Field $fieldKey
            $script:Tolerances[$fieldKey] = [pscustomobject]@{
                FieldKey = $fieldKey
                Group = ConvertTo-HooterText $item.Group
                TableName = ConvertTo-HooterText $item.TableName
                FieldName = ConvertTo-HooterText $item.FieldName
                Label = ConvertTo-HooterText $item.Label
                Mode = Normalize-HooterToleranceMode $item.Mode
                Value = ConvertTo-HooterText $item.Value
                PointValue = if ($null -ne $item.PSObject.Properties["PointValue"]) { ConvertTo-HooterText $item.PointValue } else { ConvertTo-HooterText $default.PointValue }
                CriticalFail = if ($null -ne $item.PSObject.Properties["CriticalFail"]) { ConvertTo-HooterBool $item.CriticalFail } else { Get-HooterCriticalFail -Tolerance $default }
            }
        }
        $executionTotalNumber = 0.0
        if ((ConvertTo-HooterNumber -Value $script:ExecutionPointTotal -Number ([ref]$executionTotalNumber)) -and
            [Math]::Abs($executionTotalNumber - 12.0) -lt 0.000001) {
            $script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
        }
    }
    catch {
        $script:Tolerances = @{}
        $script:FieldOrder = @{}
        $script:ExcludedFieldKeys = @{}
        $script:ScorePassPercent = $script:DefaultScorePassPercent
        $script:MaxPointLoss = $script:DefaultMaxPointLoss
        $script:PlotPointTotal = $script:DefaultPlotPointTotal
        $script:TreePointTotal = $script:DefaultTreePointTotal
        $script:RegenPointTotal = $script:DefaultRegenPointTotal
        $script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
        $script:UseStemCountPercentageForScoring = $script:DefaultUseStemCountPercentageForScoring
        Reset-HooterStemToleranceBands
    }
}

function Save-HooterSettings {
    Ensure-HooterDataRoot
    $items = @($script:Tolerances.Keys | Where-Object { -not (Test-HooterObsoleteExecutionFieldKey -FieldKey $_) -and -not (Test-HooterHiddenFieldKey -FieldKey $_) } | Sort-Object | ForEach-Object { $script:Tolerances[$_] })
    $fieldOrderItems = @($script:FieldOrder.Keys | Where-Object { -not (Test-HooterObsoleteExecutionFieldKey -FieldKey $_) -and -not (Test-HooterHiddenFieldKey -FieldKey $_) } | Sort-Object { [int]$script:FieldOrder[$_] } | ForEach-Object {
        [pscustomobject]@{
            FieldKey = $_
            Order = [int]$script:FieldOrder[$_]
        }
    })
    $excludedFieldItems = @($script:ExcludedFieldKeys.Keys | Where-Object { -not (Test-HooterObsoleteExecutionFieldKey -FieldKey $_) -and -not (Test-HooterHiddenFieldKey -FieldKey $_) } | Sort-Object | ForEach-Object {
        [pscustomobject]@{
            FieldKey = $_
        }
    })
    $settings = [pscustomobject]@{
        App = $script:AppName
        Version = $script:AppVersion
        SavedAt = (Get-Date).ToString("s")
        ScorePassPercent = Normalize-HooterScorePassPercent $script:ScorePassPercent
        MaxPointLoss = Normalize-HooterMaxPointLoss $script:MaxPointLoss
        PlotPointTotal = Normalize-HooterPointTotal $script:PlotPointTotal
        TreePointTotal = Normalize-HooterPointTotal $script:TreePointTotal
        RegenPointTotal = Normalize-HooterPointTotal $script:RegenPointTotal
        ExecutionPointTotal = Normalize-HooterPointTotal $script:ExecutionPointTotal
        UseStemCountPercentageForScoring = [bool]$script:UseStemCountPercentageForScoring
        StemToleranceBands = @(Get-HooterStemToleranceBands)
        FieldOrder = $fieldOrderItems
        ExcludedFields = $excludedFieldItems
        Tolerances = $items
    }
    $settings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
}

function Set-HooterAppIdentity {
    try {
        if (-not ([System.Management.Automation.PSTypeName]"PlotHoot.ShellAppIdentity").Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace PlotHoot {
    public static class ShellAppIdentity {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SetCurrentProcessExplicitAppUserModelID(string appID);
    }
}
"@
        }
        [void][PlotHoot.ShellAppIdentity]::SetCurrentProcessExplicitAppUserModelID($script:AppUserModelId)
    }
    catch {
    }
}

function Add-HooterAssemblies {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-HooterTouchKeyboardSupport
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

function Add-HooterTouchKeyboardSupport {
    try {
        if (-not ([System.Management.Automation.PSTypeName]"PlotHoot.TouchKeyboardInputScope").Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace PlotHoot {
    public static class TouchKeyboardInputScope {
        public const int Default = 0;
        public const int Number = 29;
        public const int Text = 57;

        [DllImport("msctf.dll", PreserveSig = true)]
        private static extern int SetInputScopes(IntPtr hwnd, int[] inputScopes, uint count, IntPtr phraseList, uint phraseCount, IntPtr regExp, IntPtr srgs);

        public static bool Set(IntPtr hwnd, int inputScope) {
            if (hwnd == IntPtr.Zero) { return false; }
            try {
                int[] scopes = new int[] { inputScope };
                return SetInputScopes(hwnd, scopes, 1, IntPtr.Zero, 0, IntPtr.Zero, IntPtr.Zero) >= 0;
            }
            catch {
                return false;
            }
        }
    }
}
"@
        }
    }
    catch {
    }
}

function Set-HooterInputScope {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Scope = "Text"
    )

    if ($null -eq $Control -or -not ([System.Management.Automation.PSTypeName]"PlotHoot.TouchKeyboardInputScope").Type) { return }
    $inputScope = switch ($Scope) {
        "Number" { [PlotHoot.TouchKeyboardInputScope]::Number; break }
        "Default" { [PlotHoot.TouchKeyboardInputScope]::Default; break }
        default { [PlotHoot.TouchKeyboardInputScope]::Text; break }
    }
    try {
        if (-not $Control.IsHandleCreated) { [void]$Control.CreateControl() }
        [void][PlotHoot.TouchKeyboardInputScope]::Set($Control.Handle, $inputScope)
    }
    catch {
    }
}

function Show-HooterTouchKeyboard {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Scope = "Text",
        [switch]$Force
    )

    if (-not $script:TouchKeyboardEnabled) { return $false }
    try {
        $now = Get-Date
        if ($null -ne $Control) {
            Set-HooterInputScope -Control $Control -Scope $Scope
            try { [void]$Control.Focus() } catch {}
        }
        if (($now - $script:LastTouchKeyboardRequest).TotalMilliseconds -lt 3500) { return $true }

        $paths = New-Object System.Collections.Generic.List[string]
        $programFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
        $programFilesX86 = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86)
        $windowsFolder = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
        if (-not [string]::IsNullOrWhiteSpace($programFiles)) {
            [void]$paths.Add((Join-Path $programFiles "Common Files\microsoft shared\ink\TabTip.exe"))
        }
        if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) {
            [void]$paths.Add((Join-Path $programFilesX86 "Common Files\microsoft shared\ink\TabTip.exe"))
        }
        if (-not [string]::IsNullOrWhiteSpace($windowsFolder)) {
            [void]$paths.Add((Join-Path $windowsFolder "System32\osk.exe"))
        }

        foreach ($path in @($paths.ToArray())) {
            if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
                Start-Process -FilePath $path -ErrorAction SilentlyContinue | Out-Null
                $script:LastTouchKeyboardRequest = $now
                return $true
            }
        }
    }
    catch {
    }
    return $false
}

function Request-HooterTouchKeyboardForControl {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Scope = "Text",
        [switch]$Force
    )

    if ($null -eq $Control) { return }
    [void](Show-HooterTouchKeyboard -Control $Control -Scope $Scope -Force:$Force)
    $launchedRecently = ((Get-Date) - $script:LastTouchKeyboardRequest).TotalMilliseconds -lt 700
    if ($Force -and $script:TouchKeyboardEnabled -and -not $launchedRecently) {
        try {
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 250
            $targetControl = $Control
            $targetScope = $Scope
            $timer.Add_Tick({
                param($TimerSender, $TimerEvent)
                try {
                    $TimerSender.Stop()
                    [void](Show-HooterTouchKeyboard -Control $targetControl -Scope $targetScope -Force)
                }
                finally {
                    try { $TimerSender.Dispose() } catch {}
                }
            }.GetNewClosure())
            $timer.Start()
        }
        catch {
        }
    }
}

function Register-HooterTextInputScope {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Scope = "Text"
    )

    if ($null -eq $Control) { return }
    $Control.Add_Enter({
        Set-HooterInputScope -Control $Control -Scope $Scope
        [void](Show-HooterTouchKeyboard -Control $Control -Scope $Scope)
    }.GetNewClosure())
    $Control.Add_GotFocus({
        Set-HooterInputScope -Control $Control -Scope $Scope
    }.GetNewClosure())
}

function Test-HooterRowNeedsNumericInput {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $Row.IsNewRow) { return $false }
    $fieldKey = if ($Row.DataGridView.Columns.Contains("FieldKey")) { ConvertTo-HooterText $Row.Cells["FieldKey"].Value } else { "" }
    $fieldName = if ($Row.DataGridView.Columns.Contains("FieldName")) { ConvertTo-HooterText $Row.Cells["FieldName"].Value } else { "" }
    $fieldLabel = if ($Row.DataGridView.Columns.Contains("FieldLabel")) { ConvertTo-HooterText $Row.Cells["FieldLabel"].Value } else { "" }
    $tolerance = if (-not [string]::IsNullOrWhiteSpace($fieldKey)) { Get-HooterTolerance -Field $fieldKey } else { $null }
    $mode = if ($null -ne $tolerance) { Normalize-HooterToleranceMode $tolerance.Mode } else { "" }
    if ($mode -in @("Range", "Percent", "Class", "StemCount", "StemPercent")) { return $true }
    if ($mode -in @("PassFail", "Filled")) { return $false }

    $nameKey = Get-NameKey ("$fieldName $fieldLabel")
    if ($nameKey -match "class|status|condition|history|problem|severity|remarks|notes|comment|type") { return $false }
    if ($nameKey -match "species|code|dbh|diameter|height|elevation|azimuth|distance|bearing|northing|easting|utm|zone|slope|percent|ratio|age|increment|defect|count|number|trees|seedling|sapling|stock|basal|acre|radius|length|width") { return $true }
    return $false
}

function Get-HooterGridInputScope {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid -or $null -eq $Grid.CurrentCell) { return "Text" }
    $columnName = $Grid.Columns[$Grid.CurrentCell.ColumnIndex].Name
    if ($columnName -in @("Notes", "FieldLabel", "CrewValue", "Rule", "Status", "Score", "TableName", "FieldName")) { return "Text" }
    if ($columnName -in @("DBH", "Distance", "Azimuth")) { return "Number" }
    if ($columnName -in @("PointValue", "MaxPointLoss", "PlotPointTotal", "TreePointTotal", "RegenPointTotal", "ExecutionPointTotal")) { return "Number" }
    if ($columnName -eq "ToleranceValue") {
        if ($Grid.Columns.Contains("StemBandKey")) { return "Number" }
        $mode = if ($Grid.Columns.Contains("Mode")) { Normalize-HooterToleranceMode $Grid.Rows[$Grid.CurrentCell.RowIndex].Cells["Mode"].Value } else { "" }
        if ($mode -in @("Range", "Percent", "Class")) { return "Number" }
        return "Text"
    }
    if ($columnName -eq "QaValue") {
        if (Test-HooterRowNeedsNumericInput -Row $Grid.Rows[$Grid.CurrentCell.RowIndex]) { return "Number" }
        return "Text"
    }
    return "Text"
}

function Test-HooterAutoAdvanceColumn {
    param([string]$ColumnName)

    return ($ColumnName -in @("QaValue", "Rating"))
}

function Test-HooterGridCellHasEntry {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$RowIndex,
        [string]$ColumnName
    )

    if ($null -eq $Grid -or $RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count -or -not $Grid.Columns.Contains($ColumnName)) { return $false }
    return -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $Grid.Rows[$RowIndex].Cells[$ColumnName].Value))
}

function Start-HooterGridCellEdit {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$RowIndex,
        [string]$ColumnName,
        [switch]$SelectAll,
        [switch]$ForceKeyboard
    )

    if ($null -eq $Grid -or $RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count -or -not $Grid.Columns.Contains($ColumnName)) { return $false }
    $row = $Grid.Rows[$RowIndex]
    if ($row.IsNewRow -or -not $row.Visible) { return $false }
    $cell = $row.Cells[$ColumnName]
    if ($cell.ReadOnly -or -not $cell.Visible) { return $false }

    try {
        $Grid.Focus()
        $Grid.ClearSelection()
        $Grid.CurrentCell = $cell
        $cell.Selected = $true
    }
    catch {
        return $false
    }

    if ($cell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) { return $true }

    $targetGrid = $Grid
    $targetCell = $cell
    $selectText = [bool]$SelectAll
    $forceKeyboardOpen = [bool]$ForceKeyboard
    $callback = {
        try {
            if ($null -ne $targetGrid.CurrentCell -and [object]::ReferenceEquals($targetGrid.CurrentCell, $targetCell)) {
                [void]$targetGrid.BeginEdit($true)
                $editor = $targetGrid.EditingControl
                if ($null -ne $editor) {
                    $scope = Get-HooterGridInputScope -Grid $targetGrid
                    Set-HooterInputScope -Control $editor -Scope $scope
                    if ($selectText -and ($editor -is [System.Windows.Forms.TextBoxBase])) {
                        $editor.SelectAll()
                    }
                    elseif ($editor -is [System.Windows.Forms.ComboBox]) {
                        $editor.DroppedDown = $true
                    }
                    if ($targetGrid.Columns[$targetGrid.CurrentCell.ColumnIndex].Name -eq "QaValue") {
                        Request-HooterTouchKeyboardForControl -Control $editor -Scope $scope -Force:$forceKeyboardOpen
                    }
                }
            }
        }
        catch {}
    }.GetNewClosure()

    try { [void]$Grid.BeginEdit($true) } catch {}
    try { [void]$Grid.BeginInvoke([System.Windows.Forms.MethodInvoker]$callback) } catch { & $callback }
    return $true
}

function Focus-HooterGridEntryCell {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$RowIndex,
        [string]$ColumnName,
        [switch]$ForceKeyboard
    )

    if ($null -eq $Grid -or $RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count -or -not $Grid.Columns.Contains($ColumnName)) { return $false }
    $row = $Grid.Rows[$RowIndex]
    if ($row.IsNewRow -or -not $row.Visible) { return $false }
    $cell = $row.Cells[$ColumnName]
    if ($cell.ReadOnly -or -not $cell.Visible) { return $false }
    try {
        if (-not (Start-HooterGridCellEdit -Grid $Grid -RowIndex $RowIndex -ColumnName $ColumnName -SelectAll -ForceKeyboard:$ForceKeyboard)) { return $false }
        if ($row.Visible) {
            try { $Grid.FirstDisplayedScrollingRowIndex = $RowIndex } catch {}
        }
        return $true
    }
    catch {
        return $false
    }
}

function Focus-HooterFirstVisibleEntryCell {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$ColumnName,
        [switch]$ForceKeyboard
    )

    if ($null -eq $Grid) { return $false }
    if ($null -ne $Grid.CurrentCell -and $Grid.CurrentCell.RowIndex -ge 0) {
        if (Focus-HooterGridEntryCell -Grid $Grid -RowIndex $Grid.CurrentCell.RowIndex -ColumnName $ColumnName -ForceKeyboard:$ForceKeyboard) { return $true }
    }
    $firstDisplayed = -1
    try { $firstDisplayed = [int]$Grid.FirstDisplayedScrollingRowIndex } catch {}
    if ($firstDisplayed -ge 0) {
        for ($index = $firstDisplayed; $index -lt $Grid.Rows.Count; $index++) {
            if (Focus-HooterGridEntryCell -Grid $Grid -RowIndex $index -ColumnName $ColumnName -ForceKeyboard:$ForceKeyboard) { return $true }
            if ($Grid.Rows[$index].Visible -eq $false -and $index -gt $firstDisplayed) { break }
        }
    }
    for ($index = 0; $index -lt $Grid.Rows.Count; $index++) {
        if (Focus-HooterGridEntryCell -Grid $Grid -RowIndex $index -ColumnName $ColumnName -ForceKeyboard:$ForceKeyboard) { return $true }
    }
    return $false
}

function Move-HooterToNextEntryCell {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$FromRowIndex,
        [string]$ColumnName
    )

    if ($null -eq $Grid -or -not (Test-HooterAutoAdvanceColumn -ColumnName $ColumnName) -or -not $Grid.Columns.Contains($ColumnName)) { return $false }
    $currentEntry = ""
    if ($FromRowIndex -ge 0 -and $FromRowIndex -lt $Grid.Rows.Count) {
        $currentEntry = Get-HooterDataEntryRowEntryNumber -Row $Grid.Rows[$FromRowIndex]
    }
    $limitScanToCurrentEntry = $false
    if ($script:Ui.ContainsKey("TreeGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.TreeGrid)) {
        $limitScanToCurrentEntry = (-not $script:Ui.ContainsKey("TreeSelectedOnlyCheck")) -or [bool]$script:Ui.TreeSelectedOnlyCheck.Checked
    }
    elseif ($script:Ui.ContainsKey("RegenGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.RegenGrid)) {
        $limitScanToCurrentEntry = (-not $script:Ui.ContainsKey("RegenSelectedOnlyCheck")) -or [bool]$script:Ui.RegenSelectedOnlyCheck.Checked
    }
    for ($index = $FromRowIndex + 1; $index -lt $Grid.Rows.Count; $index++) {
        if ($limitScanToCurrentEntry -and -not [string]::IsNullOrWhiteSpace($currentEntry) -and
            -not (Get-HooterDataEntryRowEntryNumber -Row $Grid.Rows[$index]).Equals($currentEntry, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        if (Focus-HooterGridEntryCell -Grid $Grid -RowIndex $index -ColumnName $ColumnName -ForceKeyboard) { return $true }
    }

    if ($script:Ui.ContainsKey("TreeGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.TreeGrid)) {
        $selectedOnly = (-not $script:Ui.ContainsKey("TreeSelectedOnlyCheck")) -or [bool]$script:Ui.TreeSelectedOnlyCheck.Checked
        if ($selectedOnly -and $script:Ui.ContainsKey("TreeRecordBox")) {
            $box = $script:Ui.TreeRecordBox
            if ($box.SelectedIndex -ge 0 -and $box.SelectedIndex -lt ($box.Items.Count - 1)) {
                $box.SelectedIndex = $box.SelectedIndex + 1
                Set-HooterTreeVisibleRows
                return (Focus-HooterFirstVisibleEntryCell -Grid $Grid -ColumnName $ColumnName -ForceKeyboard)
            }
        }
    }

    if ($script:Ui.ContainsKey("RegenGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.RegenGrid)) {
        $selectedOnly = (-not $script:Ui.ContainsKey("RegenSelectedOnlyCheck")) -or [bool]$script:Ui.RegenSelectedOnlyCheck.Checked
        if ($selectedOnly -and $script:Ui.ContainsKey("RegenRecordBox")) {
            $box = $script:Ui.RegenRecordBox
            if ($box.SelectedIndex -ge 0 -and $box.SelectedIndex -lt ($box.Items.Count - 1)) {
                $box.SelectedIndex = $box.SelectedIndex + 1
                Set-HooterRegenVisibleRows
                return (Focus-HooterFirstVisibleEntryCell -Grid $Grid -ColumnName $ColumnName -ForceKeyboard)
            }
        }
    }

    return $false
}

function Invoke-HooterAutoAdvanceFromCell {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$RowIndex,
        [string]$ColumnName
    )

    if ($null -eq $Grid -or -not (Test-HooterAutoAdvanceColumn -ColumnName $ColumnName)) { return }
    if (-not (Test-HooterGridCellHasEntry -Grid $Grid -RowIndex $RowIndex -ColumnName $ColumnName)) { return }
    $targetGrid = $Grid
    $targetRowIndex = $RowIndex
    $targetColumnName = $ColumnName
    $callback = {
        [void](Move-HooterToNextEntryCell -Grid $targetGrid -FromRowIndex $targetRowIndex -ColumnName $targetColumnName)
    }.GetNewClosure()
    try {
        [void]$Grid.BeginInvoke([System.Windows.Forms.MethodInvoker]$callback)
    }
    catch {
        [void](Move-HooterToNextEntryCell -Grid $targetGrid -FromRowIndex $targetRowIndex -ColumnName $targetColumnName)
    }
}

function Invoke-HooterAutoAdvanceFromEditingControl {
    param(
        [System.Windows.Forms.Control]$EditingControl,
        [System.Windows.Forms.KeyEventArgs]$KeyEvent
    )

    if ($null -eq $EditingControl -or $null -eq $KeyEvent) { return }
    if ($KeyEvent.KeyCode -notin @([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Tab)) { return }
    $grid = $null
    try { $grid = $EditingControl.EditingControlDataGridView } catch {}
    if ($null -eq $grid -or $null -eq $grid.CurrentCell) { return }
    $rowIndex = $grid.CurrentCell.RowIndex
    $columnName = $grid.Columns[$grid.CurrentCell.ColumnIndex].Name
    if (-not (Test-HooterAutoAdvanceColumn -ColumnName $columnName)) { return }
    $KeyEvent.Handled = $true
    $KeyEvent.SuppressKeyPress = $true
    try {
        if ($null -ne $grid.CurrentCell -and -not $grid.CurrentCell.ReadOnly -and $EditingControl.PSObject.Properties["Text"]) {
            $grid.CurrentCell.Value = $EditingControl.Text
        }
    }
    catch {}
    try { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}
    try { [void]$grid.EndEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch { try { [void]$grid.EndEdit() } catch {} }
    if (-not (Test-HooterGridCellHasEntry -Grid $grid -RowIndex $rowIndex -ColumnName $columnName)) { return }
    Invoke-HooterAutoAdvanceFromCell -Grid $grid -RowIndex $rowIndex -ColumnName $columnName
}

function Register-HooterGridInputScopes {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return }
    $Grid.Add_EditingControlShowing({
        param($Sender, $EventArgs)
        if ($null -eq $EventArgs.Control) { return }
        $scope = Get-HooterGridInputScope -Grid $Sender
        Set-HooterInputScope -Control $EventArgs.Control -Scope $scope
        if ($null -ne $Sender.CurrentCell -and $Sender.Columns[$Sender.CurrentCell.ColumnIndex].Name -eq "QaValue") {
            [void](Show-HooterTouchKeyboard -Control $EventArgs.Control -Scope $scope)
        }
        $handleKey = ""
        try { $handleKey = [string]$EventArgs.Control.Handle } catch {}
        if (-not [string]::IsNullOrWhiteSpace($handleKey) -and -not $script:AutoAdvanceEditingControlHandles.ContainsKey($handleKey)) {
            $script:AutoAdvanceEditingControlHandles[$handleKey] = $true
            $EventArgs.Control.Add_PreviewKeyDown({
                param($KeySender, $PreviewEvent)
                if ($PreviewEvent.KeyCode -in @([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Tab)) {
                    $PreviewEvent.IsInputKey = $true
                }
            })
            $EventArgs.Control.Add_KeyDown({
                param($KeySender, $KeyEvent)
                Invoke-HooterAutoAdvanceFromEditingControl -EditingControl $KeySender -KeyEvent $KeyEvent
            })
        }
    })
    $Grid.Add_KeyDown({
        param($Sender, $KeyEvent)
        if ($null -eq $Sender.CurrentCell) { return }
        if ($KeyEvent.KeyCode -notin @([System.Windows.Forms.Keys]::Enter, [System.Windows.Forms.Keys]::Tab)) { return }
        $rowIndex = $Sender.CurrentCell.RowIndex
        $columnName = $Sender.Columns[$Sender.CurrentCell.ColumnIndex].Name
        if (-not (Test-HooterAutoAdvanceColumn -ColumnName $columnName)) { return }
        $KeyEvent.Handled = $true
        $KeyEvent.SuppressKeyPress = $true
        try { [void]$Sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}
        try { [void]$Sender.EndEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch { try { [void]$Sender.EndEdit() } catch {} }
        if (-not (Test-HooterGridCellHasEntry -Grid $Sender -RowIndex $rowIndex -ColumnName $columnName)) { return }
        Invoke-HooterAutoAdvanceFromCell -Grid $Sender -RowIndex $rowIndex -ColumnName $columnName
    })
}

function New-HooterFont {
    param(
        [float]$Size = 9.0,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )
    return New-Object System.Drawing.Font("Segoe UI", $Size, $Style)
}

function New-HooterButton {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H = 34
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($W, $H)
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(29, 83, 79)
    $button.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 242)
    $button.ForeColor = [System.Drawing.Color]::FromArgb(16, 42, 40)
    $button.Font = New-HooterFont 9.2
    return $button
}

function New-HooterLabel {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H = 24,
        [float]$Size = 9.2,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($W, $H)
    $label.Font = New-HooterFont $Size $Style
    $label.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 42)
    return $label
}

function New-HooterTextBox {
    param(
        [int]$X,
        [int]$Y,
        [int]$W,
        [string]$Text = ""
    )

    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Size = New-Object System.Drawing.Size($W, 28)
    $box.Font = New-HooterFont 9.4
    $box.Text = $Text
    return $box
}

function Show-HooterSplashScreen {
    if (-not $script:SplashEnabled) { return $null }
    try {
        $splash = New-Object System.Windows.Forms.Form
        $splash.Text = "PlotHoot"
        $splash.StartPosition = "CenterScreen"
        $splash.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $splash.ShowInTaskbar = $false
        $splash.TopMost = $true
        $splash.Size = New-Object System.Drawing.Size(420, 230)
        $splash.BackColor = [System.Drawing.Color]::FromArgb(30, 71, 68)
        $splash.Font = New-HooterFont 9.4
        if (Test-Path -LiteralPath $script:AppIconPath) {
            try { $splash.Icon = New-Object System.Drawing.Icon($script:AppIconPath) } catch {}
        }

        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = "Fill"
        $panel.Padding = New-Object System.Windows.Forms.Padding(22)
        $panel.BackColor = [System.Drawing.Color]::FromArgb(30, 71, 68)
        $splash.Controls.Add($panel)

        if (Test-Path -LiteralPath $script:OwlImagePath) {
            $picture = New-Object System.Windows.Forms.PictureBox
            $picture.Location = New-Object System.Drawing.Point(24, 26)
            $picture.Size = New-Object System.Drawing.Size(94, 94)
            $picture.SizeMode = "Zoom"
            try {
                $image = [System.Drawing.Image]::FromFile($script:OwlImagePath)
                $picture.Image = $image
                $picture.Add_Disposed({ try { $image.Dispose() } catch {} }.GetNewClosure())
            }
            catch {}
            $panel.Controls.Add($picture)
        }

        $title = New-HooterLabel "PlotHoot" 142 30 240 36 20 ([System.Drawing.FontStyle]::Bold)
        $title.ForeColor = [System.Drawing.Color]::White
        $panel.Controls.Add($title)

        $message = New-HooterLabel "Loading offline QA tools..." 144 72 240 26 10.2
        $message.ForeColor = [System.Drawing.Color]::FromArgb(217, 236, 228)
        $panel.Controls.Add($message)

        $detail = New-HooterLabel "Preparing plot, tree, and regen checks" 144 98 240 24 8.8
        $detail.ForeColor = [System.Drawing.Color]::FromArgb(186, 214, 204)
        $panel.Controls.Add($detail)

        $bar = New-Object System.Windows.Forms.ProgressBar
        $bar.Location = New-Object System.Drawing.Point(24, 158)
        $bar.Size = New-Object System.Drawing.Size(348, 16)
        $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $bar.MarqueeAnimationSpeed = 28
        $panel.Controls.Add($bar)

        $splash.Show()
        $splash.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        return $splash
    }
    catch {
        return $null
    }
}

function Close-HooterSplashScreen {
    param([System.Windows.Forms.Form]$SplashForm)

    try {
        if ($null -ne $SplashForm -and -not $SplashForm.IsDisposed) {
            $SplashForm.Close()
            $SplashForm.Dispose()
        }
    }
    catch {
    }
}

function New-HooterGrid {
    param(
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point($X, $Y)
    $grid.Size = New-Object System.Drawing.Size($W, $H)
    $grid.Anchor = "Top,Bottom,Left,Right"
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToOrderColumns = $false
    $grid.AllowUserToResizeColumns = $false
    $grid.AllowUserToResizeRows = $false
    $grid.AutoSizeColumnsMode = "None"
    $grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = "FixedSingle"
    $grid.RowHeadersVisible = $false
    $grid.RowHeadersWidthSizeMode = [System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode]::DisableResizing
    $grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $true
    $grid.Font = New-HooterFont 9.0
    $grid.RowTemplate.Height = 28
    $grid.ColumnHeadersHeight = 30
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 71, 68)
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $grid.ColumnHeadersDefaultCellStyle.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
    $grid.EnableHeadersVisualStyles = $false
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 249)
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(207, 229, 220)
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(18, 34, 32)
    $grid.Add_SortCompare({
        param($Sender, $EventArgs)
        $leftText = ConvertTo-HooterText $EventArgs.CellValue1
        $rightText = ConvertTo-HooterText $EventArgs.CellValue2
        $leftNumber = 0.0
        $rightNumber = 0.0
        if ((ConvertTo-HooterNumber -Value $leftText -Number ([ref]$leftNumber)) -and
            (ConvertTo-HooterNumber -Value $rightText -Number ([ref]$rightNumber))) {
            $EventArgs.SortResult = $leftNumber.CompareTo($rightNumber)
        }
        else {
            $EventArgs.SortResult = [string]::Compare($leftText, $rightText, $true, [System.Globalization.CultureInfo]::CurrentCulture)
        }
        $EventArgs.Handled = $true
    })
    return $grid
}

function Enable-HooterVisibleGridScrollBars {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return }
    $Grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $showScrollBars = {
        param($Sender, $EventArgs)
        if ($null -eq $Sender) { return }
        try { $Sender.ScrollBars = [System.Windows.Forms.ScrollBars]::Both } catch {}
        foreach ($control in $Sender.Controls) {
            if ($control -is [System.Windows.Forms.VScrollBar] -or $control -is [System.Windows.Forms.HScrollBar]) {
                try { $control.Visible = $true } catch {}
                try { $control.BringToFront() } catch {}
            }
        }
    }
    $Grid.Add_HandleCreated($showScrollBars)
    $Grid.Add_Resize($showScrollBars)
    $Grid.Add_RowsAdded($showScrollBars)
    $Grid.Add_RowsRemoved($showScrollBars)
    $Grid.Add_ColumnAdded($showScrollBars)
    $Grid.Add_ColumnWidthChanged($showScrollBars)
}

function Set-HooterScrollableTab {
    param(
        [System.Windows.Forms.TabPage]$Page,
        [int]$MinimumScrollHeight = 680
    )

    if ($null -eq $Page) { return }
    $Page.AutoScroll = $false
    $Page.AutoScrollMargin = New-Object System.Drawing.Size(0, 0)
    $Page.AutoScrollMinSize = New-Object System.Drawing.Size(0, 0)
}

function Set-HooterControlBounds {
    param(
        [System.Windows.Forms.Control]$Control,
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H
    )

    if ($null -eq $Control) { return }
    $Control.Bounds = New-HooterRectangle -X $X -Y $Y -Width $W -Height $H
}

function New-HooterRectangle {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )

    return New-Object System.Drawing.Rectangle -ArgumentList @($X, $Y, [Math]::Max(1, $Width), [Math]::Max(1, $Height))
}

function Set-HooterDataEntryGridBounds {
    param(
        [System.Windows.Forms.TabPage]$Page,
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$Top,
        [int]$Margin = 10
    )

    if ($null -eq $Page -or $null -eq $Grid) { return }
    $width = [Math]::Max(260, $Page.ClientSize.Width - ($Margin * 2))
    $height = [Math]::Max(180, $Page.ClientSize.Height - $Top - $Margin)
    Set-HooterControlBounds -Control $Grid -X $Margin -Y $Top -W $width -H $height
    try {
        $Grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
        $Grid.Invalidate()
    }
    catch {}
}

function Set-HooterTabHeaderColor {
    param(
        [System.Windows.Forms.TabPage]$Page,
        [System.Drawing.Color]$Color,
        [System.Drawing.Color]$TextColor = [System.Drawing.Color]::White
    )

    if ($null -eq $Page) { return }
    $Page.Tag = [pscustomobject]@{
        TabColor = $Color
        TextColor = $TextColor
    }
}

function Register-HooterColoredTabHeaders {
    param([System.Windows.Forms.TabControl]$TabControl)

    if ($null -eq $TabControl) { return }
    $TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.Padding = New-Object System.Drawing.Point(14, 5)
    $TabControl.Add_DrawItem({
        param($Sender, $EventArgs)

        if ($EventArgs.Index -lt 0 -or $EventArgs.Index -ge $Sender.TabPages.Count) { return }
        $page = $Sender.TabPages[$EventArgs.Index]
        $style = $page.Tag
        $tabColor = [System.Drawing.Color]::FromArgb(219, 226, 224)
        $textColor = [System.Drawing.Color]::FromArgb(20, 40, 38)
        if ($null -ne $style -and $null -ne $style.PSObject.Properties["TabColor"]) {
            $tabColor = [System.Drawing.Color]$style.TabColor
        }
        if ($null -ne $style -and $null -ne $style.PSObject.Properties["TextColor"]) {
            $textColor = [System.Drawing.Color]$style.TextColor
        }

        $bounds = $Sender.GetTabRect($EventArgs.Index)
        $selected = ($EventArgs.Index -eq $Sender.SelectedIndex)
        if ($selected) {
            $bounds = New-HooterRectangle -X $bounds.X -Y ([Math]::Max(0, $bounds.Y - 2)) -Width $bounds.Width -Height ($bounds.Height + 2)
        }
        $bounds.Inflate(-1, -1)

        $fillColor = $tabColor
        $labelColor = $textColor
        if ($selected) {
            $fillColor = [System.Drawing.Color]::FromArgb(250, 253, 251)
            $labelColor = [System.Drawing.Color]::FromArgb(13, 55, 49)
        }

        $brush = New-Object System.Drawing.SolidBrush($fillColor)
        $borderPen = New-Object System.Drawing.Pen($(if ($selected) { [System.Drawing.Color]::FromArgb(25, 45, 42) } else { [System.Drawing.Color]::FromArgb(180, 190, 186) }), $(if ($selected) { 2 } else { 1 }))
        $font = $Sender.Font
        $disposeFont = $false
        if ($selected) {
            $font = New-Object System.Drawing.Font($Sender.Font, [System.Drawing.FontStyle]::Bold)
            $disposeFont = $true
        }
        try {
            $EventArgs.Graphics.FillRectangle($brush, $bounds)
            $EventArgs.Graphics.DrawRectangle($borderPen, $bounds.X, $bounds.Y, [Math]::Max(1, $bounds.Width - 1), [Math]::Max(1, $bounds.Height - 1))
            $textTop = $bounds.Y + 2
            $textBounds = New-HooterRectangle -X ($bounds.X + 4) -Y $textTop -Width ($bounds.Width - 8) -Height ($bounds.Bottom - $textTop - 2)
            $flags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
            [System.Windows.Forms.TextRenderer]::DrawText($EventArgs.Graphics, $page.Text, $font, $textBounds, $labelColor, $flags)
        }
        finally {
            if ($disposeFont) { $font.Dispose() }
            $borderPen.Dispose()
            $brush.Dispose()
        }
    })
}

function Add-GridTextColumn {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Name,
        [string]$Header,
        [int]$Width,
        [bool]$ReadOnly = $false,
        [bool]$Visible = $true
    )

    $index = $Grid.Columns.Add($Name, $Header)
    $Grid.Columns[$index].Width = $Width
    $Grid.Columns[$index].ReadOnly = $ReadOnly
    $Grid.Columns[$index].Visible = $Visible
    $Grid.Columns[$index].Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    $Grid.Columns[$index].SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    return $Grid.Columns[$index]
}

function Set-HooterGridColumnToolTip {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$ColumnName,
        [string]$Text
    )

    if ($null -eq $Grid -or -not $Grid.Columns.Contains($ColumnName)) { return }
    $Grid.Columns[$ColumnName].ToolTipText = $Text
    $Grid.Columns[$ColumnName].HeaderCell.ToolTipText = $Text
}

function Enable-HooterGridTapSort {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [scriptblock]$AfterSort = $null
    )

    if ($null -eq $Grid) { return }
    foreach ($column in $Grid.Columns) {
        try { $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic } catch {}
    }
    $Grid.Add_ColumnHeaderMouseClick({
        param($Sender, $EventArgs)
        if ($EventArgs.ColumnIndex -lt 0) { return }
        $column = $Sender.Columns[$EventArgs.ColumnIndex]
        if ($null -eq $column -or -not $column.Visible) { return }

        $direction = [System.ComponentModel.ListSortDirection]::Ascending
        $glyph = [System.Windows.Forms.SortOrder]::Ascending
        if ($column.HeaderCell.SortGlyphDirection -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $direction = [System.ComponentModel.ListSortDirection]::Descending
            $glyph = [System.Windows.Forms.SortOrder]::Descending
        }

        try { [void]$Sender.EndEdit() } catch {}
        try { $Sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}
        try { $Sender.CurrentCell = $null } catch {}
        foreach ($candidate in $Sender.Columns) {
            try { $candidate.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None } catch {}
        }
        try {
            $Sender.Sort($column, $direction)
            $column.HeaderCell.SortGlyphDirection = $glyph
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("PlotHoot could not sort this column. Try another column or clear the current cell first.", "Sort column", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        if ($null -ne $AfterSort) {
            & $AfterSort
        }
    }.GetNewClosure())
}

function Add-HooterCheckColumns {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [bool]$WithEntry
    )

    $moveColumn = Add-GridTextColumn -Grid $Grid -Name "MoveHandle" -Header "Move" -Width 50 -ReadOnly $true
    $moveColumn.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $moveColumn.DefaultCellStyle.Font = New-HooterFont 12.0 ([System.Drawing.FontStyle]::Bold)
    if ($WithEntry) {
        [void](Add-GridTextColumn -Grid $Grid -Name "EntryNumber" -Header "Entry" -Width 58 -ReadOnly $true)
        [void](Add-GridTextColumn -Grid $Grid -Name "CrewRecord" -Header "Crew record" -Width 150 -ReadOnly $true)
    }
    [void](Add-GridTextColumn -Grid $Grid -Name "FieldLabel" -Header "Field" -Width 210 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "CrewValue" -Header "Crew value" -Width 145 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "QaValue" -Header "QA value" -Width 145 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Rule" -Header "Rule" -Width 180 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "PointValue" -Header "Points" -Width 75 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Status" -Header "Status" -Width 85 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "Score" -Header "Point loss" -Width 100 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "Notes" -Header "Field notes" -Width 380 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "FieldKey" -Header "FieldKey" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Group" -Header "Group" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "TableName" -Header "Table" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "FieldName" -Header "FieldName" -Width 80 -ReadOnly $true -Visible $false)
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "MoveHandle" -Text "Drag this handle up or down to move this field."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "EntryNumber" -Text "Tree or regen entry number for this QA record."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "CrewRecord" -Text "Crew/database record that this QA row is checking."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "FieldLabel" -Text "Field being checked."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "CrewValue" -Text "Value recorded by the field crew in the database."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "QaValue" -Text "Enter the QA cruiser value here."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Rule" -Text "Tolerance rule used to compare crew and QA values."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "PointValue" -Text "Point loss for this field. Starts from AppColumns.QApoints and can be edited for the current QA work."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Status" -Text "Pass, Fail, Not checked, or Incomplete result for this row."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Score" -Text "Point loss for this row. Critical fail rows force the whole plot to fail."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Notes" -Text "Optional notes about this field check."
}

function Get-HooterExecutionScoreRule {
    param([string]$ItemKey)

    $base = Get-HooterExecutionBaseItem -ItemKey $ItemKey
    if ($null -eq $base) {
        $base = [pscustomobject]@{ ItemKey = $ItemKey; Item = $ItemKey; FairLoss = 1; PoorLoss = 2; CriticalOnPoor = $false; DefaultOrder = 0 }
    }

    $fieldKey = Get-HooterExecutionFieldKey -ItemKey $base.ItemKey
    $tolerance = Get-HooterTolerance -Field $fieldKey
    $fairLoss = 0.0
    $poorLoss = 0.0
    [void](ConvertTo-HooterNumber -Value $base.FairLoss -Number ([ref]$fairLoss))
    [void](ConvertTo-HooterNumber -Value $base.PoorLoss -Number ([ref]$poorLoss))

    $value = ConvertTo-HooterText $tolerance.Value
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        $pairMatches = [regex]::Matches($value, "(?i)\b(fair|poor)\s*[:=]\s*([+-]?\d+(?:\.\d+)?)")
        foreach ($match in $pairMatches) {
            $number = 0.0
            if (ConvertTo-HooterNumber -Value $match.Groups[2].Value -Number ([ref]$number)) {
                if ($match.Groups[1].Value.Equals("fair", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $fairLoss = $number
                }
                elseif ($match.Groups[1].Value.Equals("poor", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $poorLoss = $number
                }
            }
        }
        if ($pairMatches.Count -eq 0) {
            $numberMatches = [regex]::Matches($value, "[+-]?\d+(?:\.\d+)?")
            if ($numberMatches.Count -ge 2) {
                [void](ConvertTo-HooterNumber -Value $numberMatches[0].Value -Number ([ref]$fairLoss))
                [void](ConvertTo-HooterNumber -Value $numberMatches[1].Value -Number ([ref]$poorLoss))
            }
            elseif ($numberMatches.Count -eq 1) {
                [void](ConvertTo-HooterNumber -Value $numberMatches[0].Value -Number ([ref]$poorLoss))
            }
        }
    }

    $fairLoss = [Math]::Max(0.0, $fairLoss)
    $poorLoss = [Math]::Max(0.0, $poorLoss)
    $possiblePoints = [Math]::Max($fairLoss, $poorLoss)

    return [pscustomobject]@{
        ItemKey = $base.ItemKey
        Item = $base.Item
        FairLoss = $fairLoss
        PoorLoss = $poorLoss
        PossiblePoints = $possiblePoints
        CriticalOnPoor = Get-HooterCriticalFail -Tolerance $tolerance
        Tolerance = $tolerance
        ToleranceValue = Format-HooterExecutionToleranceValue -FairLoss $fairLoss -PoorLoss $poorLoss
    }
}

function Get-HooterExecutionItems {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($base in Get-HooterExecutionBaseItems) {
        $rule = Get-HooterExecutionScoreRule -ItemKey $base.ItemKey
        [void]$items.Add([pscustomobject]@{
            ItemKey = $base.ItemKey
            Item = $base.Item
            FieldKey = Get-HooterExecutionFieldKey -ItemKey $base.ItemKey
            FairLoss = $rule.FairLoss
            MaxLoss = $rule.PoorLoss
            PoorLoss = $rule.PoorLoss
            CriticalOnPoor = $rule.CriticalOnPoor
            DefaultOrder = $base.DefaultOrder
            ToleranceValue = $rule.ToleranceValue
        })
    }
    return @(Sort-HooterFieldsBySavedOrder -Fields @($items.ToArray()))
}

function Add-HooterExecutionColumns {
    param([System.Windows.Forms.DataGridView]$Grid)

    [void](Add-GridTextColumn -Grid $Grid -Name "Item" -Header "Item" -Width 620 -ReadOnly $true)
    $ratingColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $ratingColumn.Name = "Rating"
    $ratingColumn.HeaderText = "Rating"
    $ratingColumn.Width = 110
    $ratingColumn.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    [void]$ratingColumn.Items.Add("")
    [void]$ratingColumn.Items.Add("Good")
    [void]$ratingColumn.Items.Add("Fair")
    [void]$ratingColumn.Items.Add("Poor")
    $ratingColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    [void]$Grid.Columns.Add($ratingColumn)
    [void](Add-GridTextColumn -Grid $Grid -Name "PointLoss" -Header "Point loss" -Width 95 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "Status" -Header "Status" -Width 95 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "Notes" -Header "Notes" -Width 360 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "ItemKey" -Header "ItemKey" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "MaxLoss" -Header "MaxLoss" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "CriticalOnPoor" -Header "CriticalOnPoor" -Width 80 -ReadOnly $true -Visible $false)
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Item" -Text "Table D location/execution item."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Rating" -Text "Good/Fair/Poor point loss is controlled by this item in Tolerance Setup."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "PointLoss" -Text "Point loss for this Table D item. Poor can also be marked critical in Tolerance Setup."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Status" -Text "Pass or Fail result for this Table D item."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Notes" -Text "Optional notes about the location/execution check."
}

function Add-HooterMissedTreeColumns {
    param([System.Windows.Forms.DataGridView]$Grid)

    [void](Add-GridTextColumn -Grid $Grid -Name "EntryNumber" -Header "Missed tree" -Width 95 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $Grid -Name "DBH" -Header "DBH" -Width 90 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Distance" -Header "Distance to plot center" -Width 170 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Azimuth" -Header "Azimuth to plot center" -Width 170 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Notes" -Header "Missed tree notes" -Width 520 -ReadOnly $false)
    [void](Add-GridTextColumn -Grid $Grid -Name "Status" -Header "Status" -Width 160 -ReadOnly $true)
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "EntryNumber" -Text "Missed tree number for this plot QA check."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "DBH" -Text "DBH for the tree found by the QA cruiser but missed by the crew."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Distance" -Text "Distance from the geometric center of the tree to plot center."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Azimuth" -Text "Azimuth from the tree to plot center."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Notes" -Text "Notes about the missed tree."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "Status" -Text "Crew-missed trees force an automatic plot failure."
}

function Get-HooterNextMissedTreeNumber {
    if (-not $script:Ui.ContainsKey("MissedTreeGrid")) { return 1 }
    $max = 0
    foreach ($row in $script:Ui.MissedTreeGrid.Rows) {
        if ($row.IsNewRow) { continue }
        $entry = 0
        if ([int]::TryParse((ConvertTo-HooterText $row.Cells["EntryNumber"].Value), [ref]$entry)) {
            $max = [Math]::Max($max, $entry)
        }
    }
    return ($max + 1)
}

function Add-HooterMissedTreeRow {
    param(
        [object]$DBH = "",
        [object]$Distance = "",
        [object]$Azimuth = "",
        [object]$Notes = "",
        [object]$EntryNumber = ""
    )

    if (-not $script:Ui.ContainsKey("MissedTreeGrid")) { return }
    $grid = $script:Ui.MissedTreeGrid
    $entry = ConvertTo-HooterText $EntryNumber
    if ([string]::IsNullOrWhiteSpace($entry)) { $entry = [string](Get-HooterNextMissedTreeNumber) }
    $rowIndex = $grid.Rows.Add($entry, (ConvertTo-HooterText $DBH), (ConvertTo-HooterText $Distance), (ConvertTo-HooterText $Azimuth), (ConvertTo-HooterText $Notes), "Critical fail")
    if ($rowIndex -ge 0) {
        $grid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 242)
    }
    if ($script:Ui.ContainsKey("CrewMissedTreeCheck")) {
        $script:Ui.CrewMissedTreeCheck.Checked = $true
    }
    if (-not $script:SuppressMissedTreeCheckEvent) {
        Show-HooterAutoSaveNotice -Text "Updating missed tree check..."
        Request-HooterMissedTreeRefresh
    }
}

function Remove-HooterSelectedMissedTrees {
    if (-not $script:Ui.ContainsKey("MissedTreeGrid")) { return }
    $grid = $script:Ui.MissedTreeGrid
    if ($grid.SelectedRows.Count -eq 0 -and $null -ne $grid.CurrentRow) {
        $grid.CurrentRow.Selected = $true
    }
    for ($i = $grid.SelectedRows.Count - 1; $i -ge 0; $i--) {
        $row = $grid.SelectedRows[$i]
        if ($row.IsNewRow) { continue }
        $grid.Rows.Remove($row)
    }
    Show-HooterAutoSaveNotice -Text "Updating missed tree check..."
    Request-HooterMissedTreeRefresh
}

function Update-HooterMissedTreeStatus {
    if ($script:Ui.ContainsKey("MissedTreeGrid")) {
        foreach ($row in $script:Ui.MissedTreeGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $row.Cells["Status"].Value = "Critical fail"
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 242)
        }
    }
    if ($script:Ui.ContainsKey("MissedTreeStatusLabel")) {
        $count = Get-HooterMissedTreeRowCount
        $flag = Test-HooterCrewMissedTreeFail
        if ($flag) {
            $script:Ui.MissedTreeStatusLabel.Text = "Automatic plot fail: crew missed tree(s) found by QA. Logged missed trees: $count."
            $script:Ui.MissedTreeStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(170, 54, 35)
        }
        else {
            $script:Ui.MissedTreeStatusLabel.Text = "No crew-missed trees logged."
            $script:Ui.MissedTreeStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 75, 71)
        }
    }
}

function Populate-HooterExecutionGrid {
    if (-not $script:Ui.ContainsKey("ExecutionGrid")) { return }
    $grid = $script:Ui.ExecutionGrid
    $state = @{}
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $itemKey = ConvertTo-HooterText $row.Cells["ItemKey"].Value
        if ([string]::IsNullOrWhiteSpace($itemKey)) { continue }
        $state[$itemKey] = [pscustomobject]@{
            Rating = ConvertTo-HooterText $row.Cells["Rating"].Value
            Notes = ConvertTo-HooterText $row.Cells["Notes"].Value
        }
    }
    $grid.Rows.Clear()
    foreach ($item in Get-HooterExecutionItems) {
        $rowIndex = $grid.Rows.Add($item.Item, "", "", "", "", $item.ItemKey, $item.MaxLoss, $item.CriticalOnPoor)
        if ($state.ContainsKey($item.ItemKey)) {
            $grid.Rows[$rowIndex].Cells["Rating"].Value = $state[$item.ItemKey].Rating
            $grid.Rows[$rowIndex].Cells["Notes"].Value = $state[$item.ItemKey].Notes
        }
    }
    Refresh-HooterExecutionGridStatuses
}

function Get-HooterExecutionRowResult {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    $rating = ConvertTo-HooterText $Row.Cells["Rating"].Value
    $itemKey = ConvertTo-HooterText $Row.Cells["ItemKey"].Value
    $rule = Get-HooterExecutionScoreRule -ItemKey $itemKey
    $fairLoss = $rule.FairLoss
    $poorLoss = $rule.PoorLoss
    $maxLoss = $rule.PossiblePoints
    $criticalOnPoor = [bool]$rule.CriticalOnPoor
    $loss = 0.0
    $status = ""
    $display = ""
    $criticalFailure = $false

    switch ($rating) {
        "Good" {
            $status = "Pass"
            $display = "0"
        }
        "Fair" {
            $loss = $fairLoss
            $status = if ($loss -gt 0) { "Fail" } else { "Pass" }
            $display = Format-HooterScoreNumber $loss
        }
        "Poor" {
            $loss = $poorLoss
            $criticalFailure = $criticalOnPoor
            $status = if ($criticalFailure -or $loss -gt 0) { "Fail" } else { "Pass" }
            $display = if ($criticalFailure) { "Critical fail" } else { (Format-HooterScoreNumber $loss) }
        }
        default {
            $display = ""
        }
    }

    return [pscustomobject]@{
        Rating = $rating
        Status = $status
        Checked = -not [string]::IsNullOrWhiteSpace($rating)
        PointLoss = $loss
        PossiblePoints = $maxLoss
        EarnedPoints = if ([string]::IsNullOrWhiteSpace($rating)) { 0.0 } else { [Math]::Max(0.0, $maxLoss - $loss) }
        CriticalFailure = $criticalFailure
        Display = $display
    }
}

function Update-HooterExecutionRowStatus {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $Row.IsNewRow) { return }
    $result = Get-HooterExecutionRowResult -Row $Row
    $Row.Cells["PointLoss"].Value = $result.Display
    $Row.Cells["Status"].Value = $result.Status
    switch ($result.Status) {
        "Pass" { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(236, 249, 241) }
        "Fail" { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 230) }
        default { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White }
    }
}

function Refresh-HooterExecutionGridStatuses {
    if (-not $script:Ui.ContainsKey("ExecutionGrid")) { return }
    foreach ($row in $script:Ui.ExecutionGrid.Rows) {
        Update-HooterExecutionRowStatus -Row $row
    }
}

function Get-HooterRuleText {
    param([object]$Tolerance)

    $mode = Normalize-HooterToleranceMode $Tolerance.Mode
    $value = ConvertTo-HooterText $Tolerance.Value
    $points = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $Tolerance)
    $criticalText = if (Get-HooterCriticalFail -Tolerance $Tolerance) { "; critical" } else { "" }
    $scoreText = "; $points pt$criticalText"
    switch ($mode) {
        "Range" { return "+/- $value$scoreText" }
        "Percent" { return "$value%$scoreText" }
        "Class" { return "+/- $value class$scoreText" }
        "StemCount" { return "Stem count table$scoreText" }
        "StemPercent" { return "$(Get-HooterStemPercentRuleText -Tolerance $Tolerance)" }
        "PassFail" { return "Pass/fail$scoreText" }
        "Filled" { return "Filled in$scoreText" }
        "GoodFairPoor" { return "Good/Fair/Poor $value$scoreText" }
        default { return "Exact$scoreText" }
    }
}

function Copy-HooterTolerance {
    param([object]$Tolerance)

    if ($null -eq $Tolerance) {
        return [pscustomobject]@{ FieldKey = ""; Group = ""; TableName = ""; FieldName = ""; Label = ""; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
    }
    return [pscustomobject]@{
        FieldKey = if ($null -ne $Tolerance.PSObject.Properties["FieldKey"]) { ConvertTo-HooterText $Tolerance.FieldKey } else { "" }
        Group = if ($null -ne $Tolerance.PSObject.Properties["Group"]) { ConvertTo-HooterText $Tolerance.Group } else { "" }
        TableName = if ($null -ne $Tolerance.PSObject.Properties["TableName"]) { ConvertTo-HooterText $Tolerance.TableName } else { "" }
        FieldName = if ($null -ne $Tolerance.PSObject.Properties["FieldName"]) { ConvertTo-HooterText $Tolerance.FieldName } else { "" }
        Label = if ($null -ne $Tolerance.PSObject.Properties["Label"]) { ConvertTo-HooterText $Tolerance.Label } else { "" }
        Mode = if ($null -ne $Tolerance.PSObject.Properties["Mode"]) { Normalize-HooterToleranceMode $Tolerance.Mode } else { "Exact" }
        Value = if ($null -ne $Tolerance.PSObject.Properties["Value"]) { ConvertTo-HooterText $Tolerance.Value } else { "" }
        PointValue = if ($null -ne $Tolerance.PSObject.Properties["PointValue"]) { ConvertTo-HooterText $Tolerance.PointValue } else { "1" }
        CriticalFail = if ($null -ne $Tolerance.PSObject.Properties["CriticalFail"]) { ConvertTo-HooterBool $Tolerance.CriticalFail } else { $false }
    }
}

function Get-HooterRowTolerance {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $Row.IsNewRow) { return $null }
    $fieldKey = ConvertTo-HooterText $Row.Cells["FieldKey"].Value
    $tolerance = Copy-HooterTolerance (Get-HooterTolerance -Field $fieldKey)
    if ($Row.DataGridView.Columns.Contains("PointValue")) {
        $rowPointValue = Normalize-HooterQaPointValue $Row.Cells["PointValue"].Value
        if (-not [string]::IsNullOrWhiteSpace($rowPointValue)) {
            Set-HooterTolerancePointValue -Tolerance $tolerance -PointValue $rowPointValue
        }
    }
    return $tolerance
}

function Set-HooterWorkingPointOverride {
    param(
        [string]$FieldKey,
        [object]$PointValue
    )

    $fieldKeyText = ConvertTo-HooterText $FieldKey
    if ([string]::IsNullOrWhiteSpace($fieldKeyText)) { return "" }
    $points = Normalize-HooterQaPointValue $PointValue
    if ([string]::IsNullOrWhiteSpace($points)) {
        $points = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance (Get-HooterTolerance -Field $fieldKeyText))
    }

    $tolerance = Get-HooterTolerance -Field $fieldKeyText
    Set-HooterTolerancePointValue -Tolerance $tolerance -PointValue $points
    $script:Tolerances[$fieldKeyText] = $tolerance

    $script:SuppressGridEvents = $true
    try {
        foreach ($grid in @($script:Ui.PlotGrid, $script:Ui.TreeGrid, $script:Ui.RegenGrid)) {
            if ($null -eq $grid -or -not $grid.Columns.Contains("FieldKey") -or -not $grid.Columns.Contains("PointValue")) { continue }
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) { continue }
                if ((ConvertTo-HooterText $row.Cells["FieldKey"].Value) -ne $fieldKeyText) { continue }
                $row.Cells["PointValue"].Value = $points
                $row.Cells["Rule"].Value = Get-HooterRuleText (Get-HooterRowTolerance -Row $row)
            }
        }
        if ($script:Ui.ContainsKey("SettingsGrid")) {
            $settingsGrid = $script:Ui.SettingsGrid
            if ($settingsGrid.Columns.Contains("FieldKey") -and $settingsGrid.Columns.Contains("PointValue")) {
                foreach ($row in $settingsGrid.Rows) {
                    if ($row.IsNewRow) { continue }
                    if ((ConvertTo-HooterText $row.Cells["FieldKey"].Value) -eq $fieldKeyText) {
                        $row.Cells["PointValue"].Value = $points
                    }
                }
            }
        }
    }
    finally {
        $script:SuppressGridEvents = $false
    }

    foreach ($grid in @($script:Ui.PlotGrid, $script:Ui.TreeGrid, $script:Ui.RegenGrid)) {
        Refresh-HooterGridStatuses -Grid $grid
    }
    Request-HooterTreeProgressRefresh
    Request-HooterRegenProgressRefresh
    Request-HooterOverallRefresh
    return $points
}

function Update-HooterWorkingPointOverrideFromRow {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $Row.IsNewRow -or -not $Row.DataGridView.Columns.Contains("PointValue")) { return }
    $fieldKey = ConvertTo-HooterText $Row.Cells["FieldKey"].Value
    $points = Set-HooterWorkingPointOverride -FieldKey $fieldKey -PointValue $Row.Cells["PointValue"].Value
    if (-not [string]::IsNullOrWhiteSpace($points)) {
        $Row.Cells["PointValue"].Value = $points
    }
}

function Set-HooterFilledCheckCell {
    param(
        [System.Windows.Forms.DataGridViewRow]$Row,
        [object]$Tolerance
    )

    if ($null -eq $Row -or $null -eq $Row.DataGridView -or -not $Row.DataGridView.Columns.Contains("QaValue")) { return }
    if ((ConvertTo-HooterText $Tolerance.Mode) -ne "Filled") { return }
    $qaColumnIndex = $Row.DataGridView.Columns["QaValue"].Index
    $cell = New-Object System.Windows.Forms.DataGridViewCheckBoxCell
    $cell.ThreeState = $false
    $cell.TrueValue = "Yes"
    $cell.FalseValue = "No"
    $cell.IndeterminateValue = ""
    $cell.Value = $null
    $cell.Style.NullValue = $false
    $cell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $cell.ToolTipText = "Starts unchecked. Check this only after confirming the crew/database value is filled in."
    $Row.Cells[$qaColumnIndex] = $cell
}

function Add-HooterCheckRow {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [object]$Field,
        [string]$CrewValue,
        [int]$EntryNumber = 1,
        [string]$CrewRecord = ""
    )

    if (Test-HooterFieldExcluded -Field $Field) { return }
    $tolerance = Get-HooterTolerance -Field $Field
    $pointValue = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $tolerance)
    if ($Grid.Columns.Contains("EntryNumber")) {
        [void]$Grid.Rows.Add("|||", $EntryNumber, $CrewRecord, $Field.Label, $CrewValue, "", (Get-HooterRuleText $tolerance), $pointValue, "", "", "", $Field.FieldKey, $Field.Group, $Field.TableName, $Field.FieldName)
    }
    else {
        [void]$Grid.Rows.Add("|||", $Field.Label, $CrewValue, "", (Get-HooterRuleText $tolerance), $pointValue, "", "", "", $Field.FieldKey, $Field.Group, $Field.TableName, $Field.FieldName)
    }
    Set-HooterFilledCheckCell -Row $Grid.Rows[$Grid.Rows.Count - 1] -Tolerance $tolerance
}

function Update-HooterRowStatus {
    param(
        [System.Windows.Forms.DataGridViewRow]$Row
    )

    if ($null -eq $Row -or $Row.IsNewRow) { return }
    $tolerance = Get-HooterRowTolerance -Row $Row
    $Row.Cells["Rule"].Value = Get-HooterRuleText $tolerance
    $result = Test-HooterFieldPass -CrewValue $Row.Cells["CrewValue"].Value -QaValue $Row.Cells["QaValue"].Value -Tolerance $tolerance
    $Row.Cells["Status"].Value = $result.Status
    if ($Row.DataGridView.Columns.Contains("Score")) {
        $score = Get-HooterRowScoreResult -Status $result.Status -Tolerance $tolerance -CrewValue $Row.Cells["CrewValue"].Value -QaValue $Row.Cells["QaValue"].Value
        $Row.Cells["Score"].Value = $score.Display
    }
    switch ($result.Status) {
        "Pass" { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(236, 249, 241) }
        "Fail" { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 230) }
        default { $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White }
    }
}

function Refresh-HooterGridStatuses {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return }
    foreach ($row in $Grid.Rows) {
        Update-HooterRowStatus -Row $row
    }
}

function Get-HooterGridCheckStats {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Scope = ""
    )

    $total = 0
    $checked = 0
    $failed = 0
    $earnedPoints = 0.0
    $possiblePoints = 0.0
    $rowEarnedPoints = 0.0
    $rowPossiblePoints = 0.0
    $lostPoints = 0.0
    $criticalFailures = 0
    if ($null -ne $Grid) {
        foreach ($row in $Grid.Rows) {
            if ($row.IsNewRow) { continue }
            $total++
            $tolerance = Get-HooterRowTolerance -Row $row
            $status = ConvertTo-HooterText $row.Cells["Status"].Value
            $score = Get-HooterRowScoreResult -Status $status -Tolerance $tolerance -CrewValue $row.Cells["CrewValue"].Value -QaValue $row.Cells["QaValue"].Value
            $rowPossiblePoints += $score.PossiblePoints
            $rowEarnedPoints += $score.EarnedPoints
            $lostPoints += $score.LostPoints
            if ($score.CriticalFailure) { $criticalFailures++ }
            if ($status -eq "Fail") { $failed++ }
            if (-not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.Cells["QaValue"].Value))) { $checked++ }
        }
    }
    $possiblePoints = if ($total -gt 0) { Get-HooterSectionPointTotal -Scope $Scope -Fallback $rowPossiblePoints } else { 0.0 }
    $earnedPoints = if ($total -gt 0 -and $checked -eq $total) {
        [Math]::Max(0.0, $possiblePoints - $lostPoints)
    }
    else {
        [Math]::Min($possiblePoints, $rowEarnedPoints)
    }
    return [pscustomobject]@{
        Total = $total
        Checked = $checked
        Failed = $failed
        EarnedPoints = $earnedPoints
        PossiblePoints = $possiblePoints
        RowPossiblePoints = $rowPossiblePoints
        LostPoints = $lostPoints
        CriticalFailures = $criticalFailures
        MaxExceeded = ($possiblePoints -gt 0 -and $lostPoints -gt $possiblePoints)
    }
}

function Get-HooterExecutionGridStats {
    $total = 0
    $checked = 0
    $failed = 0
    $earnedPoints = 0.0
    $rowPossiblePoints = 0.0
    $lostPoints = 0.0
    $criticalFailures = 0

    if ($script:Ui.ContainsKey("ExecutionGrid")) {
        foreach ($row in $script:Ui.ExecutionGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $result = Get-HooterExecutionRowResult -Row $row
            $total++
            $rowPossiblePoints += $result.PossiblePoints
            $earnedPoints += $result.EarnedPoints
            $lostPoints += $result.PointLoss
            if ($result.Checked) { $checked++ }
            if ($result.Status -eq "Fail") { $failed++ }
            if ($result.CriticalFailure) { $criticalFailures++ }
        }
    }

    $possiblePoints = Get-HooterSectionPointTotal -Scope "Execution" -Fallback $rowPossiblePoints
    $earnedPoints = if ($total -gt 0 -and $checked -eq $total) {
        [Math]::Max(0.0, $possiblePoints - $lostPoints)
    }
    else {
        [Math]::Min($possiblePoints, $earnedPoints)
    }

    return [pscustomobject]@{
        Total = $total
        Checked = $checked
        Failed = $failed
        EarnedPoints = $earnedPoints
        PossiblePoints = $possiblePoints
        RowPossiblePoints = $rowPossiblePoints
        LostPoints = $lostPoints
        CriticalFailures = $criticalFailures
        MaxExceeded = ($possiblePoints -gt 0 -and $lostPoints -gt $possiblePoints)
    }
}

function Get-HooterMissedTreeRowCount {
    if (-not $script:Ui.ContainsKey("MissedTreeGrid")) { return 0 }
    $count = 0
    foreach ($row in $script:Ui.MissedTreeGrid.Rows) {
        if ($row.IsNewRow) { continue }
        $count++
    }
    return $count
}

function Test-HooterCrewMissedTreeFail {
    $checked = $false
    if ($script:Ui.ContainsKey("CrewMissedTreeCheck") -and $null -ne $script:Ui.CrewMissedTreeCheck) {
        $checked = [bool]$script:Ui.CrewMissedTreeCheck.Checked
    }
    return ($checked -or (Get-HooterMissedTreeRowCount) -gt 0)
}

function Get-HooterMissedTreeStats {
    $rowCount = Get-HooterMissedTreeRowCount
    $hasFail = Test-HooterCrewMissedTreeFail
    $total = if ($hasFail) { [Math]::Max(1, $rowCount) } else { 0 }

    return [pscustomobject]@{
        Total = $total
        Checked = $total
        Failed = $total
        EarnedPoints = 0.0
        PossiblePoints = 0.0
        RowPossiblePoints = 0.0
        LostPoints = 0.0
        CriticalFailures = $total
        MaxExceeded = $false
        MissedTreeCount = $rowCount
    }
}

function Get-HooterOverallStatusFromStats {
    param(
        [int]$Total,
        [int]$Checked,
        [int]$Failed,
        [double]$EarnedPoints,
        [double]$PossiblePoints,
        [double]$LostPoints,
        [int]$CriticalFailures,
        [int]$SectionFailures = 0
    )

    $scorePercent = if ($PossiblePoints -gt 0) { [Math]::Round(($EarnedPoints / $PossiblePoints) * 100.0, 2) } else { 0.0 }
    $passPercent = Get-HooterScorePassPercent
    $maxPointLoss = Get-HooterMaxPointLoss
    $maxPointLossConfigured = Test-HooterMaxPointLossConfigured
    $status = "Not checked"
    if ($CriticalFailures -gt 0) {
        $status = "Fail"
    }
    elseif ($SectionFailures -gt 0) {
        $status = "Fail"
    }
    elseif ($Total -gt 0 -and $Checked -eq $Total) {
        if ($maxPointLossConfigured -and $LostPoints -gt $maxPointLoss) { $status = "Fail" } else { $status = "Pass" }
    }
    elseif ($Checked -gt 0) {
        $status = "Incomplete"
    }

    return [pscustomobject]@{
        Status = $status
        Total = $Total
        Checked = $Checked
        Failed = $Failed
        EarnedPoints = $EarnedPoints
        PossiblePoints = $PossiblePoints
        LostPoints = $LostPoints
        ScorePercent = $scorePercent
        ScorePassPercent = $passPercent
        MaxPointLoss = if ($maxPointLossConfigured) { $maxPointLoss } else { "" }
        MaxPointLossConfigured = $maxPointLossConfigured
        CriticalFailures = $CriticalFailures
        SectionFailures = $SectionFailures
        UncheckedCount = [Math]::Max(0, $Total - $Checked)
    }
}

function Get-HooterOverallStatusFromGrids {
    $plotStats = Get-HooterGridCheckStats -Grid $script:Ui.PlotGrid -Scope "Plot"
    $treeStats = Get-HooterGridCheckStats -Grid $script:Ui.TreeGrid -Scope "Tree"
    $regenStats = Get-HooterGridCheckStats -Grid $script:Ui.RegenGrid -Scope "Regen"
    $executionStats = Get-HooterExecutionGridStats
    $missedTreeStats = Get-HooterMissedTreeStats

    $total = $plotStats.Total + $treeStats.Total + $regenStats.Total + $executionStats.Total + $missedTreeStats.Total
    $checked = $plotStats.Checked + $treeStats.Checked + $regenStats.Checked + $executionStats.Checked + $missedTreeStats.Checked
    $failed = $plotStats.Failed + $treeStats.Failed + $regenStats.Failed + $executionStats.Failed + $missedTreeStats.Failed
    $earnedPoints = $plotStats.EarnedPoints + $treeStats.EarnedPoints + $regenStats.EarnedPoints + $executionStats.EarnedPoints + $missedTreeStats.EarnedPoints
    $possiblePoints = $plotStats.PossiblePoints + $treeStats.PossiblePoints + $regenStats.PossiblePoints + $executionStats.PossiblePoints + $missedTreeStats.PossiblePoints
    $lostPoints = $plotStats.LostPoints + $treeStats.LostPoints + $regenStats.LostPoints + $executionStats.LostPoints + $missedTreeStats.LostPoints
    $criticalFailures = $plotStats.CriticalFailures + $treeStats.CriticalFailures + $regenStats.CriticalFailures + $executionStats.CriticalFailures + $missedTreeStats.CriticalFailures
    $sectionFailures = @(@($plotStats, $treeStats, $regenStats, $executionStats) | Where-Object { $_.MaxExceeded }).Count

    $overall = Get-HooterOverallStatusFromStats -Total $total -Checked $checked -Failed $failed -EarnedPoints $earnedPoints -PossiblePoints $possiblePoints -LostPoints $lostPoints -CriticalFailures $criticalFailures -SectionFailures $sectionFailures
    $overall | Add-Member -NotePropertyName PlotPointTotal -NotePropertyValue $plotStats.PossiblePoints -Force
    $overall | Add-Member -NotePropertyName TreePointTotal -NotePropertyValue $treeStats.PossiblePoints -Force
    $overall | Add-Member -NotePropertyName RegenPointTotal -NotePropertyValue $regenStats.PossiblePoints -Force
    $overall | Add-Member -NotePropertyName ExecutionPointTotal -NotePropertyValue $executionStats.PossiblePoints -Force
    $overall | Add-Member -NotePropertyName PlotPointLoss -NotePropertyValue $plotStats.LostPoints -Force
    $overall | Add-Member -NotePropertyName TreePointLoss -NotePropertyValue $treeStats.LostPoints -Force
    $overall | Add-Member -NotePropertyName RegenPointLoss -NotePropertyValue $regenStats.LostPoints -Force
    $overall | Add-Member -NotePropertyName ExecutionPointLoss -NotePropertyValue $executionStats.LostPoints -Force
    $overall | Add-Member -NotePropertyName MissedTreeCount -NotePropertyValue $missedTreeStats.MissedTreeCount -Force
    return $overall
}

function Update-HooterErrorSummary {
    if (-not $script:Ui.ContainsKey("ErrorSummaryGrid")) { return }

    $grid = $script:Ui.ErrorSummaryGrid
    $grid.Rows.Clear()

    $plotStats = Get-HooterGridCheckStats -Grid $script:Ui.PlotGrid -Scope "Plot"
    $treeStats = Get-HooterGridCheckStats -Grid $script:Ui.TreeGrid -Scope "Tree"
    $regenStats = Get-HooterGridCheckStats -Grid $script:Ui.RegenGrid -Scope "Regen"
    $executionStats = Get-HooterExecutionGridStats
    $missedTreeStats = Get-HooterMissedTreeStats
    $maxPointLossDisplay = Get-HooterMaxPointLossDisplay
    $totalLost = $plotStats.LostPoints + $treeStats.LostPoints + $regenStats.LostPoints + $executionStats.LostPoints
    $totalPossible = $plotStats.PossiblePoints + $treeStats.PossiblePoints + $regenStats.PossiblePoints + $executionStats.PossiblePoints

    $addSummaryRow = {
        param([string]$Label, [object]$Stats)
        $rowIndex = $grid.Rows.Add($Label, (Format-HooterScoreNumber $Stats.LostPoints), (Format-HooterScoreNumber $Stats.PossiblePoints))
        if ($rowIndex -ge 0 -and $Stats.MaxExceeded) {
            $grid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 230)
            $grid.Rows[$rowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(130, 36, 24)
            $grid.Rows[$rowIndex].DefaultCellStyle.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
        }
    }
    & $addSummaryRow "Total From Table A - Plot Classification" $plotStats
    & $addSummaryRow "Total from Table B - Tree Classification" $treeStats
    if ($missedTreeStats.Total -gt 0) {
        $rowIndex = $grid.Rows.Add("Crew missed tree(s) found by QA - automatic plot failure", "Critical", "0")
        if ($rowIndex -ge 0) {
            $grid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 230)
            $grid.Rows[$rowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(130, 36, 24)
            $grid.Rows[$rowIndex].DefaultCellStyle.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
        }
    }
    & $addSummaryRow "Total from Table C - Regen / Tree Count" $regenStats
    & $addSummaryRow "Total from Table D - Plot Location and Execution" $executionStats
    $totalIndex = $grid.Rows.Add("Total max point loss (section max sum)", (Format-HooterScoreNumber $totalLost), (Format-HooterScoreNumber $totalPossible))
    if ($totalIndex -ge 0) {
        $grid.Rows[$totalIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(170, 170, 170)
        $grid.Rows[$totalIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        $grid.Rows[$totalIndex].DefaultCellStyle.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
    }
    $thresholdIndex = $grid.Rows.Add("Failure threshold (total error greater than this fails)", "", $maxPointLossDisplay)
    if ($thresholdIndex -ge 0) {
        $grid.Rows[$thresholdIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 241, 225)
        $grid.Rows[$thresholdIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(92, 65, 18)
        $grid.Rows[$thresholdIndex].DefaultCellStyle.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
    }
}

function Update-HooterTreeMaxSummaryLabel {
    if (-not $script:Ui.ContainsKey("TreeMaxSummaryLabel")) { return }

    $summary = Get-HooterTreeMaxLossSummary
    $script:Ui.TreeMaxSummaryLabel.Text = "Tree Count: $($summary.TreeCount)    Tree total max loss: $($summary.TotalText)"
}

function Update-HooterOverallLabel {
    $overall = Get-HooterOverallStatusFromGrids
    $status = $overall.Status
    if ($script:Ui.ContainsKey("OverallLabel")) {
        $remaining = [Math]::Max(0, $overall.Total - $overall.Checked)
        $failedLabel = if ($overall.Failed -eq 1) { "1 failed check row" } else { "$($overall.Failed) failed check rows" }
        $remainingText = if ($remaining -gt 0) { "; $remaining unchecked rows" } else { "" }
        $criticalText = if ($overall.CriticalFailures -gt 0) { "; $($overall.CriticalFailures) critical failures" } else { "" }
        $sectionText = if ($overall.SectionFailures -gt 0) { "; $($overall.SectionFailures) section over max" } else { "" }
        $thresholdText = Get-HooterMaxPointLossSummaryText -Value $overall.MaxPointLoss
        $script:Ui.OverallLabel.Text = "Overall QA status: $status (total error $(Format-HooterScoreNumber $overall.LostPoints); $thresholdText; $failedLabel$remainingText$criticalText$sectionText)"
        switch ($status) {
            "Pass" { $script:Ui.OverallLabel.ForeColor = [System.Drawing.Color]::FromArgb(31, 112, 74) }
            "Fail" { $script:Ui.OverallLabel.ForeColor = [System.Drawing.Color]::FromArgb(170, 54, 35) }
            "Incomplete" { $script:Ui.OverallLabel.ForeColor = [System.Drawing.Color]::FromArgb(128, 92, 22) }
            default { $script:Ui.OverallLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 42) }
        }
    }
    Update-HooterErrorSummary
    Update-HooterTreeMaxSummaryLabel
    Update-HooterInventoryProgress
}

function Refresh-HooterAllStatuses {
    Refresh-HooterGridStatuses -Grid $script:Ui.PlotGrid
    Refresh-HooterGridStatuses -Grid $script:Ui.TreeGrid
    Refresh-HooterGridStatuses -Grid $script:Ui.RegenGrid
    Refresh-HooterExecutionGridStatuses
    Update-HooterOverallLabel
}

function Request-HooterTimerRefresh {
    param(
        [string]$TimerName,
        [scriptblock]$Action,
        [int]$Interval = 180
    )

    try {
        $timer = Get-Variable -Scope Script -Name $TimerName -ValueOnly -ErrorAction SilentlyContinue
        if ($null -eq $timer) {
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = $Interval
            $timer.Add_Tick({
                param($Sender, $EventArgs)
                try {
                    $Sender.Stop()
                    & $Action
                }
                catch {
                }
            }.GetNewClosure())
            Set-Variable -Scope Script -Name $TimerName -Value $timer
        }
        $timer.Stop()
        $timer.Start()
    }
    catch {
        & $Action
    }
}

function Test-HooterEntryGridIsEditing {
    foreach ($key in @("PlotGrid", "TreeGrid", "RegenGrid")) {
        if (-not $script:Ui.ContainsKey($key)) { continue }
        $grid = $script:Ui[$key]
        if ($null -eq $grid) { continue }
        try {
            if ($grid.IsCurrentCellInEditMode -or $null -ne $grid.EditingControl) { return $true }
        }
        catch {}
    }
    return $false
}

function Request-HooterOverallRefresh {
    Request-HooterTimerRefresh -TimerName "OverallRefreshTimer" -Interval 1600 -Action {
        if (Test-HooterEntryGridIsEditing) {
            Request-HooterOverallRefresh
            return
        }
        Update-HooterOverallLabel
    }
}

function Request-HooterTreeProgressRefresh {
    Request-HooterTimerRefresh -TimerName "TreeProgressRefreshTimer" -Interval 1500 -Action {
        if (Test-HooterEntryGridIsEditing) {
            Request-HooterTreeProgressRefresh
            return
        }
        Update-HooterTreeProgressLabel
    }
}

function Request-HooterRegenProgressRefresh {
    Request-HooterTimerRefresh -TimerName "RegenProgressRefreshTimer" -Interval 1500 -Action {
        if (Test-HooterEntryGridIsEditing) {
            Request-HooterRegenProgressRefresh
            return
        }
        Update-HooterRegenProgressLabel
    }
}

function Request-HooterDeferredSettingsSave {
    Request-HooterTimerRefresh -TimerName "DeferredSettingsSaveTimer" -Action {
        Invoke-HooterAutoSaveOperation -Text "Auto saving setup..." -Action { Save-HooterSettings }
    } -Interval 3000
}

function Request-HooterStemScoringRefresh {
    Request-HooterTimerRefresh -TimerName "StemScoringRefreshTimer" -Action {
        Invoke-HooterAutoSaveOperation -Text "Updating regen scoring..." -Action {
            Update-HooterStemScoringRowsInSettingsGrid
            Update-HooterSettingsFromGrid
            Refresh-HooterAllStatuses
            Save-HooterSettings
        }
        Set-HooterStatus $(if ($script:UseStemCountPercentageForScoring) { "Using regen stem count percentage scoring. The stem-count table is off." } else { "Using regen stem count table scoring. The table is active." })
    } -Interval 250
}

function Request-HooterMissedTreeRefresh {
    Request-HooterTimerRefresh -TimerName "MissedTreeRefreshTimer" -Action {
        Invoke-HooterAutoSaveOperation -Text "Updating missed tree check..." -CompletedText "Updated" -Action {
            Update-HooterMissedTreeStatus
            Update-HooterOverallLabel
        }
        if (Test-HooterCrewMissedTreeFail) {
            Set-HooterStatus "Crew missed tree check is marked. This forces the plot to fail."
        }
        else {
            Set-HooterStatus "Crew missed tree check is clear."
        }
    } -Interval 220
}

function Stop-HooterDeferredSettingsSave {
    try {
        if ($null -ne $script:DeferredSettingsSaveTimer) {
            $script:DeferredSettingsSaveTimer.Stop()
        }
    }
    catch {}
}

function Set-HooterStatus {
    param([string]$Text)

    if ($script:Ui.ContainsKey("StatusLabel")) {
        $script:Ui.StatusLabel.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Show-HooterAutoSaveNotice {
    param([string]$Text = "Auto saving...")

    try {
        if (-not $script:Ui.ContainsKey("AutoSaveNoticePanel")) { return }
        $panel = $script:Ui.AutoSaveNoticePanel
        if ($null -eq $panel) { return }
        if ($script:Ui.ContainsKey("AutoSaveNoticeLabel") -and $null -ne $script:Ui.AutoSaveNoticeLabel) {
            $script:Ui.AutoSaveNoticeLabel.Text = $Text
        }
        if ($null -ne $script:AutoSaveNoticeHideTimer) {
            $script:AutoSaveNoticeHideTimer.Stop()
        }
        $panel.Visible = $true
        $panel.BringToFront()
        $panel.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
    }
}

function Hide-HooterAutoSaveNotice {
    param([int]$DelayMs = 0)

    try {
        if ($DelayMs -gt 0) {
            if ($null -eq $script:AutoSaveNoticeHideTimer) {
                $script:AutoSaveNoticeHideTimer = New-Object System.Windows.Forms.Timer
                $script:AutoSaveNoticeHideTimer.Add_Tick({
                    try {
                        $script:AutoSaveNoticeHideTimer.Stop()
                        Hide-HooterAutoSaveNotice
                    }
                    catch {
                    }
                })
            }
            $script:AutoSaveNoticeHideTimer.Interval = $DelayMs
            $script:AutoSaveNoticeHideTimer.Stop()
            $script:AutoSaveNoticeHideTimer.Start()
            return
        }

        if ($script:Ui.ContainsKey("AutoSaveNoticePanel") -and $null -ne $script:Ui.AutoSaveNoticePanel) {
            $script:Ui.AutoSaveNoticePanel.Visible = $false
        }
    }
    catch {
    }
}

function Invoke-HooterAutoSaveOperation {
    param(
        [string]$Text = "Auto saving setup...",
        [string]$CompletedText = "Saved",
        [scriptblock]$Action
    )

    if ($null -eq $Action) { return }
    try {
        Show-HooterAutoSaveNotice -Text $Text
        & $Action
        if ($script:Ui.ContainsKey("AutoSaveNoticeLabel") -and $null -ne $script:Ui.AutoSaveNoticeLabel) {
            $script:Ui.AutoSaveNoticeLabel.Text = $CompletedText
        }
        Hide-HooterAutoSaveNotice -DelayMs 900
    }
    catch {
        Hide-HooterAutoSaveNotice
        throw
    }
}

function Show-HooterBusyDialog {
    param(
        [string]$Title = "Working",
        [string]$Message = "Please wait..."
    )

    if (-not $script:Ui.ContainsKey("Form")) { return $null }
    $owner = $script:Ui.Form

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.StartPosition = "Manual"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.ControlBox = $false
    $dialog.ShowInTaskbar = $false
    $dialog.Size = New-Object System.Drawing.Size(560, 190)
    $dialog.MinimumSize = New-Object System.Drawing.Size(560, 190)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    $dialog.Font = New-HooterFont 9.4
    $dialog.UseWaitCursor = $true

    $label = New-HooterLabel $Message 28 24 488 44 10.0 ([System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $false
    $dialog.Controls.Add($label)

    $hint = New-HooterLabel "PlotHoot is reading the database and building QA rows." 28 78 488 24 9.0
    $dialog.Controls.Add($hint)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(28, 116)
    $progress.Size = New-Object System.Drawing.Size(488, 22)
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progress.MarqueeAnimationSpeed = 35
    $dialog.Controls.Add($progress)

    try {
        $ownerBounds = $owner.RectangleToScreen($owner.ClientRectangle)
        if ($ownerBounds.Width -le 0 -or $ownerBounds.Height -le 0) {
            $ownerBounds = $owner.Bounds
        }
        $screenBounds = [System.Windows.Forms.Screen]::FromControl($owner).WorkingArea
        $x = [int][Math]::Round($ownerBounds.Left + (($ownerBounds.Width - $dialog.Width) / 2.0))
        $y = [int][Math]::Round($ownerBounds.Top + (($ownerBounds.Height - $dialog.Height) / 2.0))
        $maxX = [Math]::Max($screenBounds.Left, $screenBounds.Right - $dialog.Width)
        $maxY = [Math]::Max($screenBounds.Top, $screenBounds.Bottom - $dialog.Height)
        $x = [Math]::Min([Math]::Max($screenBounds.Left, $x), $maxX)
        $y = [Math]::Min([Math]::Max($screenBounds.Top, $y), $maxY)
        $dialog.Location = New-Object System.Drawing.Point($x, $y)
    }
    catch {
        $dialog.StartPosition = "CenterScreen"
    }

    $owner.UseWaitCursor = $true
    $owner.Enabled = $false
    try {
        $dialog.Show($owner)
    }
    catch {
        $dialog.Show()
    }
    $dialog.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    return $dialog
}

function Close-HooterBusyDialog {
    param([System.Windows.Forms.Form]$Dialog)

    if ($script:Ui.ContainsKey("Form")) {
        $script:Ui.Form.Enabled = $true
        $script:Ui.Form.UseWaitCursor = $false
    }
    if ($null -ne $Dialog) {
        try { $Dialog.Close() } catch {}
        try { $Dialog.Dispose() } catch {}
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-HooterStemToleranceBandSetting {
    param([string]$Key)

    $keyText = ConvertTo-HooterText $Key
    if ($keyText.Contains("|")) {
        $keyText = ConvertTo-HooterText (@($keyText -split "\|")[-1])
    }
    foreach ($band in @(Get-HooterStemToleranceBands)) {
        if ($keyText.Equals((ConvertTo-HooterText $band.Key), [System.StringComparison]::OrdinalIgnoreCase) -or
            $keyText.Equals((ConvertTo-HooterText $band.Label), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $band
        }
    }
    return $null
}

function Format-HooterStemToleranceValue {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    return Format-HooterScoreNumber $text
}

function Populate-HooterStemToleranceGrid {
    param([System.Windows.Forms.DataGridView]$Grid = $null)

    if ($null -eq $Grid) {
        if (-not $script:Ui.ContainsKey("StemToleranceGrid")) { return }
        $Grid = $script:Ui.StemToleranceGrid
    }
    $Grid.Rows.Clear()
    foreach ($band in @(Get-HooterStemToleranceBands)) {
        [void]$Grid.Rows.Add((ConvertTo-HooterText $band.Label), (Format-HooterStemToleranceValue $band.Tolerance), (ConvertTo-HooterText $band.Key))
    }
}

function Update-HooterStemToleranceBandsFromGrid {
    if (-not $script:Ui.ContainsKey("StemToleranceGrid")) { return }
    $grid = $script:Ui.StemToleranceGrid
    try { [void]$grid.EndEdit() } catch {}
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $key = ConvertTo-HooterText $row.Cells["StemBandKey"].Value
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        [void](Set-HooterStemToleranceBand -Key $key -Tolerance $row.Cells["ToleranceValue"].Value)
        $band = Get-HooterStemToleranceBandSetting -Key $key
        if ($null -ne $band) {
            $row.Cells["ToleranceValue"].Value = Format-HooterStemToleranceValue $band.Tolerance
        }
    }
}

function Update-HooterStemScoringControls {
    $usePercent = [bool]$script:UseStemCountPercentageForScoring
    if ($script:Ui.ContainsKey("UseStemPercentCheck") -and $null -ne $script:Ui.UseStemPercentCheck) {
        try {
            $script:SuppressStemScoringToggleEvent = $true
            $script:Ui.UseStemPercentCheck.Checked = $usePercent
        }
        finally {
            $script:SuppressStemScoringToggleEvent = $false
        }
    }
    if ($script:Ui.ContainsKey("StemToleranceLabel") -and $null -ne $script:Ui.StemToleranceLabel) {
        $script:Ui.StemToleranceLabel.Text = if ($usePercent) { "Regen StemCount table (off)" } else { "Regen StemCount table" }
        $script:Ui.StemToleranceLabel.ForeColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(105, 112, 108) } else { [System.Drawing.Color]::FromArgb(25, 75, 71) }
    }
    if ($script:Ui.ContainsKey("StemToleranceGrid") -and $null -ne $script:Ui.StemToleranceGrid) {
        $grid = $script:Ui.StemToleranceGrid
        $grid.Enabled = (-not $usePercent)
        $grid.ReadOnly = $usePercent
        $grid.BackgroundColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(226, 230, 228) } else { [System.Drawing.Color]::White }
        $grid.DefaultCellStyle.BackColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(226, 230, 228) } else { [System.Drawing.Color]::White }
        $grid.DefaultCellStyle.ForeColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(95, 101, 98) } else { [System.Drawing.Color]::FromArgb(25, 45, 42) }
        $grid.AlternatingRowsDefaultCellStyle.BackColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(226, 230, 228) } else { [System.Drawing.Color]::FromArgb(248, 250, 249) }
        $grid.ColumnHeadersDefaultCellStyle.BackColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(112, 122, 118) } else { [System.Drawing.Color]::FromArgb(30, 71, 68) }
        $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) { continue }
            $row.DefaultCellStyle.BackColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(226, 230, 228) } else { [System.Drawing.Color]::White }
            $row.DefaultCellStyle.ForeColor = if ($usePercent) { [System.Drawing.Color]::FromArgb(95, 101, 98) } else { [System.Drawing.Color]::FromArgb(25, 45, 42) }
        }
        try { $grid.Invalidate() } catch {}
    }
}

function Update-HooterStemScoringRowsInSettingsGrid {
    if (-not $script:Ui.ContainsKey("SettingsGrid")) { return }
    $grid = $script:Ui.SettingsGrid
    if (-not ($grid.Columns.Contains("FieldKey") -and $grid.Columns.Contains("Mode") -and $grid.Columns.Contains("ToleranceValue"))) { return }
    $previousSuppressGridEvents = [bool]$script:SuppressGridEvents
    try {
        $script:SuppressGridEvents = $true
        try { $grid.SuspendLayout() } catch {}
        foreach ($row in $grid.Rows) {
            if ($row.IsNewRow) { continue }
            $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
            $group = if ($grid.Columns.Contains("Group")) { ConvertTo-HooterText $row.Cells["Group"].Value } else { "" }
            $fieldName = if ($grid.Columns.Contains("FieldName")) { ConvertTo-HooterText $row.Cells["FieldName"].Value } else { "" }
            $fieldLabel = if ($grid.Columns.Contains("FieldLabel")) { ConvertTo-HooterText $row.Cells["FieldLabel"].Value } else { "" }
            if (-not ((Test-HooterRegenStemCountFieldKey -FieldKey $fieldKey) -or (Test-HooterRegenStemCountField -Group $group -FieldName $fieldName -Label $fieldLabel))) { continue }
            if ($script:UseStemCountPercentageForScoring) {
                $row.Cells["Mode"].Value = "StemPercent"
                if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.Cells["ToleranceValue"].Value))) {
                    $row.Cells["ToleranceValue"].Value = "Allowance=10; Step=5"
                }
            }
            else {
                $row.Cells["Mode"].Value = "StemCount"
                $row.Cells["ToleranceValue"].Value = ""
            }
        }
    }
    finally {
        try { $grid.ResumeLayout() } catch {}
        try { $grid.Invalidate() } catch {}
        $script:SuppressGridEvents = $previousSuppressGridEvents
    }
}

function Set-HooterUseStemCountPercentageForScoring {
    param(
        [bool]$Enabled,
        [switch]$Refresh
    )

    $script:UseStemCountPercentageForScoring = [bool]$Enabled
    Update-HooterStemScoringControls
    if ($Refresh) {
        Show-HooterAutoSaveNotice -Text "Updating regen scoring..."
        Request-HooterStemScoringRefresh
    }
    else {
        Update-HooterStemScoringRowsInSettingsGrid
    }
}

function Update-HooterSettingsFromGrid {
    param([switch]$UpdateFieldOrder)

    if (-not $script:Ui.ContainsKey("SettingsGrid")) { return }
    Update-HooterStemToleranceBandsFromGrid
    if ($script:Ui.ContainsKey("UseStemPercentCheck")) {
        $script:UseStemCountPercentageForScoring = [bool]$script:Ui.UseStemPercentCheck.Checked
    }
    if ($script:Ui.ContainsKey("MaxPointLossBox")) {
        $script:MaxPointLoss = Normalize-HooterMaxPointLoss $script:Ui.MaxPointLossBox.Text
        $script:Ui.MaxPointLossBox.Text = $script:MaxPointLoss
    }
    if ($script:Ui.ContainsKey("PlotPointTotalBox")) {
        $script:PlotPointTotal = Normalize-HooterPointTotal $script:Ui.PlotPointTotalBox.Text
        $script:Ui.PlotPointTotalBox.Text = $script:PlotPointTotal
    }
    if ($script:Ui.ContainsKey("TreePointTotalBox")) {
        $script:TreePointTotal = Normalize-HooterPointTotal $script:Ui.TreePointTotalBox.Text
        $script:Ui.TreePointTotalBox.Text = $script:TreePointTotal
    }
    if ($script:Ui.ContainsKey("RegenPointTotalBox")) {
        $script:RegenPointTotal = Normalize-HooterPointTotal $script:Ui.RegenPointTotalBox.Text
        $script:Ui.RegenPointTotalBox.Text = $script:RegenPointTotal
    }
    if ($script:Ui.ContainsKey("ExecutionPointTotalBox")) {
        $script:ExecutionPointTotal = Normalize-HooterPointTotal $script:Ui.ExecutionPointTotalBox.Text
        $script:Ui.ExecutionPointTotalBox.Text = $script:ExecutionPointTotal
    }
    $grid = $script:Ui.SettingsGrid
    if ($UpdateFieldOrder) {
        $script:FieldOrder = @{}
    }
    $fallbackOrder = 0
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        if ([string]::IsNullOrWhiteSpace($fieldKey)) { continue }
        $group = ConvertTo-HooterText $row.Cells["Group"].Value
        $fieldName = ConvertTo-HooterText $row.Cells["FieldName"].Value
        $fieldLabel = ConvertTo-HooterText $row.Cells["FieldLabel"].Value
        $useField = $true
        if ($grid.Columns.Contains("UseField")) {
            $useField = ConvertTo-HooterBool $row.Cells["UseField"].Value
        }
        if (-not (Test-HooterFieldUseCanBeChanged -Field $fieldKey)) {
            $useField = $true
            if ($grid.Columns.Contains("UseField")) { $row.Cells["UseField"].Value = $true }
        }
        if ($useField) {
            if ($script:ExcludedFieldKeys.ContainsKey($fieldKey)) { $script:ExcludedFieldKeys.Remove($fieldKey) }
        }
        else {
            $script:ExcludedFieldKeys[$fieldKey] = $true
        }
        $fieldOrder = $fallbackOrder
        if ($grid.Columns.Contains("FieldOrder")) {
            [void][int]::TryParse((ConvertTo-HooterText $row.Cells["FieldOrder"].Value), [ref]$fieldOrder)
        }
        $mode = Normalize-HooterToleranceMode $row.Cells["Mode"].Value
        $row.Cells["Mode"].Value = $mode
        if ((Test-HooterFilledOnlyFieldKey -FieldKey $fieldKey) -or (Test-HooterFilledOnlyField -Group $group -FieldName $fieldName -Label $fieldLabel)) {
            $mode = "Filled"
            $row.Cells["Mode"].Value = "Filled"
            $row.Cells["ToleranceValue"].Value = ""
        }
        if ($UpdateFieldOrder) {
            $script:FieldOrder[$fieldKey] = $fieldOrder
        }
        $fallbackOrder++
        if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) {
            $mode = "GoodFairPoor"
            $row.Cells["Mode"].Value = "GoodFairPoor"
        }
        if ($mode -eq "StemPercent" -and [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.Cells["ToleranceValue"].Value))) {
            $row.Cells["ToleranceValue"].Value = "Allowance=10; Step=5"
        }
        if ($mode -eq "Exact") {
            $row.Cells["ToleranceValue"].Value = ""
        }
        if ((Test-HooterRegenStemCountFieldKey -FieldKey $fieldKey) -or (Test-HooterRegenStemCountField -Group $group -FieldName $fieldName -Label $fieldLabel)) {
            if ($script:UseStemCountPercentageForScoring) {
                $mode = "StemPercent"
                $row.Cells["Mode"].Value = "StemPercent"
                if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.Cells["ToleranceValue"].Value))) {
                    $row.Cells["ToleranceValue"].Value = "Allowance=10; Step=5"
                }
            }
            else {
                $mode = "StemCount"
                $row.Cells["Mode"].Value = "StemCount"
                $row.Cells["ToleranceValue"].Value = ""
            }
        }
        $toleranceValue = ConvertTo-HooterText $row.Cells["ToleranceValue"].Value
        if ($mode -eq "Exact") { $toleranceValue = "" }
        $tolerance = [pscustomobject]@{
            FieldKey = $fieldKey
            Group = $group
            TableName = ConvertTo-HooterText $row.Cells["TableName"].Value
            FieldName = $fieldName
            Label = $fieldLabel
            Mode = $mode
            Value = $toleranceValue
            PointValue = ConvertTo-HooterText $row.Cells["PointValue"].Value
            CriticalFail = ConvertTo-HooterBool $row.Cells["CriticalFail"].Value
        }
        $script:Tolerances[$fieldKey] = $tolerance
        if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) {
            $rule = Get-HooterExecutionScoreRule -ItemKey $tolerance.FieldName
            $tolerance.Value = ConvertTo-HooterText $rule.ToleranceValue
            $tolerance.PointValue = Format-HooterScoreNumber $rule.PossiblePoints
            $row.Cells["ToleranceValue"].Value = $tolerance.Value
            $row.Cells["PointValue"].Value = $tolerance.PointValue
        }
        Update-HooterSettingsGridRowStyle -Row $row
    }
    if ($UpdateFieldOrder) {
        Apply-HooterSavedFieldOrderToCatalog
    }
}

function Set-HooterSettingsGridOrderFromRows {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid -or -not $Grid.Columns.Contains("FieldOrder")) { return }
    $order = 0
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $row.Cells["FieldOrder"].Value = $order
        $order++
    }
}

function Update-HooterSettingsGridRowStyle {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $Row.IsNewRow -or $null -eq $Row.DataGridView) { return }
    $grid = $Row.DataGridView
    $mode = if ($grid.Columns.Contains("Mode")) { Normalize-HooterToleranceMode $Row.Cells["Mode"].Value } else { "" }
    $fieldKey = if ($grid.Columns.Contains("FieldKey")) { ConvertTo-HooterText $Row.Cells["FieldKey"].Value } else { "" }
    $enabled = $true
    if ($grid.Columns.Contains("UseField")) {
        $enabled = ConvertTo-HooterBool $Row.Cells["UseField"].Value
    }

    if ($grid.Columns.Contains("ToleranceValue")) {
        $valueCell = $Row.Cells["ToleranceValue"]
        $valueCell.ReadOnly = ($mode -eq "Exact")
        if ($mode -eq "Exact") { $valueCell.Value = "" }
        $valueCell.Style.BackColor = if ($mode -eq "Exact" -or -not $enabled) { [System.Drawing.Color]::FromArgb(235, 238, 236) } else { [System.Drawing.Color]::White }
        $valueCell.Style.ForeColor = if ($mode -eq "Exact" -or -not $enabled) { [System.Drawing.Color]::FromArgb(105, 112, 108) } else { [System.Drawing.Color]::FromArgb(25, 45, 42) }
    }
    if ($grid.Columns.Contains("UnitHint")) {
        $Row.Cells["UnitHint"].Value = Get-HooterToleranceUnitHint -Tolerance ([pscustomobject]@{
            Mode = $mode
            FieldName = if ($grid.Columns.Contains("FieldName")) { ConvertTo-HooterText $Row.Cells["FieldName"].Value } else { "" }
            Label = if ($grid.Columns.Contains("FieldLabel")) { ConvertTo-HooterText $Row.Cells["FieldLabel"].Value } else { "" }
        })
    }
    if ($grid.Columns.Contains("UseField") -and -not (Test-HooterFieldUseCanBeChanged -Field $fieldKey)) {
        $Row.Cells["UseField"].ReadOnly = $true
        $Row.Cells["UseField"].Value = $true
    }
    $Row.DefaultCellStyle.ForeColor = if ($enabled) { [System.Drawing.Color]::FromArgb(25, 45, 42) } else { [System.Drawing.Color]::FromArgb(115, 122, 118) }
    $Row.DefaultCellStyle.BackColor = if ($enabled) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::FromArgb(242, 244, 243) }
}

function Refresh-HooterSettingsGridStyles {
    if (-not $script:Ui.ContainsKey("SettingsGrid")) { return }
    foreach ($row in $script:Ui.SettingsGrid.Rows) {
        Update-HooterSettingsGridRowStyle -Row $row
    }
}

function Apply-HooterSettingsFilter {
    if (-not ($script:Ui.ContainsKey("SettingsGrid") -and $script:Ui.ContainsKey("SettingsFilterBox"))) { return }
    $grid = $script:Ui.SettingsGrid
    $filter = ConvertTo-HooterText $script:Ui.SettingsFilterBox.Text
    try { $grid.CurrentCell = $null } catch {}
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $group = if ($grid.Columns.Contains("Group")) { ConvertTo-HooterText $row.Cells["Group"].Value } else { "" }
        $visible = $true
        if ($filter.Equals("Plot", [System.StringComparison]::OrdinalIgnoreCase)) { $visible = $group.Equals("Plot", [System.StringComparison]::OrdinalIgnoreCase) }
        elseif ($filter.Equals("Tree", [System.StringComparison]::OrdinalIgnoreCase)) { $visible = $group.Equals("Tree", [System.StringComparison]::OrdinalIgnoreCase) }
        elseif ($filter.Equals("Regen", [System.StringComparison]::OrdinalIgnoreCase)) { $visible = $group.Equals("Regen", [System.StringComparison]::OrdinalIgnoreCase) }
        elseif ($filter.Equals("Location / Execution", [System.StringComparison]::OrdinalIgnoreCase)) { $visible = $group.Equals("Execution", [System.StringComparison]::OrdinalIgnoreCase) }
        try { $row.Visible = $visible } catch {}
    }
}

function Populate-HooterSettingsGrid {
    if (-not $script:Ui.ContainsKey("SettingsGrid")) { return }
    if ($script:Ui.ContainsKey("StemToleranceGrid")) {
        Populate-HooterStemToleranceGrid
    }
    Update-HooterStemScoringControls
    if ($script:Ui.ContainsKey("MaxPointLossBox")) {
        $script:Ui.MaxPointLossBox.Text = Normalize-HooterMaxPointLoss $script:MaxPointLoss
    }
    if ($script:Ui.ContainsKey("PlotPointTotalBox")) {
        $script:Ui.PlotPointTotalBox.Text = Normalize-HooterPointTotal $script:PlotPointTotal
    }
    if ($script:Ui.ContainsKey("TreePointTotalBox")) {
        $script:Ui.TreePointTotalBox.Text = Normalize-HooterPointTotal $script:TreePointTotal
    }
    if ($script:Ui.ContainsKey("RegenPointTotalBox")) {
        $script:Ui.RegenPointTotalBox.Text = Normalize-HooterPointTotal $script:RegenPointTotal
    }
    if ($script:Ui.ContainsKey("ExecutionPointTotalBox")) {
        $script:Ui.ExecutionPointTotalBox.Text = Normalize-HooterPointTotal $script:ExecutionPointTotal
    }
    $grid = $script:Ui.SettingsGrid
    $grid.Rows.Clear()
    $allFields = @(Get-HooterAllScoredFields)
    $displayOrder = 0
    foreach ($field in $allFields) {
        $tolerance = Get-HooterTolerance -Field $field
        $fieldKey = ConvertTo-HooterText $field.FieldKey
        $fieldOrder = if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and $script:FieldOrder.ContainsKey($fieldKey)) { [int]$script:FieldOrder[$fieldKey] } else { $displayOrder }
        $mode = Normalize-HooterToleranceMode $tolerance.Mode
        $tolerance.Mode = $mode
        $toleranceValue = ConvertTo-HooterText $tolerance.Value
        if ($mode -eq "Exact") { $toleranceValue = "" }
        $unitHint = Get-HooterToleranceUnitHint -Field $field -Tolerance $tolerance
        $pointDisplay = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $tolerance)
        if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) {
            $executionRule = Get-HooterExecutionScoreRule -ItemKey $field.FieldName
            $toleranceValue = ConvertTo-HooterText $executionRule.ToleranceValue
            $pointDisplay = Format-HooterScoreNumber $executionRule.PossiblePoints
            $unitHint = "point loss"
        }
        $useField = (-not (Test-HooterFieldExcluded -Field $field)) -or (-not (Test-HooterFieldUseCanBeChanged -Field $field))
        [void]$grid.Rows.Add($useField, $field.Group, $field.Label, $mode, $toleranceValue, $unitHint, $pointDisplay, (Get-HooterCriticalFail -Tolerance $tolerance), $field.FieldKey, $field.TableName, $field.FieldName, $fieldOrder)
        $displayOrder++
    }
    Refresh-HooterSettingsGridStyles
    Apply-HooterSettingsFilter
}

function Get-HooterGridEditState {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Scope
    )

    $state = @{}
    if ($null -eq $Grid) { return $state }
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $entryNumber = if ($Grid.Columns.Contains("EntryNumber")) { ConvertTo-HooterText $row.Cells["EntryNumber"].Value } else { "1" }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        if ([string]::IsNullOrWhiteSpace($fieldKey)) { continue }
        $key = "$Scope|$entryNumber|$fieldKey"
        $state[$key] = [pscustomobject]@{
            QaValue = ConvertTo-HooterText $row.Cells["QaValue"].Value
            Notes = ConvertTo-HooterText $row.Cells["Notes"].Value
        }
    }
    return $state
}

function Restore-HooterGridEditState {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Scope,
        [hashtable]$State
    )

    if ($null -eq $Grid -or $null -eq $State) { return }
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $entryNumber = if ($Grid.Columns.Contains("EntryNumber")) { ConvertTo-HooterText $row.Cells["EntryNumber"].Value } else { "1" }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        $key = "$Scope|$entryNumber|$fieldKey"
        if (-not $State.ContainsKey($key)) { continue }
        $row.Cells["QaValue"].Value = $State[$key].QaValue
        $row.Cells["Notes"].Value = $State[$key].Notes
    }
}

function Apply-HooterFieldOrderToLoadedPlot {
    if ($null -eq $script:CurrentPlot) { return }
    if (-not ($script:Ui.ContainsKey("PlotGrid") -and $script:Ui.ContainsKey("TreeGrid") -and $script:Ui.ContainsKey("RegenGrid"))) { return }

    $plotState = Get-HooterGridEditState -Grid $script:Ui.PlotGrid -Scope "Plot"
    $treeState = Get-HooterGridEditState -Grid $script:Ui.TreeGrid -Scope "Tree"
    $regenState = Get-HooterGridEditState -Grid $script:Ui.RegenGrid -Scope "Regen"
    $selectedTree = if ($script:Ui.ContainsKey("TreeRecordBox")) { $script:Ui.TreeRecordBox.SelectedIndex } else { -1 }
    $selectedRegen = if ($script:Ui.ContainsKey("RegenRecordBox")) { $script:Ui.RegenRecordBox.SelectedIndex } else { -1 }

    $script:Ui.PlotGrid.Rows.Clear()
    foreach ($field in @($script:FieldCatalog.Plot)) {
        Add-HooterCheckRow -Grid $script:Ui.PlotGrid -Field $field -CrewValue (Get-RecordFieldValue -Record $script:CurrentPlot -Field $field)
    }
    Add-HooterAllTreeEntries
    Add-HooterAllRegenEntries

    Restore-HooterGridEditState -Grid $script:Ui.PlotGrid -Scope "Plot" -State $plotState
    Restore-HooterGridEditState -Grid $script:Ui.TreeGrid -Scope "Tree" -State $treeState
    Restore-HooterGridEditState -Grid $script:Ui.RegenGrid -Scope "Regen" -State $regenState

    if ($script:Ui.ContainsKey("TreeRecordBox") -and $selectedTree -ge 0 -and $selectedTree -lt $script:Ui.TreeRecordBox.Items.Count) {
        $script:Ui.TreeRecordBox.SelectedIndex = $selectedTree
    }
    if ($script:Ui.ContainsKey("RegenRecordBox") -and $selectedRegen -ge 0 -and $selectedRegen -lt $script:Ui.RegenRecordBox.Items.Count) {
        $script:Ui.RegenRecordBox.SelectedIndex = $selectedRegen
    }
    Refresh-HooterAllStatuses
    Set-HooterTreeVisibleRows
    Set-HooterRegenVisibleRows
}

function Copy-HooterGridRowForReorder {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row) { return $null }
    $copy = $Row.Clone()
    $copy.Tag = $Row.Tag
    try { $copy.DefaultCellStyle = $Row.DefaultCellStyle.Clone() } catch {}
    for ($index = 0; $index -lt $Row.Cells.Count; $index++) {
        $copy.Cells[$index].Value = $Row.Cells[$index].Value
        $copy.Cells[$index].ToolTipText = $Row.Cells[$index].ToolTipText
        $copy.Cells[$index].ReadOnly = $Row.Cells[$index].ReadOnly
        try { $copy.Cells[$index].Style = $Row.Cells[$index].Style.Clone() } catch {}
    }
    return $copy
}

function Apply-HooterDataEntryFieldOrderInPlace {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Scope,
        [string]$SourceKey,
        [string]$TargetKey,
        [string]$EntryNumber
    )

    if ($null -eq $Grid -or -not $Grid.Columns.Contains("FieldKey")) { return $false }
    $sourceKeyText = ConvertTo-HooterText $SourceKey
    $targetKeyText = ConvertTo-HooterText $TargetKey
    if ([string]::IsNullOrWhiteSpace($sourceKeyText) -or [string]::IsNullOrWhiteSpace($targetKeyText)) { return $false }

    $movesByEntry = @{}
    $moves = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $Grid.Rows.Count; $index++) {
        $row = $Grid.Rows[$index]
        if ($row.IsNewRow) { continue }
        $rowScope = if ($Grid.Columns.Contains("Group")) { ConvertTo-HooterText $row.Cells["Group"].Value } else { $Scope }
        if (-not $rowScope.Equals($Scope, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        if (-not ($fieldKey.Equals($sourceKeyText, [System.StringComparison]::OrdinalIgnoreCase) -or $fieldKey.Equals($targetKeyText, [System.StringComparison]::OrdinalIgnoreCase))) { continue }
        $entry = Get-HooterDataEntryRowEntryNumber -Row $row
        if (-not $movesByEntry.ContainsKey($entry)) {
            $move = [pscustomobject]@{
                EntryNumber = $entry
                SourceIndex = -1
                TargetIndex = -1
            }
            $movesByEntry[$entry] = $move
            [void]$moves.Add($move)
        }
        $entryMove = $movesByEntry[$entry]
        if ($fieldKey.Equals($sourceKeyText, [System.StringComparison]::OrdinalIgnoreCase)) {
            $entryMove.SourceIndex = $index
        }
        elseif ($fieldKey.Equals($targetKeyText, [System.StringComparison]::OrdinalIgnoreCase)) {
            $entryMove.TargetIndex = $index
        }
    }
    if ($moves.Count -eq 0) { return $false }

    $sourceCell = $null
    $moved = $false
    $script:SuppressGridEvents = $true
    try {
        try { [void]$Grid.EndEdit() } catch {}
        try { $Grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}
        try { $Grid.CurrentCell = $null } catch {}
        $Grid.SuspendLayout()

        foreach ($entryMove in @($moves.ToArray())) {
            $sourceIndex = [int]$entryMove.SourceIndex
            $targetIndex = [int]$entryMove.TargetIndex
            if ($sourceIndex -lt 0 -or $targetIndex -lt 0 -or $sourceIndex -eq $targetIndex) { continue }

            $rowCopy = Copy-HooterGridRowForReorder -Row $Grid.Rows[$sourceIndex]
            $visible = [bool]$Grid.Rows[$sourceIndex].Visible
            $Grid.Rows.RemoveAt($sourceIndex)
            if ($sourceIndex -lt $targetIndex) { $targetIndex-- }
            $targetIndex = [Math]::Max(0, [Math]::Min($targetIndex, $Grid.Rows.Count))
            $Grid.Rows.Insert($targetIndex, [System.Windows.Forms.DataGridViewRow]$rowCopy)
            try { $Grid.Rows[$targetIndex].Visible = $visible } catch {}
            if ($null -eq $sourceCell -and
                (ConvertTo-HooterText $entryMove.EntryNumber).Equals((ConvertTo-HooterText $EntryNumber), [System.StringComparison]::OrdinalIgnoreCase)) {
                $sourceCell = $Grid.Rows[$targetIndex].Cells["FieldLabel"]
            }
            $moved = $true
        }
    }
    finally {
        try { $Grid.ResumeLayout() } catch {}
        $script:SuppressGridEvents = $false
    }

    if (-not $moved) { return $false }

    if ($null -ne $sourceCell) {
        try {
            $Grid.ClearSelection()
            $Grid.CurrentCell = $sourceCell
            $sourceCell.OwningRow.Selected = $true
            $Grid.FirstDisplayedScrollingRowIndex = $sourceCell.RowIndex
        }
        catch {}
    }
    $Grid.Invalidate()
    if ($Scope -eq "Tree") { Request-HooterTreeProgressRefresh }
    elseif ($Scope -eq "Regen") { Request-HooterRegenProgressRefresh }
    return $true
}

function Get-HooterDataEntryGridScope {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return "" }
    if ($script:Ui.ContainsKey("PlotGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.PlotGrid)) { return "Plot" }
    if ($script:Ui.ContainsKey("TreeGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.TreeGrid)) { return "Tree" }
    if ($script:Ui.ContainsKey("RegenGrid") -and [object]::ReferenceEquals($Grid, $script:Ui.RegenGrid)) { return "Regen" }
    if ($null -ne $Grid.CurrentRow -and $Grid.Columns.Contains("Group")) {
        return (ConvertTo-HooterText $Grid.CurrentRow.Cells["Group"].Value)
    }
    return ""
}

function Get-HooterDataEntryRowEntryNumber {
    param([System.Windows.Forms.DataGridViewRow]$Row)

    if ($null -eq $Row -or $null -eq $Row.DataGridView) { return "" }
    if ($Row.DataGridView.Columns.Contains("EntryNumber")) {
        return ConvertTo-HooterText $Row.Cells["EntryNumber"].Value
    }
    return "1"
}

function Test-HooterDataEntryMoveTarget {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$SourceIndex,
        [int]$TargetIndex
    )

    if ($null -eq $Grid) { return $false }
    if (-not ($Grid.Columns.Contains("FieldKey") -and $Grid.Columns.Contains("FieldLabel"))) { return $false }
    if ($SourceIndex -lt 0 -or $TargetIndex -lt 0 -or $SourceIndex -ge $Grid.Rows.Count -or $TargetIndex -ge $Grid.Rows.Count) { return $false }
    $sourceRow = $Grid.Rows[$SourceIndex]
    $targetRow = $Grid.Rows[$TargetIndex]
    if ($sourceRow.IsNewRow -or $targetRow.IsNewRow) { return $false }
    $sourceKey = ConvertTo-HooterText $sourceRow.Cells["FieldKey"].Value
    $targetKey = ConvertTo-HooterText $targetRow.Cells["FieldKey"].Value
    if ([string]::IsNullOrWhiteSpace($sourceKey) -or [string]::IsNullOrWhiteSpace($targetKey)) { return $false }
    if ($sourceKey.Equals($targetKey, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

    $sourceGroup = if ($Grid.Columns.Contains("Group")) { ConvertTo-HooterText $sourceRow.Cells["Group"].Value } else { Get-HooterDataEntryGridScope -Grid $Grid }
    $targetGroup = if ($Grid.Columns.Contains("Group")) { ConvertTo-HooterText $targetRow.Cells["Group"].Value } else { Get-HooterDataEntryGridScope -Grid $Grid }
    if (-not $sourceGroup.Equals($targetGroup, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

    $sourceEntry = Get-HooterDataEntryRowEntryNumber -Row $sourceRow
    $targetEntry = Get-HooterDataEntryRowEntryNumber -Row $targetRow
    return $sourceEntry.Equals($targetEntry, [System.StringComparison]::OrdinalIgnoreCase)
}

function Set-HooterFieldOrderFromDataEntryMove {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$SourceIndex,
        [int]$TargetIndex
    )

    if (-not (Test-HooterDataEntryMoveTarget -Grid $Grid -SourceIndex $SourceIndex -TargetIndex $TargetIndex)) { return $false }
    $sourceRow = $Grid.Rows[$SourceIndex]
    $targetRow = $Grid.Rows[$TargetIndex]
    $scope = if ($Grid.Columns.Contains("Group")) { ConvertTo-HooterText $sourceRow.Cells["Group"].Value } else { Get-HooterDataEntryGridScope -Grid $Grid }
    $entryNumber = Get-HooterDataEntryRowEntryNumber -Row $sourceRow
    $sourceKey = ConvertTo-HooterText $sourceRow.Cells["FieldKey"].Value
    $targetKey = ConvertTo-HooterText $targetRow.Cells["FieldKey"].Value

    $keys = New-Object System.Collections.ArrayList
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        if ((Get-HooterDataEntryRowEntryNumber -Row $row) -ne $entryNumber) { continue }
        $rowScope = if ($Grid.Columns.Contains("Group")) { ConvertTo-HooterText $row.Cells["Group"].Value } else { $scope }
        if (-not $rowScope.Equals($scope, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        if ([string]::IsNullOrWhiteSpace($fieldKey) -or $keys.Contains($fieldKey)) { continue }
        [void]$keys.Add($fieldKey)
    }

    $sourcePosition = $keys.IndexOf($sourceKey)
    $targetPosition = $keys.IndexOf($targetKey)
    if ($sourcePosition -lt 0 -or $targetPosition -lt 0 -or $sourcePosition -eq $targetPosition) { return $false }
    $keys.RemoveAt($sourcePosition)
    if ($sourcePosition -lt $targetPosition) { $targetPosition-- }
    $keys.Insert($targetPosition, $sourceKey)

    for ($index = 0; $index -lt $keys.Count; $index++) {
        $script:FieldOrder[[string]$keys[$index]] = $index
    }
    Apply-HooterSavedFieldOrderToCatalog
    Request-HooterDeferredSettingsSave
    if ($script:Ui.ContainsKey("SettingsGrid") -and $script:Ui.SettingsGrid.Visible -and $script:Ui.SettingsGrid.Columns.Contains("FieldKey") -and $script:Ui.SettingsGrid.Columns.Contains("FieldOrder")) {
        foreach ($row in $script:Ui.SettingsGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
            if ($script:FieldOrder.ContainsKey($fieldKey)) {
                $row.Cells["FieldOrder"].Value = $script:FieldOrder[$fieldKey]
            }
        }
    }
    $targetGrid = switch ($scope) {
        "Plot" { if ($script:Ui.ContainsKey("PlotGrid")) { $script:Ui.PlotGrid } else { $Grid } }
        "Tree" { if ($script:Ui.ContainsKey("TreeGrid")) { $script:Ui.TreeGrid } else { $Grid } }
        "Regen" { if ($script:Ui.ContainsKey("RegenGrid")) { $script:Ui.RegenGrid } else { $Grid } }
        default { $Grid }
    }
    if (-not (Apply-HooterDataEntryFieldOrderInPlace -Grid $targetGrid -Scope $scope -SourceKey $sourceKey -TargetKey $targetKey -EntryNumber $entryNumber)) {
        Apply-HooterFieldOrderToLoadedPlot
    }
    if ($script:Ui.ContainsKey("StatusLabel")) {
        $script:Ui.StatusLabel.Text = "Moved $scope field order. Setup will save after you pause moving rows."
    }
    return $true
}

function Reset-HooterDataEntryDragState {
    $script:DataEntryDragSourceGrid = $null
    $script:DataEntryDragSourceIndex = -1
    $script:DataEntryDragStartPoint = $null
}

function Register-HooterDataEntryGridDragReorder {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return }
    $Grid.AllowDrop = $true
    $Grid.Add_MouseDown({
        param($Sender, $EventArgs)
        Stop-HooterDeferredSettingsSave
        Reset-HooterDataEntryDragState
        if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $hit = $Sender.HitTest($EventArgs.X, $EventArgs.Y)
        if ($hit.RowIndex -lt 0 -or $hit.RowIndex -ge $Sender.Rows.Count -or $Sender.Rows[$hit.RowIndex].IsNewRow) { return }
        if ($hit.ColumnIndex -lt 0 -or $hit.ColumnIndex -ge $Sender.Columns.Count) { return }
        $columnName = $Sender.Columns[$hit.ColumnIndex].Name
        if ($columnName -ne "MoveHandle" -and $columnName -ne "FieldLabel") { return }
        $script:DataEntryDragSourceGrid = $Sender
        $script:DataEntryDragSourceIndex = $hit.RowIndex
        $script:DataEntryDragStartPoint = New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)
    })
    $Grid.Add_MouseMove({
        param($Sender, $EventArgs)
        if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        if (-not [object]::ReferenceEquals($script:DataEntryDragSourceGrid, $Sender)) { return }
        if ($script:DataEntryDragSourceIndex -lt 0 -or $null -eq $script:DataEntryDragStartPoint) { return }
        $dx = [Math]::Abs($EventArgs.X - $script:DataEntryDragStartPoint.X)
        $dy = [Math]::Abs($EventArgs.Y - $script:DataEntryDragStartPoint.Y)
        if ($dx -lt [System.Windows.Forms.SystemInformation]::DragSize.Width -and $dy -lt [System.Windows.Forms.SystemInformation]::DragSize.Height) { return }
        Stop-HooterDeferredSettingsSave
        [void]$Sender.DoDragDrop("PlotHootDataEntryFieldRow", [System.Windows.Forms.DragDropEffects]::Move)
    })
    $Grid.Add_DragOver({
        param($Sender, $EventArgs)
        $EventArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
        if (-not [object]::ReferenceEquals($script:DataEntryDragSourceGrid, $Sender)) { return }
        $point = $Sender.PointToClient((New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)))
        $hit = $Sender.HitTest($point.X, $point.Y)
        if (Test-HooterDataEntryMoveTarget -Grid $Sender -SourceIndex $script:DataEntryDragSourceIndex -TargetIndex $hit.RowIndex) {
            $EventArgs.Effect = [System.Windows.Forms.DragDropEffects]::Move
        }
    })
    $Grid.Add_DragDrop({
        param($Sender, $EventArgs)
        try {
            if (-not [object]::ReferenceEquals($script:DataEntryDragSourceGrid, $Sender)) { return }
            $point = $Sender.PointToClient((New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)))
            $hit = $Sender.HitTest($point.X, $point.Y)
            if (-not (Set-HooterFieldOrderFromDataEntryMove -Grid $Sender -SourceIndex $script:DataEntryDragSourceIndex -TargetIndex $hit.RowIndex)) {
                Set-HooterStatus "Drag the Move handle within the same plot, tree, or regen entry to move that field."
            }
        }
        finally {
            Reset-HooterDataEntryDragState
        }
    })
    $Grid.Add_MouseUp({ Reset-HooterDataEntryDragState })
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "MoveHandle" -Text "Drag this handle up or down to move this field."
    Set-HooterGridColumnToolTip -Grid $Grid -ColumnName "FieldLabel" -Text "Field being checked. You can also drag this cell up or down to move this field order."
}

function Move-HooterSettingsFieldRow {
    param([int]$Offset)

    if (-not $script:Ui.ContainsKey("SettingsGrid")) { return }
    $grid = $script:Ui.SettingsGrid
    if ($null -eq $grid.CurrentRow -or $grid.CurrentRow.IsNewRow) { return }
    Move-HooterSettingsFieldRowToIndex -Grid $grid -SourceIndex $grid.CurrentRow.Index -TargetIndex ($grid.CurrentRow.Index + $Offset)
}

function Move-HooterSettingsFieldRowToIndex {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$SourceIndex,
        [int]$TargetIndex
    )

    if ($null -eq $Grid) { return $false }
    try { [void]$grid.EndEdit() } catch {}
    try { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } catch {}

    if ($SourceIndex -lt 0 -or $SourceIndex -ge $Grid.Rows.Count) { return $false }
    if ($TargetIndex -lt 0) { $TargetIndex = 0 }
    if ($TargetIndex -ge $Grid.Rows.Count) { $TargetIndex = $Grid.Rows.Count - 1 }
    if ($Grid.Rows[$SourceIndex].IsNewRow -or $Grid.Rows[$TargetIndex].IsNewRow -or $SourceIndex -eq $TargetIndex) { return $false }

    $sourceGroup = ConvertTo-HooterText $Grid.Rows[$SourceIndex].Cells["Group"].Value
    $targetGroup = ConvertTo-HooterText $Grid.Rows[$TargetIndex].Cells["Group"].Value
    if (-not $sourceGroup.Equals($targetGroup, [System.StringComparison]::OrdinalIgnoreCase)) {
        Set-HooterStatus "Fields can be moved within Plot, Tree, Regen, or Execution, but not across sections."
        return $false
    }

    $values = New-Object object[] $Grid.Columns.Count
    for ($columnIndex = 0; $columnIndex -lt $Grid.Columns.Count; $columnIndex++) {
        $values[$columnIndex] = $Grid.Rows[$SourceIndex].Cells[$columnIndex].Value
    }
    $insertIndex = $TargetIndex
    if ($SourceIndex -lt $TargetIndex) { $insertIndex-- }
    $Grid.Rows.RemoveAt($SourceIndex)
    $Grid.Rows.Insert($insertIndex, $values)
    $Grid.ClearSelection()
    $Grid.Rows[$insertIndex].Selected = $true
    try { $Grid.CurrentCell = $Grid.Rows[$insertIndex].Cells["FieldLabel"] } catch {}

    Set-HooterSettingsGridOrderFromRows -Grid $Grid
    Update-HooterSettingsFromGrid -UpdateFieldOrder
    Apply-HooterFieldOrderToLoadedPlot
    if ($sourceGroup.Equals("Execution", [System.StringComparison]::OrdinalIgnoreCase)) {
        Populate-HooterExecutionGrid
    }
    Set-HooterStatus "Moved field order. Click Save setup to keep this order for future plots."
    return $true
}

function Register-HooterSettingsGridDragReorder {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid) { return }
    $Grid.AllowDrop = $true
    $Grid.Add_MouseDown({
        param($Sender, $EventArgs)
        $script:SettingsDragSourceIndex = -1
        $script:SettingsDragStartPoint = $null
        if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $hit = $Sender.HitTest($EventArgs.X, $EventArgs.Y)
        if ($hit.RowIndex -lt 0 -or $hit.RowIndex -ge $Sender.Rows.Count -or $Sender.Rows[$hit.RowIndex].IsNewRow) { return }
        $script:SettingsDragSourceIndex = $hit.RowIndex
        $script:SettingsDragStartPoint = New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)
    })
    $Grid.Add_MouseMove({
        param($Sender, $EventArgs)
        if ($EventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        if ($script:SettingsDragSourceIndex -lt 0 -or $null -eq $script:SettingsDragStartPoint) { return }
        $dx = [Math]::Abs($EventArgs.X - $script:SettingsDragStartPoint.X)
        $dy = [Math]::Abs($EventArgs.Y - $script:SettingsDragStartPoint.Y)
        if ($dx -lt [System.Windows.Forms.SystemInformation]::DragSize.Width -and $dy -lt [System.Windows.Forms.SystemInformation]::DragSize.Height) { return }
        [void]$Sender.DoDragDrop("PlotHootSettingsFieldRow", [System.Windows.Forms.DragDropEffects]::Move)
    })
    $Grid.Add_DragOver({
        param($Sender, $EventArgs)
        $EventArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
        if ($script:SettingsDragSourceIndex -lt 0 -or $script:SettingsDragSourceIndex -ge $Sender.Rows.Count) { return }
        $point = $Sender.PointToClient((New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)))
        $hit = $Sender.HitTest($point.X, $point.Y)
        if ($hit.RowIndex -lt 0 -or $hit.RowIndex -ge $Sender.Rows.Count -or $Sender.Rows[$hit.RowIndex].IsNewRow) { return }
        $sourceGroup = ConvertTo-HooterText $Sender.Rows[$script:SettingsDragSourceIndex].Cells["Group"].Value
        $targetGroup = ConvertTo-HooterText $Sender.Rows[$hit.RowIndex].Cells["Group"].Value
        if ($sourceGroup.Equals($targetGroup, [System.StringComparison]::OrdinalIgnoreCase)) {
            $EventArgs.Effect = [System.Windows.Forms.DragDropEffects]::Move
        }
    })
    $Grid.Add_DragDrop({
        param($Sender, $EventArgs)
        try {
            $point = $Sender.PointToClient((New-Object System.Drawing.Point($EventArgs.X, $EventArgs.Y)))
            $hit = $Sender.HitTest($point.X, $point.Y)
            if ($hit.RowIndex -ge 0) {
                [void](Move-HooterSettingsFieldRowToIndex -Grid $Sender -SourceIndex $script:SettingsDragSourceIndex -TargetIndex $hit.RowIndex)
            }
        }
        finally {
            $script:SettingsDragSourceIndex = -1
            $script:SettingsDragStartPoint = $null
        }
    })
    $Grid.Add_MouseUp({
        $script:SettingsDragSourceIndex = -1
        $script:SettingsDragStartPoint = $null
    })
}

function Show-HooterCrewValuePopup {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [int]$RowIndex
    )

    if ($null -eq $Grid -or $RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count -or -not $Grid.Columns.Contains("CrewValue")) { return }
    $row = $Grid.Rows[$RowIndex]
    if ($row.IsNewRow) { return }
    $fieldLabel = if ($Grid.Columns.Contains("FieldLabel")) { ConvertTo-HooterText $row.Cells["FieldLabel"].Value } else { "Crew value" }
    $crewValue = ConvertTo-HooterText $row.Cells["CrewValue"].Value

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Crew value - $fieldLabel"
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dialog.MinimumSize = New-Object System.Drawing.Size(520, 320)
    $dialog.Size = New-Object System.Drawing.Size(760, 460)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    $dialog.Font = New-HooterFont 9.4

    $label = New-HooterLabel $fieldLabel 18 16 700 26 10.0 ([System.Drawing.FontStyle]::Bold)
    $label.Anchor = "Top,Left,Right"
    $dialog.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(18, 52)
    $textBox.Size = New-Object System.Drawing.Size(706, 318)
    $textBox.Anchor = "Top,Bottom,Left,Right"
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = "Both"
    $textBox.WordWrap = $true
    $textBox.Font = New-HooterFont 10.0
    $textBox.Text = $crewValue
    $dialog.Controls.Add($textBox)

    $closeButton = New-HooterButton "Close" 640 382 84 30
    $closeButton.Anchor = "Bottom,Right"
    $closeButton.Add_Click({ $dialog.Close() })
    $dialog.Controls.Add($closeButton)

    try {
        if ($script:Ui.ContainsKey("Form") -and $null -ne $script:Ui.Form) {
            [void]$dialog.ShowDialog($script:Ui.Form)
        }
        else {
            [void]$dialog.ShowDialog()
        }
    }
    finally {
        $dialog.Dispose()
    }
}

function Get-HooterPlotByDisplayText {
    param([string]$Text)

    $clean = ConvertTo-HooterText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    foreach ($plot in @($script:Plots)) {
        $plotLabel = ConvertTo-HooterText $plot.PlotLabel
        if ($plot.Display.Equals($clean, [System.StringComparison]::OrdinalIgnoreCase) -or
            $plot.PlotNumber.Equals($clean, [System.StringComparison]::OrdinalIgnoreCase) -or
            (-not [string]::IsNullOrWhiteSpace($plotLabel) -and $plotLabel.Equals($clean, [System.StringComparison]::OrdinalIgnoreCase))) {
            return $plot
        }
    }
    $firstToken = ($clean -split "\s+-\s+", 2)[0]
    foreach ($plot in @($script:Plots)) {
        if ($plot.PlotNumber.Equals($firstToken, [System.StringComparison]::OrdinalIgnoreCase)) { return $plot }
    }
    return [pscustomobject]@{ PlotNumber = $clean; Display = $clean }
}

function Set-HooterPlotLookupList {
    param(
        [System.Windows.Forms.ComboBox]$PlotBox,
        [object[]]$Plots
    )

    if ($null -eq $PlotBox) { return }
    $source = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $seen = New-InsensitiveHashtable

    $PlotBox.BeginUpdate()
    try {
        $PlotBox.Items.Clear()
        foreach ($plot in @($Plots)) {
            $display = ConvertTo-HooterText $plot.Display
            $plotNumber = ConvertTo-HooterText $plot.PlotNumber
            $plotLabel = ConvertTo-HooterText $plot.PlotLabel
            if (-not [string]::IsNullOrWhiteSpace($display)) {
                [void]$PlotBox.Items.Add($display)
                if (-not $seen.ContainsKey($display)) {
                    [void]$source.Add($display)
                    $seen[$display] = $true
                }
            }
            foreach ($value in @($plotNumber, $plotLabel)) {
                if (-not [string]::IsNullOrWhiteSpace($value) -and -not $seen.ContainsKey($value)) {
                    [void]$source.Add($value)
                    $seen[$value] = $true
                }
            }
        }
        $PlotBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
        $PlotBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource
        $PlotBox.AutoCompleteCustomSource = $source
        $PlotBox.SelectedIndex = -1
        $PlotBox.Text = ""
    }
    finally {
        $PlotBox.EndUpdate()
    }
}

function Connect-HooterDatabase {
    param([switch]$CompletedOnly)

    try {
        $path = Get-HooterDatabaseBoxPath
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
            throw "Choose a master database first."
        }
        $plotListMode = if ($CompletedOnly) { "completed plots" } else { "all plots" }
        Set-HooterStatus "Connecting to database and loading $plotListMode..."
        $connectionString = New-AccessConnectionString -Path $path
        $connection = Open-AccessConnection -ConnectionString $connectionString
        try {
            $script:DatabasePath = [System.IO.Path]::GetFullPath($path)
            Set-HooterDatabaseBoxPath -TextBox $script:Ui.DatabaseBox -Path $script:DatabasePath
            $script:ConnectionString = $connectionString
            $script:ProjectName = Get-HooterProjectNameFromDatabase -Connection $connection -DatabasePath $script:DatabasePath
            $script:FieldCatalog = Get-HooterFieldCatalog -Connection $connection
            $script:InventoryTotalPlotCount = Get-HooterPlotInventoryCount -Connection $connection
            $script:Plots = @(Get-HooterPlots -Connection $connection -CompletedOnly:$CompletedOnly)
            $script:PlotListMode = if ($CompletedOnly) { "Completed" } else { "All" }
        }
        finally {
            if ($connection.State -eq "Open") { $connection.Close() }
            $connection.Dispose()
        }

        Set-HooterPlotLookupList -PlotBox $script:Ui.PlotBox -Plots $script:Plots
        Update-HooterProjectDisplay
        Update-HooterInventoryProgress
        Populate-HooterSettingsGrid
        $projectText = if ([string]::IsNullOrWhiteSpace((Get-HooterCurrentProjectName))) { "Connected." } else { "Connected to $(Get-HooterCurrentProjectName)." }
        $fieldOrderText = if ($script:FieldOrder.Count -gt 0) { " Order: custom." } else { " Order: database/AppColumns." }
        if ($CompletedOnly) {
            Set-HooterStatus "$projectText Completed plots loaded: $($script:Plots.Count)/$($script:InventoryTotalPlotCount).$fieldOrderText Type plot #, then Enter or Load plot."
        }
        else {
            Set-HooterStatus "$projectText All plots loaded: $($script:Plots.Count).$fieldOrderText Type plot #, then Enter or Load plot."
        }
    }
    catch {
        $script:DatabasePath = ""
        $script:ConnectionString = ""
        $script:ProjectName = ""
        $script:Plots = @()
        $script:InventoryTotalPlotCount = 0
        Update-HooterProjectDisplay
        Set-HooterStatus "Could not connect."
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "PlotHoot connection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Clear-HooterCurrentGrids {
    foreach ($grid in @($script:Ui.PlotGrid, $script:Ui.TreeGrid, $script:Ui.RegenGrid)) {
        if ($null -ne $grid) { $grid.Rows.Clear() }
    }
    if ($script:Ui.ContainsKey("TreeRecordBox")) { $script:Ui.TreeRecordBox.Items.Clear() }
    if ($script:Ui.ContainsKey("RegenRecordBox")) { $script:Ui.RegenRecordBox.Items.Clear() }
    if ($script:Ui.ContainsKey("NotesBox")) { $script:Ui.NotesBox.Text = "" }
    Clear-HooterCheckCruiseMetadata
    if ($script:Ui.ContainsKey("MissedTreeGrid")) { $script:Ui.MissedTreeGrid.Rows.Clear() }
    if ($script:Ui.ContainsKey("CrewMissedTreeCheck")) {
        try {
            $script:SuppressMissedTreeCheckEvent = $true
            $script:Ui.CrewMissedTreeCheck.Checked = $false
        }
        finally {
            $script:SuppressMissedTreeCheckEvent = $false
        }
    }
    Populate-HooterExecutionGrid
    Update-HooterMissedTreeStatus
    Update-HooterTreeProgressLabel
    Update-HooterRegenProgressLabel
}

function Load-HooterPlot {
    $busyDialog = $null
    try {
        $requestedPlotText = ConvertTo-HooterText $script:Ui.PlotBox.Text
        if ([string]::IsNullOrWhiteSpace($script:ConnectionString)) {
            $busyDialog = Show-HooterBusyDialog -Title "Load plot" -Message "Connecting to database..."
            Connect-HooterDatabase -CompletedOnly
            if ([string]::IsNullOrWhiteSpace($script:ConnectionString)) { return }
            if (-not [string]::IsNullOrWhiteSpace($requestedPlotText)) {
                $script:Ui.PlotBox.Text = $requestedPlotText
            }
        }
        $plot = Get-HooterPlotByDisplayText -Text $script:Ui.PlotBox.Text
        if ($null -eq $plot -or [string]::IsNullOrWhiteSpace($plot.PlotNumber)) { throw "Type a plot number or choose one from the list." }
        if ($null -eq $busyDialog) {
            $busyDialog = Show-HooterBusyDialog -Title "Load plot" -Message "Loading plot $($plot.PlotNumber)..."
        }
        Set-HooterStatus "Loading plot $($plot.PlotNumber)..."
        $connection = Open-AccessConnection -ConnectionString $script:ConnectionString
        try {
            $script:CurrentPlot = Get-HooterPlotData -Connection $connection -PlotNumber $plot.PlotNumber
        }
        finally {
            if ($connection.State -eq "Open") { $connection.Close() }
            $connection.Dispose()
        }
        $script:CurrentSessionId = ([guid]::NewGuid()).ToString()
        Clear-HooterCurrentGrids

        foreach ($field in @($script:FieldCatalog.Plot)) {
            Add-HooterCheckRow -Grid $script:Ui.PlotGrid -Field $field -CrewValue (Get-RecordFieldValue -Record $script:CurrentPlot -Field $field)
        }
        foreach ($treeRecord in @($script:CurrentPlot.TreeRecords)) {
            [void]$script:Ui.TreeRecordBox.Items.Add($treeRecord.Display)
        }
        if ($script:Ui.TreeRecordBox.Items.Count -gt 0) { $script:Ui.TreeRecordBox.SelectedIndex = 0 }
        Add-HooterAllTreeEntries
        foreach ($regenRecord in @($script:CurrentPlot.RegenRecords)) {
            [void]$script:Ui.RegenRecordBox.Items.Add($regenRecord.Display)
        }
        if ($script:Ui.RegenRecordBox.Items.Count -gt 0) { $script:Ui.RegenRecordBox.SelectedIndex = 0 }
        Add-HooterAllRegenEntries

        Apply-HooterExistingSessionForCurrentPlot
        Refresh-HooterAllStatuses
        Set-HooterTreeVisibleRows
        Set-HooterRegenVisibleRows
        $loadNotes = New-Object System.Collections.Generic.List[string]
        if ($script:CurrentPlot.TreeRecords.Count -gt 0 -and $script:FieldCatalog.Tree.Count -eq 0) { [void]$loadNotes.Add("tree records found, but no active tree fields were found in AppColumns") }
        if ($script:CurrentPlot.RegenRecords.Count -gt 0 -and $script:FieldCatalog.Regen.Count -eq 0) { [void]$loadNotes.Add("regen records found, but no active regen fields were found in AppColumns") }
        $noteText = if ($loadNotes.Count -gt 0) { " Note: " + ([string]::Join("; ", [string[]]$loadNotes.ToArray())) + "." } else { "" }
        $metadataReminder = if ([string]::IsNullOrWhiteSpace((Get-HooterCheckCruiserName))) { " Add the check cruiser name on Review / Export before saving." } else { "" }
        Set-HooterStatus "Loaded plot $($script:CurrentPlot.PlotNumber). Tree records: $($script:CurrentPlot.TreeRecords.Count). Regen records: $($script:CurrentPlot.RegenRecords.Count).$noteText$metadataReminder"
    }
    catch {
        Close-HooterBusyDialog -Dialog $busyDialog
        $busyDialog = $null
        Set-HooterStatus "Could not load plot."
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Load plot", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    finally {
        Close-HooterBusyDialog -Dialog $busyDialog
    }
}

function Get-HooterNextEntryNumber {
    param([System.Windows.Forms.DataGridView]$Grid)

    $max = 0
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $value = 0
        if ([int]::TryParse((ConvertTo-HooterText $row.Cells["EntryNumber"].Value), [ref]$value) -and $value -gt $max) {
            $max = $value
        }
    }
    return ($max + 1)
}

function Add-HooterRecordEntry {
    param(
        [string]$Kind
    )

    if ($null -eq $script:CurrentPlot) {
        [System.Windows.Forms.MessageBox]::Show("Load a plot before adding $Kind QA entries.", "PlotHoot", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if ($Kind -eq "Tree") {
        $records = @($script:CurrentPlot.TreeRecords)
        $box = $script:Ui.TreeRecordBox
        $grid = $script:Ui.TreeGrid
        $fields = @($script:FieldCatalog.Tree)
    }
    else {
        $records = @($script:CurrentPlot.RegenRecords)
        $box = $script:Ui.RegenRecordBox
        $grid = $script:Ui.RegenGrid
        $fields = @($script:FieldCatalog.Regen)
    }
    if ($records.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No $Kind records were found for this plot.", "PlotHoot", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if ($fields.Count -eq 0 -and $Kind -ne "Tree") {
        [System.Windows.Forms.MessageBox]::Show("No active $Kind fields were found in AppColumns.", "PlotHoot", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $index = if ($box.SelectedIndex -ge 0) { $box.SelectedIndex } else { 0 }
    $record = $records[$index]
    $entryNumber = Get-HooterNextEntryNumber -Grid $grid
    if ($Kind -eq "Tree") {
        Add-HooterCheckRow -Grid $grid -Field (Get-HooterTreeFoundField) -CrewValue "Pass" -EntryNumber $entryNumber -CrewRecord $record.Display
    }
    foreach ($field in $fields) {
        Add-HooterCheckRow -Grid $grid -Field $field -CrewValue (Get-RecordFieldValue -Record $record -Field $field) -EntryNumber $entryNumber -CrewRecord $record.Display
    }
    Refresh-HooterGridStatuses -Grid $grid
    Update-HooterOverallLabel
}

function Add-HooterAllTreeEntries {
    if ($null -eq $script:CurrentPlot) { return }
    if (-not $script:Ui.ContainsKey("TreeGrid")) { return }

    $grid = $script:Ui.TreeGrid
    $grid.Rows.Clear()
    $records = @($script:CurrentPlot.TreeRecords)
    $fields = @(Get-HooterTreeEntryFields)
    if ($records.Count -eq 0) { return }

    $entryNumber = 1
    foreach ($record in $records) {
        foreach ($field in $fields) {
            $crewValue = if ((ConvertTo-HooterText $field.FieldKey) -eq (ConvertTo-HooterText (Get-HooterTreeFoundField).FieldKey)) { "Pass" } else { Get-RecordFieldValue -Record $record -Field $field }
            Add-HooterCheckRow -Grid $grid -Field $field -CrewValue $crewValue -EntryNumber $entryNumber -CrewRecord $record.Display
        }
        $entryNumber++
    }
}

function Get-HooterVisibleTreeEntryNumber {
    if (-not $script:Ui.ContainsKey("TreeRecordBox")) { return 0 }
    if ($script:Ui.TreeRecordBox.SelectedIndex -lt 0) { return 0 }
    return ($script:Ui.TreeRecordBox.SelectedIndex + 1)
}

function Update-HooterTreeProgressLabel {
    if (-not $script:Ui.ContainsKey("TreeProgressLabel")) { return }
    $entryNumber = Get-HooterVisibleTreeEntryNumber
    $treeCount = if ($script:Ui.ContainsKey("TreeRecordBox")) { $script:Ui.TreeRecordBox.Items.Count } else { 0 }
    if ($treeCount -le 0 -or $entryNumber -le 0) {
        $script:Ui.TreeProgressLabel.Text = "No tree records loaded."
        Update-HooterTreeMaxSummaryLabel
        return
    }

    $failed = 0
    foreach ($row in $script:Ui.TreeGrid.Rows) {
        if ($row.IsNewRow) { continue }
        if ((ConvertTo-HooterText $row.Cells["EntryNumber"].Value) -ne ([string]$entryNumber)) { continue }
        if ((ConvertTo-HooterText $row.Cells["Status"].Value) -eq "Fail") { $failed++ }
    }
    $treeStats = Get-HooterGridCheckStats -Grid $script:Ui.TreeGrid -Scope "Tree"
    $treeMaxSummary = Get-HooterTreeMaxLossSummary
    $script:Ui.TreeProgressLabel.Text = "Tree $entryNumber/$treeCount - $failed fail; section loss $(Format-HooterScoreNumber $treeStats.LostPoints)/$(Format-HooterScoreNumber $treeStats.PossiblePoints); max/tree $($treeMaxSummary.PerTreeText); total max $($treeMaxSummary.TotalText)"
    Update-HooterTreeMaxSummaryLabel
}

function Set-HooterTreeVisibleRows {
    if (-not $script:Ui.ContainsKey("TreeGrid") -or -not $script:Ui.ContainsKey("TreeRecordBox")) { return }
    $grid = $script:Ui.TreeGrid
    $entryNumber = Get-HooterVisibleTreeEntryNumber
    $showSelectedOnly = $true
    if ($script:Ui.ContainsKey("TreeSelectedOnlyCheck")) {
        $showSelectedOnly = [bool]$script:Ui.TreeSelectedOnlyCheck.Checked
    }

    try { $grid.CurrentCell = $null } catch {}
    $firstVisibleCell = $null
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $rowEntry = ConvertTo-HooterText $row.Cells["EntryNumber"].Value
        $visible = (-not $showSelectedOnly) -or ($entryNumber -le 0) -or ($rowEntry -eq ([string]$entryNumber))
        try { $row.Visible = $visible } catch {}
        if ($visible -and $null -eq $firstVisibleCell) {
            $firstVisibleCell = $row.Cells["QaValue"]
        }
    }
    if ($null -ne $firstVisibleCell) {
        try { $grid.CurrentCell = $firstVisibleCell } catch {}
    }
    Update-HooterTreeProgressLabel
}

function Select-HooterTreeOffset {
    param([int]$Offset)

    if (-not $script:Ui.ContainsKey("TreeRecordBox")) { return }
    $box = $script:Ui.TreeRecordBox
    if ($box.Items.Count -eq 0) { return }
    $index = $box.SelectedIndex
    if ($index -lt 0) { $index = 0 }
    $index = [Math]::Max(0, [Math]::Min($box.Items.Count - 1, $index + $Offset))
    $box.SelectedIndex = $index
    Set-HooterTreeVisibleRows
}

function Add-HooterAllRegenEntries {
    if ($null -eq $script:CurrentPlot) { return }
    if (-not $script:Ui.ContainsKey("RegenGrid")) { return }

    $grid = $script:Ui.RegenGrid
    $grid.Rows.Clear()
    $records = @($script:CurrentPlot.RegenRecords)
    $fields = @($script:FieldCatalog.Regen)
    if ($records.Count -eq 0 -or $fields.Count -eq 0) { return }

    $entryNumber = 1
    foreach ($record in $records) {
        foreach ($field in $fields) {
            Add-HooterCheckRow -Grid $grid -Field $field -CrewValue (Get-RecordFieldValue -Record $record -Field $field) -EntryNumber $entryNumber -CrewRecord $record.Display
        }
        $entryNumber++
    }
}

function Get-HooterVisibleRegenEntryNumber {
    if (-not $script:Ui.ContainsKey("RegenRecordBox")) { return 0 }
    if ($script:Ui.RegenRecordBox.SelectedIndex -lt 0) { return 0 }
    return ($script:Ui.RegenRecordBox.SelectedIndex + 1)
}

function Update-HooterRegenProgressLabel {
    if (-not $script:Ui.ContainsKey("RegenProgressLabel")) { return }
    $entryNumber = Get-HooterVisibleRegenEntryNumber
    $regenCount = if ($script:Ui.ContainsKey("RegenRecordBox")) { $script:Ui.RegenRecordBox.Items.Count } else { 0 }
    if ($regenCount -le 0) {
        $script:Ui.RegenProgressLabel.Text = "No regen records loaded."
        return
    }

    $showSelectedOnly = $false
    if ($script:Ui.ContainsKey("RegenSelectedOnlyCheck")) {
        $showSelectedOnly = [bool]$script:Ui.RegenSelectedOnlyCheck.Checked
    }

    $failed = 0
    foreach ($row in $script:Ui.RegenGrid.Rows) {
        if ($row.IsNewRow) { continue }
        if ($showSelectedOnly -and (ConvertTo-HooterText $row.Cells["EntryNumber"].Value) -ne ([string]$entryNumber)) { continue }
        if ((ConvertTo-HooterText $row.Cells["Status"].Value) -eq "Fail") { $failed++ }
    }
    $regenStats = Get-HooterGridCheckStats -Grid $script:Ui.RegenGrid -Scope "Regen"
    if ($showSelectedOnly -and $entryNumber -gt 0) {
        $script:Ui.RegenProgressLabel.Text = "Regen $entryNumber/$regenCount - $failed fail; section loss $(Format-HooterScoreNumber $regenStats.LostPoints)/$(Format-HooterScoreNumber $regenStats.PossiblePoints)"
    }
    else {
        $script:Ui.RegenProgressLabel.Text = "All regen records visible: $regenCount record(s), $failed fail; section loss $(Format-HooterScoreNumber $regenStats.LostPoints)/$(Format-HooterScoreNumber $regenStats.PossiblePoints)"
    }
}

function Set-HooterRegenVisibleRows {
    if (-not $script:Ui.ContainsKey("RegenGrid") -or -not $script:Ui.ContainsKey("RegenRecordBox")) { return }
    $grid = $script:Ui.RegenGrid
    $entryNumber = Get-HooterVisibleRegenEntryNumber
    $showSelectedOnly = $false
    if ($script:Ui.ContainsKey("RegenSelectedOnlyCheck")) {
        $showSelectedOnly = [bool]$script:Ui.RegenSelectedOnlyCheck.Checked
    }

    try { $grid.CurrentCell = $null } catch {}
    $firstVisibleCell = $null
    $selectedEntryCell = $null
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        $rowEntry = ConvertTo-HooterText $row.Cells["EntryNumber"].Value
        $visible = (-not $showSelectedOnly) -or ($entryNumber -le 0) -or ($rowEntry -eq ([string]$entryNumber))
        try { $row.Visible = $visible } catch {}
        if ($visible -and $null -eq $firstVisibleCell) {
            $firstVisibleCell = $row.Cells["QaValue"]
        }
        if ($visible -and $entryNumber -gt 0 -and $rowEntry -eq ([string]$entryNumber) -and $null -eq $selectedEntryCell) {
            $selectedEntryCell = $row.Cells["QaValue"]
        }
    }
    $targetCell = if ($null -ne $selectedEntryCell) { $selectedEntryCell } else { $firstVisibleCell }
    if ($null -ne $targetCell) {
        try { $grid.CurrentCell = $targetCell } catch {}
        try { $grid.FirstDisplayedScrollingRowIndex = $targetCell.RowIndex } catch {}
    }
    Update-HooterRegenProgressLabel
}

function Select-HooterRegenOffset {
    param([int]$Offset)

    if (-not $script:Ui.ContainsKey("RegenRecordBox")) { return }
    $box = $script:Ui.RegenRecordBox
    if ($box.Items.Count -eq 0) { return }
    $index = $box.SelectedIndex
    if ($index -lt 0) { $index = 0 }
    $index = [Math]::Max(0, [Math]::Min($box.Items.Count - 1, $index + $Offset))
    $box.SelectedIndex = $index
    Set-HooterRegenVisibleRows
}

function Remove-HooterSelectedEntry {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid -or -not $Grid.Columns.Contains("EntryNumber") -or $Grid.SelectedRows.Count -eq 0) { return }
    $entries = @{}
    foreach ($row in $Grid.SelectedRows) {
        if ($row.IsNewRow) { continue }
        $entries[(ConvertTo-HooterText $row.Cells["EntryNumber"].Value)] = $true
    }
    for ($i = $Grid.Rows.Count - 1; $i -ge 0; $i--) {
        $row = $Grid.Rows[$i]
        if ($row.IsNewRow) { continue }
        $entry = ConvertTo-HooterText $row.Cells["EntryNumber"].Value
        if ($entries.ContainsKey($entry)) {
            $Grid.Rows.RemoveAt($i)
        }
    }
    Update-HooterOverallLabel
}

function Get-HooterRowsFromGrid {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Scope
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $entryNumber = if ($Grid.Columns.Contains("EntryNumber")) { ConvertTo-HooterText $row.Cells["EntryNumber"].Value } else { "1" }
        $crewRecord = if ($Grid.Columns.Contains("CrewRecord")) { ConvertTo-HooterText $row.Cells["CrewRecord"].Value } else { $script:CurrentPlot.PlotNumber }
        $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
        $fieldOrder = if ($script:FieldOrder.ContainsKey($fieldKey)) { ConvertTo-HooterText $script:FieldOrder[$fieldKey] } else { "" }
        $tolerance = Get-HooterRowTolerance -Row $row
        $status = ConvertTo-HooterText $row.Cells["Status"].Value
        if ([string]::IsNullOrWhiteSpace($status)) { $status = "Not checked" }
        $score = Get-HooterRowScoreResult -Status $status -Tolerance $tolerance -CrewValue $row.Cells["CrewValue"].Value -QaValue $row.Cells["QaValue"].Value
        [void]$rows.Add([pscustomobject]@{
            Scope = $Scope
            EntryNumber = $entryNumber
            CrewRecord = $crewRecord
            TableName = ConvertTo-HooterText $row.Cells["TableName"].Value
            FieldName = ConvertTo-HooterText $row.Cells["FieldName"].Value
            FieldLabel = ConvertTo-HooterText $row.Cells["FieldLabel"].Value
            FieldKey = $fieldKey
            FieldOrder = $fieldOrder
            CrewValue = ConvertTo-HooterText $row.Cells["CrewValue"].Value
            QaValue = ConvertTo-HooterText $row.Cells["QaValue"].Value
            RuleMode = ConvertTo-HooterText $tolerance.Mode
            ToleranceValue = ConvertTo-HooterText $tolerance.Value
            Status = $status
            PointValue = Format-HooterScoreNumber $score.PointValue
            CriticalFail = $score.CriticalFail
            EarnedPoints = Format-HooterScoreNumber $score.EarnedPoints
            PossiblePoints = Format-HooterScoreNumber $score.PossiblePoints
            LostPoints = Format-HooterScoreNumber $score.LostPoints
            CriticalFailure = $score.CriticalFailure
            FieldNotes = ConvertTo-HooterText $row.Cells["Notes"].Value
        })
    }
    return @($rows.ToArray())
}

function Get-HooterRowsFromExecutionGrid {
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $script:Ui.ContainsKey("ExecutionGrid")) { return @() }
    foreach ($row in $script:Ui.ExecutionGrid.Rows) {
        if ($row.IsNewRow) { continue }
        $result = Get-HooterExecutionRowResult -Row $row
        $status = ConvertTo-HooterText $result.Status
        if ([string]::IsNullOrWhiteSpace($status)) { $status = "Not checked" }
        $fieldKey = Get-HooterFieldKey -Group "Execution" -TableName "TableD" -FieldName (ConvertTo-HooterText $row.Cells["ItemKey"].Value)
        $tolerance = Get-HooterTolerance -Field $fieldKey
        $rule = Get-HooterExecutionScoreRule -ItemKey (ConvertTo-HooterText $row.Cells["ItemKey"].Value)
        [void]$rows.Add([pscustomobject]@{
            Scope = "Execution"
            EntryNumber = "1"
            CrewRecord = if ($null -ne $script:CurrentPlot) { $script:CurrentPlot.PlotNumber } else { "" }
            TableName = "TableD"
            FieldName = ConvertTo-HooterText $row.Cells["ItemKey"].Value
            FieldLabel = ConvertTo-HooterText $row.Cells["Item"].Value
            FieldKey = $fieldKey
            CrewValue = ""
            QaValue = ConvertTo-HooterText $row.Cells["Rating"].Value
            RuleMode = ConvertTo-HooterText $tolerance.Mode
            ToleranceValue = ConvertTo-HooterText $rule.ToleranceValue
            Status = $status
            PointValue = Format-HooterScoreNumber $rule.PossiblePoints
            CriticalFail = [bool]$rule.CriticalOnPoor
            EarnedPoints = Format-HooterScoreNumber $result.EarnedPoints
            PossiblePoints = Format-HooterScoreNumber $result.PossiblePoints
            LostPoints = Format-HooterScoreNumber $result.PointLoss
            CriticalFailure = $result.CriticalFailure
            FieldNotes = ConvertTo-HooterText $row.Cells["Notes"].Value
        })
    }
    return @($rows.ToArray())
}

function Get-HooterRowsFromMissedTreeGrid {
    $rows = New-Object System.Collections.Generic.List[object]
    $plotNumber = if ($null -ne $script:CurrentPlot) { ConvertTo-HooterText $script:CurrentPlot.PlotNumber } else { "" }
    $fieldKey = Get-HooterFieldKey -Group "MissedTree" -TableName "PlotHoot QA" -FieldName "CrewMissedTree"

    if ($script:Ui.ContainsKey("MissedTreeGrid")) {
        foreach ($row in $script:Ui.MissedTreeGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $entryNumber = ConvertTo-HooterText $row.Cells["EntryNumber"].Value
            $dbh = ConvertTo-HooterText $row.Cells["DBH"].Value
            $distance = ConvertTo-HooterText $row.Cells["Distance"].Value
            $azimuth = ConvertTo-HooterText $row.Cells["Azimuth"].Value
            $notes = ConvertTo-HooterText $row.Cells["Notes"].Value
            $qaValue = "DBH=$dbh; Distance=$distance; Azimuth=$azimuth"
            [void]$rows.Add([pscustomobject]@{
                Scope = "MissedTree"
                EntryNumber = $entryNumber
                CrewRecord = $plotNumber
                TableName = "PlotHoot QA"
                FieldName = "CrewMissedTree"
                FieldLabel = "Crew missed tree found by QA"
                FieldKey = $fieldKey
                CrewValue = "Not recorded by crew"
                QaValue = $qaValue
                RuleMode = "Critical"
                ToleranceValue = "Automatic plot fail"
                PointValue = "0"
                CriticalFail = $true
                Status = "Fail"
                EarnedPoints = "0"
                PossiblePoints = "0"
                LostPoints = "0"
                CriticalFailure = $true
                FieldNotes = $notes
                MissedTreeDBH = $dbh
                MissedTreeDistance = $distance
                MissedTreeAzimuth = $azimuth
            })
        }
    }

    if ($rows.Count -eq 0 -and (Test-HooterCrewMissedTreeFail)) {
        [void]$rows.Add([pscustomobject]@{
            Scope = "MissedTree"
            EntryNumber = "1"
            CrewRecord = $plotNumber
            TableName = "PlotHoot QA"
            FieldName = "CrewMissedTreeFlag"
            FieldLabel = "Crew missed tree(s) found by QA"
            FieldKey = Get-HooterFieldKey -Group "MissedTree" -TableName "PlotHoot QA" -FieldName "CrewMissedTreeFlag"
            CrewValue = "Not recorded by crew"
            QaValue = "Yes"
            RuleMode = "Critical"
            ToleranceValue = "Automatic plot fail"
            PointValue = "0"
            CriticalFail = $true
            Status = "Fail"
            EarnedPoints = "0"
            PossiblePoints = "0"
            LostPoints = "0"
            CriticalFailure = $true
            FieldNotes = "Crew missed tree checkbox was marked, but no missed-tree attributes were entered."
            MissedTreeDBH = ""
            MissedTreeDistance = ""
            MissedTreeAzimuth = ""
        })
    }

    return @($rows.ToArray())
}

function Get-HooterCheckCruiserName {
    if ($script:Ui.ContainsKey("CheckCruiserBox") -and $null -ne $script:Ui.CheckCruiserBox) {
        return ConvertTo-HooterText $script:Ui.CheckCruiserBox.Text
    }
    return ""
}

function Get-HooterCheckCruiseDateText {
    if ($script:Ui.ContainsKey("CheckCruiseDatePicker") -and $null -ne $script:Ui.CheckCruiseDatePicker -and [bool]$script:Ui.CheckCruiseDatePicker.Checked) {
        return $script:Ui.CheckCruiseDatePicker.Value.ToString("yyyy-MM-dd")
    }
    return ""
}

function Clear-HooterCheckCruiseMetadata {
    if ($script:Ui.ContainsKey("CheckCruiserBox") -and $null -ne $script:Ui.CheckCruiserBox) {
        $script:Ui.CheckCruiserBox.Text = ""
    }
    if ($script:Ui.ContainsKey("CheckCruiseDatePicker") -and $null -ne $script:Ui.CheckCruiseDatePicker) {
        $script:Ui.CheckCruiseDatePicker.Checked = $true
        $script:Ui.CheckCruiseDatePicker.Value = Get-Date
    }
}

function Set-HooterCheckCruiseMetadata {
    param([object]$Session)

    if ($null -eq $Session) {
        Clear-HooterCheckCruiseMetadata
        return
    }

    if ($script:Ui.ContainsKey("CheckCruiserBox") -and $null -ne $script:Ui.CheckCruiserBox) {
        $script:Ui.CheckCruiserBox.Text = if ($null -ne $Session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $Session.CheckCruiserName } else { "" }
    }
    if ($script:Ui.ContainsKey("CheckCruiseDatePicker") -and $null -ne $script:Ui.CheckCruiseDatePicker) {
        $dateText = if ($null -ne $Session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $Session.CheckCruiseDate } else { "" }
        if ([string]::IsNullOrWhiteSpace($dateText)) {
            $script:Ui.CheckCruiseDatePicker.Checked = $true
            $script:Ui.CheckCruiseDatePicker.Value = Get-Date
        }
        else {
            try {
                $script:Ui.CheckCruiseDatePicker.Value = [datetime]::Parse($dateText)
                $script:Ui.CheckCruiseDatePicker.Checked = $true
            }
            catch {
                $script:Ui.CheckCruiseDatePicker.Checked = $true
                $script:Ui.CheckCruiseDatePicker.Value = Get-Date
            }
        }
    }
}

function New-HooterSessionFromCurrent {
    if ($null -eq $script:CurrentPlot) { throw "Load a plot before saving QA." }
    Refresh-HooterAllStatuses
    $rows = @(
        Get-HooterRowsFromGrid -Grid $script:Ui.PlotGrid -Scope "Plot"
        Get-HooterRowsFromGrid -Grid $script:Ui.TreeGrid -Scope "Tree"
        Get-HooterRowsFromGrid -Grid $script:Ui.RegenGrid -Scope "Regen"
        Get-HooterRowsFromExecutionGrid
        Get-HooterRowsFromMissedTreeGrid
    )
    $failed = @($rows | Where-Object { $_.Status -eq "Fail" }).Count
    $checked = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.QaValue) }).Count
    $overall = Get-HooterOverallStatusFromGrids
    return [pscustomobject]@{
        SessionID = if ([string]::IsNullOrWhiteSpace($script:CurrentSessionId)) { ([guid]::NewGuid()).ToString() } else { $script:CurrentSessionId }
        SavedAt = (Get-Date).ToString("s")
        Database = $script:DatabasePath
        ProjectName = Get-HooterCurrentProjectName
        CheckCruiserName = Get-HooterCheckCruiserName
        CheckCruiseDate = Get-HooterCheckCruiseDateText
        PlotNumber = $script:CurrentPlot.PlotNumber
        UTMEasting = $script:CurrentPlot.UTMEasting
        UTMNorthing = $script:CurrentPlot.UTMNorthing
        UTMZone = $script:CurrentPlot.UTMZone
        OverallStatus = $overall.Status
        CheckedCount = $checked
        FailedCount = $failed
        UncheckedCount = $overall.UncheckedCount
        MissedTreeCount = $overall.MissedTreeCount
        ScoreEarned = Format-HooterScoreNumber $overall.EarnedPoints
        ScorePossible = Format-HooterScoreNumber $overall.PossiblePoints
        ScoreLost = Format-HooterScoreNumber $overall.LostPoints
        ScorePercent = Format-HooterScoreNumber $overall.ScorePercent
        ScorePassPercent = Format-HooterScoreNumber $overall.ScorePassPercent
        MaxPointLoss = if ($overall.MaxPointLossConfigured) { Format-HooterScoreNumber $overall.MaxPointLoss } else { "" }
        PlotPointTotal = Format-HooterScoreNumber $overall.PlotPointTotal
        TreePointTotal = Format-HooterScoreNumber $overall.TreePointTotal
        RegenPointTotal = Format-HooterScoreNumber $overall.RegenPointTotal
        ExecutionPointTotal = Format-HooterScoreNumber $overall.ExecutionPointTotal
        PlotPointLoss = Format-HooterScoreNumber $overall.PlotPointLoss
        TreePointLoss = Format-HooterScoreNumber $overall.TreePointLoss
        RegenPointLoss = Format-HooterScoreNumber $overall.RegenPointLoss
        ExecutionPointLoss = Format-HooterScoreNumber $overall.ExecutionPointLoss
        CriticalFailCount = $overall.CriticalFailures
        OverallNotes = ConvertTo-HooterText $script:Ui.NotesBox.Text
        Rows = @($rows)
    }
}

function Save-HooterCurrentSession {
    try {
        if ([string]::IsNullOrWhiteSpace((Get-HooterCheckCruiserName))) {
            if ($script:Ui.ContainsKey("TabControl") -and $script:Ui.ContainsKey("ReviewPage")) {
                try { $script:Ui.TabControl.SelectedTab = $script:Ui.ReviewPage } catch {}
            }
            if ($script:Ui.ContainsKey("CheckCruiserBox")) {
                try { $script:Ui.CheckCruiserBox.Focus() } catch {}
            }
            throw "Enter the check cruiser name before saving QA."
        }
        if ($script:Ui.ContainsKey("CheckCruiseDatePicker") -and $null -ne $script:Ui.CheckCruiseDatePicker) {
            $script:Ui.CheckCruiseDatePicker.Checked = $true
        }
        $session = New-HooterSessionFromCurrent
        for ($i = $script:QaSessions.Count - 1; $i -ge 0; $i--) {
            $existing = $script:QaSessions[$i]
            if ($existing.PlotNumber.Equals($session.PlotNumber, [System.StringComparison]::OrdinalIgnoreCase) -and
                (ConvertTo-HooterText $existing.Database).Equals((ConvertTo-HooterText $session.Database), [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:QaSessions.RemoveAt($i)
            }
        }
        [void]$script:QaSessions.Add($session)
        Update-HooterReviewGrid
        if ((ConvertTo-HooterText $session.OverallStatus) -eq "Incomplete") {
            Set-HooterStatus "Saved QA for plot $($session.PlotNumber): Incomplete - $($session.UncheckedCount) unchecked row(s) still need QA values."
        }
        else {
            Set-HooterStatus "Saved QA for plot $($session.PlotNumber): $($session.OverallStatus), total error $($session.ScoreLost), $(Get-HooterMaxPointLossSummaryText -Value $session.MaxPointLoss)."
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Save QA", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Update-HooterReviewGrid {
    if (-not $script:Ui.ContainsKey("ReviewGrid")) { return }
    $grid = $script:Ui.ReviewGrid
    $grid.Rows.Clear()
    foreach ($session in @($script:QaSessions.ToArray() | Sort-Object PlotNumber)) {
        $utm = @($session.UTMEasting, $session.UTMNorthing, $session.UTMZone) | Where-Object { -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $_)) }
        $missedCount = 0
        if ($null -ne $session.PSObject.Properties["MissedTreeCount"]) {
            [void][int]::TryParse((ConvertTo-HooterText $session.MissedTreeCount), [ref]$missedCount)
        }
        if ($missedCount -le 0) {
            $missedCount = @($session.Rows | Where-Object { $_.Scope -eq "MissedTree" -and $_.FieldName -eq "CrewMissedTree" }).Count
        }
        $missedText = if ($missedCount -gt 0) { " Missed trees $missedCount" } else { "" }
        $uncheckedCount = 0
        if ($null -ne $session.PSObject.Properties["UncheckedCount"]) {
            [void][int]::TryParse((ConvertTo-HooterText $session.UncheckedCount), [ref]$uncheckedCount)
        }
        else {
            $checkedCount = 0
            [void][int]::TryParse((ConvertTo-HooterText $session.CheckedCount), [ref]$checkedCount)
            $uncheckedCount = [Math]::Max(0, @($session.Rows).Count - $checkedCount)
        }
        $scoreText = "total error $(Format-HooterScoreNumber $session.ScoreLost); $(Get-HooterMaxPointLossSummaryText -Value $session.MaxPointLoss); max possible $(Format-HooterScoreNumber $session.ScorePossible); Plot max $(Format-HooterScoreNumber $session.PlotPointTotal) Tree max effective $(Format-HooterScoreNumber $session.TreePointTotal) Regen max $(Format-HooterScoreNumber $session.RegenPointTotal) Table D max $(Format-HooterScoreNumber $session.ExecutionPointTotal)"
        $scoreText = "$scoreText$missedText"
        $checkCruiserName = if ($null -ne $session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $session.CheckCruiserName } else { "" }
        $checkCruiseDate = if ($null -ne $session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $session.CheckCruiseDate } else { "" }
        [void]$grid.Rows.Add($session.PlotNumber, $session.OverallStatus, $scoreText, $session.CriticalFailCount, $session.FailedCount, $uncheckedCount, ([string]::Join(" / ", [string[]]$utm)), $checkCruiserName, $checkCruiseDate, $session.OverallNotes, $session.SavedAt, $session.SessionID)
    }
    Update-HooterInventoryProgress
}

function Apply-HooterExistingSessionForCurrentPlot {
    if ($null -eq $script:CurrentPlot) { return }
    $match = $null
    foreach ($session in @($script:QaSessions.ToArray())) {
        if ($session.PlotNumber.Equals($script:CurrentPlot.PlotNumber, [System.StringComparison]::OrdinalIgnoreCase) -and
            (ConvertTo-HooterText $session.Database).Equals((ConvertTo-HooterText $script:DatabasePath), [System.StringComparison]::OrdinalIgnoreCase)) {
            $match = $session
            break
        }
    }
    if ($null -eq $match) { return }
    $script:CurrentSessionId = $match.SessionID
    $script:Ui.NotesBox.Text = ConvertTo-HooterText $match.OverallNotes
    Set-HooterCheckCruiseMetadata -Session $match

    $treeRows = @($match.Rows | Where-Object { $_.Scope -eq "Tree" })
    $regenRows = @($match.Rows | Where-Object { $_.Scope -eq "Regen" })
    if ($script:Ui.TreeGrid.Rows.Count -eq 0) {
        foreach ($entryGroup in @($treeRows | Group-Object EntryNumber)) {
            $entryNumber = 0
            if (-not [int]::TryParse($entryGroup.Name, [ref]$entryNumber)) { $entryNumber = Get-HooterNextEntryNumber -Grid $script:Ui.TreeGrid }
            $crewRecord = ConvertTo-HooterText ($entryGroup.Group | Select-Object -First 1).CrewRecord
            foreach ($savedRow in $entryGroup.Group) {
                $field = @($script:FieldCatalog.Tree | Where-Object { $_.FieldKey -eq $savedRow.FieldKey } | Select-Object -First 1)
                if ($field.Count -eq 0) { continue }
                Add-HooterCheckRow -Grid $script:Ui.TreeGrid -Field $field[0] -CrewValue $savedRow.CrewValue -EntryNumber $entryNumber -CrewRecord $crewRecord
            }
        }
    }
    if ($script:Ui.RegenGrid.Rows.Count -eq 0) {
        foreach ($entryGroup in @($regenRows | Group-Object EntryNumber)) {
            $entryNumber = 0
            if (-not [int]::TryParse($entryGroup.Name, [ref]$entryNumber)) { $entryNumber = Get-HooterNextEntryNumber -Grid $script:Ui.RegenGrid }
            $crewRecord = ConvertTo-HooterText ($entryGroup.Group | Select-Object -First 1).CrewRecord
            foreach ($savedRow in $entryGroup.Group) {
                $field = @($script:FieldCatalog.Regen | Where-Object { $_.FieldKey -eq $savedRow.FieldKey } | Select-Object -First 1)
                if ($field.Count -eq 0) { continue }
                Add-HooterCheckRow -Grid $script:Ui.RegenGrid -Field $field[0] -CrewValue $savedRow.CrewValue -EntryNumber $entryNumber -CrewRecord $crewRecord
            }
        }
    }
    $missedRows = @($match.Rows | Where-Object { $_.Scope -eq "MissedTree" })
    try {
        $script:SuppressMissedTreeCheckEvent = $true
        if ($script:Ui.ContainsKey("MissedTreeGrid")) {
            $script:Ui.MissedTreeGrid.Rows.Clear()
            foreach ($missedRow in @($missedRows | Where-Object { $_.FieldName -eq "CrewMissedTree" })) {
                Add-HooterMissedTreeRow `
                    -EntryNumber $missedRow.EntryNumber `
                    -DBH $missedRow.MissedTreeDBH `
                    -Distance $missedRow.MissedTreeDistance `
                    -Azimuth $missedRow.MissedTreeAzimuth `
                    -Notes $missedRow.FieldNotes
            }
        }
        if ($script:Ui.ContainsKey("CrewMissedTreeCheck")) {
            $script:Ui.CrewMissedTreeCheck.Checked = ($missedRows.Count -gt 0)
        }
    }
    finally {
        $script:SuppressMissedTreeCheckEvent = $false
    }
    Update-HooterMissedTreeStatus

    foreach ($scopeGrid in @(
        [pscustomobject]@{ Scope = "Plot"; Grid = $script:Ui.PlotGrid },
        [pscustomobject]@{ Scope = "Tree"; Grid = $script:Ui.TreeGrid },
        [pscustomobject]@{ Scope = "Regen"; Grid = $script:Ui.RegenGrid }
    )) {
        foreach ($row in $scopeGrid.Grid.Rows) {
            if ($row.IsNewRow) { continue }
            $entryNumber = if ($scopeGrid.Grid.Columns.Contains("EntryNumber")) { ConvertTo-HooterText $row.Cells["EntryNumber"].Value } else { "1" }
            $crewRecord = if ($scopeGrid.Grid.Columns.Contains("CrewRecord")) { ConvertTo-HooterText $row.Cells["CrewRecord"].Value } else { $script:CurrentPlot.PlotNumber }
            $fieldKey = ConvertTo-HooterText $row.Cells["FieldKey"].Value
            $saved = @($match.Rows | Where-Object {
                $_.Scope -eq $scopeGrid.Scope -and
                $_.FieldKey -eq $fieldKey -and
                ($_.EntryNumber -eq $entryNumber -or (-not [string]::IsNullOrWhiteSpace($crewRecord) -and $_.CrewRecord -eq $crewRecord))
            } | Select-Object -First 1)
            if ($saved.Count -eq 0) { continue }
            if ($scopeGrid.Grid.Columns.Contains("PointValue") -and $null -ne $saved[0].PSObject.Properties["PointValue"]) {
                $savedPointValue = Normalize-HooterQaPointValue $saved[0].PointValue
                if (-not [string]::IsNullOrWhiteSpace($savedPointValue)) {
                    $row.Cells["PointValue"].Value = $savedPointValue
                }
            }
            $row.Cells["QaValue"].Value = $saved[0].QaValue
            $row.Cells["Notes"].Value = $saved[0].FieldNotes
        }
    }
    if ($script:Ui.ContainsKey("ExecutionGrid")) {
        foreach ($row in $script:Ui.ExecutionGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $fieldKey = Get-HooterFieldKey -Group "Execution" -TableName "TableD" -FieldName (ConvertTo-HooterText $row.Cells["ItemKey"].Value)
            $saved = @($match.Rows | Where-Object {
                $_.Scope -eq "Execution" -and ($_.FieldKey -eq $fieldKey -or $_.FieldName -eq (ConvertTo-HooterText $row.Cells["ItemKey"].Value))
            } | Select-Object -First 1)
            if ($saved.Count -eq 0) { continue }
            $row.Cells["Rating"].Value = $saved[0].QaValue
            $row.Cells["Notes"].Value = $saved[0].FieldNotes
            Update-HooterExecutionRowStatus -Row $row
        }
    }
    Set-HooterTreeVisibleRows
    Set-HooterRegenVisibleRows
}

function Get-HooterDefaultExportPath {
    param(
        [string]$Suffix,
        [string]$Extension
    )

    $folder = [Environment]::GetFolderPath("MyDocuments")
    $stamp = Get-Date -Format "yyyyMMdd_HHmm"
    return Join-Path $folder ("PlotHoot_{0}_{1}.{2}" -f $Suffix, $stamp, $Extension.TrimStart("."))
}

function New-HooterExportRow {
    param(
        [string]$RecordType = "QA",
        [string]$SessionID = "",
        [string]$SavedAt = "",
        [string]$Database = "",
        [string]$ProjectName = "",
        [string]$CheckCruiserName = "",
        [string]$CheckCruiseDate = "",
        [object]$InventoryTotalPlots = "",
        [object]$InventoryCheckedPlots = "",
        [object]$InventoryTargetPercent = "",
        [object]$InventoryRequiredPlots = "",
        [object]$InventoryCompletionPercent = "",
        [string]$PlotNumber = "",
        [string]$UTMEasting = "",
        [string]$UTMNorthing = "",
        [string]$UTMZone = "",
        [string]$OverallStatus = "",
        [object]$ScoreEarned = "",
        [object]$ScorePossible = "",
        [object]$ScoreLost = "",
        [object]$ScorePercent = "",
        [object]$ScorePassPercent = "",
        [object]$MaxPointLoss = "",
        [object]$PlotPointTotal = "",
        [object]$TreePointTotal = "",
        [object]$RegenPointTotal = "",
        [object]$ExecutionPointTotal = "",
        [object]$PlotPointLoss = "",
        [object]$TreePointLoss = "",
        [object]$RegenPointLoss = "",
        [object]$ExecutionPointLoss = "",
        [object]$CriticalFailCount = "",
        [object]$FailedCount = "",
        [object]$UncheckedCount = "",
        [object]$MissedTreeCount = "",
        [string]$OverallNotes = "",
        [string]$Scope = "",
        [string]$EntryNumber = "",
        [string]$CrewRecord = "",
        [string]$TableName = "",
        [string]$FieldName = "",
        [string]$FieldLabel = "",
        [string]$FieldKey = "",
        [string]$FieldOrder = "",
        [string]$CrewValue = "",
        [string]$QaValue = "",
        [string]$RuleMode = "",
        [string]$ToleranceValue = "",
        [string]$PointValue = "",
        [string]$CriticalFail = "",
        [string]$Status = "",
        [string]$EarnedPoints = "",
        [string]$PossiblePoints = "",
        [string]$LostPoints = "",
        [string]$CriticalFailure = "",
        [string]$FieldNotes = "",
        [string]$MissedTreeDBH = "",
        [string]$MissedTreeDistance = "",
        [string]$MissedTreeAzimuth = ""
    )

    return [pscustomobject]@{
        RecordType = $RecordType
        SessionID = $SessionID
        SavedAt = $SavedAt
        Database = $Database
        ProjectName = $ProjectName
        CheckCruiserName = $CheckCruiserName
        CheckCruiseDate = $CheckCruiseDate
        InventoryTotalPlots = $InventoryTotalPlots
        InventoryCheckedPlots = $InventoryCheckedPlots
        InventoryTargetPercent = $InventoryTargetPercent
        InventoryRequiredPlots = $InventoryRequiredPlots
        InventoryCompletionPercent = $InventoryCompletionPercent
        PlotNumber = $PlotNumber
        UTMEasting = $UTMEasting
        UTMNorthing = $UTMNorthing
        UTMZone = $UTMZone
        OverallStatus = $OverallStatus
        ScoreEarned = $ScoreEarned
        ScorePossible = $ScorePossible
        ScoreLost = $ScoreLost
        ScorePercent = $ScorePercent
        ScorePassPercent = $ScorePassPercent
        MaxPointLoss = $MaxPointLoss
        PlotPointTotal = $PlotPointTotal
        TreePointTotal = $TreePointTotal
        RegenPointTotal = $RegenPointTotal
        ExecutionPointTotal = $ExecutionPointTotal
        PlotPointLoss = $PlotPointLoss
        TreePointLoss = $TreePointLoss
        RegenPointLoss = $RegenPointLoss
        ExecutionPointLoss = $ExecutionPointLoss
        CriticalFailCount = $CriticalFailCount
        FailedCount = $FailedCount
        UncheckedCount = $UncheckedCount
        MissedTreeCount = $MissedTreeCount
        OverallNotes = $OverallNotes
        Scope = $Scope
        EntryNumber = $EntryNumber
        CrewRecord = $CrewRecord
        TableName = $TableName
        FieldName = $FieldName
        FieldLabel = $FieldLabel
        FieldKey = $FieldKey
        FieldOrder = $FieldOrder
        CrewValue = $CrewValue
        QaValue = $QaValue
        RuleMode = $RuleMode
        ToleranceValue = $ToleranceValue
        PointValue = $PointValue
        CriticalFail = $CriticalFail
        Status = $Status
        EarnedPoints = $EarnedPoints
        PossiblePoints = $PossiblePoints
        LostPoints = $LostPoints
        CriticalFailure = $CriticalFailure
        FieldNotes = $FieldNotes
        MissedTreeDBH = $MissedTreeDBH
        MissedTreeDistance = $MissedTreeDistance
        MissedTreeAzimuth = $MissedTreeAzimuth
    }
}

function Get-HooterSetupExportRows {
    if ($script:Ui.ContainsKey("SettingsGrid")) {
        Update-HooterSettingsFromGrid
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $fields = @(Get-HooterAllScoredFields)
    $inventoryProgress = Get-HooterInventoryQaProgress
    foreach ($field in $fields) {
        $fieldKey = ConvertTo-HooterText $field.FieldKey
        if ([string]::IsNullOrWhiteSpace($fieldKey) -or $seen.ContainsKey($fieldKey)) { continue }
        $seen[$fieldKey] = $true
        $tolerance = Get-HooterTolerance -Field $field
        $fieldOrder = if ($script:FieldOrder.ContainsKey($fieldKey)) { ConvertTo-HooterText $script:FieldOrder[$fieldKey] } else { "" }
        $toleranceValue = ConvertTo-HooterText $tolerance.Value
        $pointValue = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $tolerance)
        if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) {
            $executionRule = Get-HooterExecutionScoreRule -ItemKey $field.FieldName
            $toleranceValue = ConvertTo-HooterText $executionRule.ToleranceValue
            $pointValue = Format-HooterScoreNumber $executionRule.PossiblePoints
        }
        [void]$rows.Add((New-HooterExportRow `
            -RecordType "Setting" `
            -SessionID "SETUP" `
            -SavedAt (Get-Date).ToString("s") `
            -Database $script:DatabasePath `
            -ProjectName (Get-HooterCurrentProjectName) `
            -InventoryTotalPlots $inventoryProgress.TotalPlots `
            -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
            -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
            -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
            -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
            -Scope $field.Group `
            -TableName $field.TableName `
            -FieldName $field.FieldName `
            -FieldLabel $field.Label `
            -FieldKey $fieldKey `
            -FieldOrder $fieldOrder `
            -RuleMode $tolerance.Mode `
            -ToleranceValue $toleranceValue `
            -PointValue $pointValue `
            -CriticalFail (Get-HooterCriticalFail -Tolerance $tolerance) `
            -ScorePassPercent (Normalize-HooterScorePassPercent $script:ScorePassPercent) `
            -MaxPointLoss (Normalize-HooterMaxPointLoss $script:MaxPointLoss) `
            -PlotPointTotal (Normalize-HooterPointTotal $script:PlotPointTotal) `
            -TreePointTotal (Normalize-HooterPointTotal $script:TreePointTotal) `
            -RegenPointTotal (Normalize-HooterPointTotal $script:RegenPointTotal) `
            -ExecutionPointTotal (Normalize-HooterPointTotal $script:ExecutionPointTotal) `
            -Status "Setting"))
    }

    foreach ($fieldKey in @($script:Tolerances.Keys | Sort-Object)) {
        if ($seen.ContainsKey($fieldKey)) { continue }
        if (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey) { continue }
        if (Test-HooterHiddenFieldKey -FieldKey $fieldKey) { continue }
        $tolerance = $script:Tolerances[$fieldKey]
        $fieldOrder = if ($script:FieldOrder.ContainsKey($fieldKey)) { ConvertTo-HooterText $script:FieldOrder[$fieldKey] } else { "" }
        $toleranceValue = ConvertTo-HooterText $tolerance.Value
        $pointValue = Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $tolerance)
        if (Test-HooterExecutionFieldKey -FieldKey $fieldKey) {
            $executionRule = Get-HooterExecutionScoreRule -ItemKey $tolerance.FieldName
            $toleranceValue = ConvertTo-HooterText $executionRule.ToleranceValue
            $pointValue = Format-HooterScoreNumber $executionRule.PossiblePoints
        }
        [void]$rows.Add((New-HooterExportRow `
            -RecordType "Setting" `
            -SessionID "SETUP" `
            -SavedAt (Get-Date).ToString("s") `
            -Database $script:DatabasePath `
            -ProjectName (Get-HooterCurrentProjectName) `
            -InventoryTotalPlots $inventoryProgress.TotalPlots `
            -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
            -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
            -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
            -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
            -Scope $tolerance.Group `
            -TableName $tolerance.TableName `
            -FieldName $tolerance.FieldName `
            -FieldLabel $tolerance.Label `
            -FieldKey $fieldKey `
            -FieldOrder $fieldOrder `
            -RuleMode $tolerance.Mode `
            -ToleranceValue $toleranceValue `
            -PointValue $pointValue `
            -CriticalFail (Get-HooterCriticalFail -Tolerance $tolerance) `
            -ScorePassPercent (Normalize-HooterScorePassPercent $script:ScorePassPercent) `
            -MaxPointLoss (Normalize-HooterMaxPointLoss $script:MaxPointLoss) `
            -PlotPointTotal (Normalize-HooterPointTotal $script:PlotPointTotal) `
            -TreePointTotal (Normalize-HooterPointTotal $script:TreePointTotal) `
            -RegenPointTotal (Normalize-HooterPointTotal $script:RegenPointTotal) `
            -ExecutionPointTotal (Normalize-HooterPointTotal $script:ExecutionPointTotal) `
            -Status "Setting"))
    }

    [void]$rows.Add((New-HooterExportRow `
        -RecordType "Setting" `
        -SessionID "SETUP" `
        -SavedAt (Get-Date).ToString("s") `
        -Database $script:DatabasePath `
        -ProjectName (Get-HooterCurrentProjectName) `
        -InventoryTotalPlots $inventoryProgress.TotalPlots `
        -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
        -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
        -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
        -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
        -Scope "Scoring" `
        -TableName "PlotHoot QA" `
        -FieldName "UseStemCountPercentageForScoring" `
        -FieldLabel "Use Stem Count Percentage For Scoring" `
        -RuleMode "StemPercentToggle" `
        -ToleranceValue ([string][bool]$script:UseStemCountPercentageForScoring) `
        -ScorePassPercent (Normalize-HooterScorePassPercent $script:ScorePassPercent) `
        -MaxPointLoss (Normalize-HooterMaxPointLoss $script:MaxPointLoss) `
        -PlotPointTotal (Normalize-HooterPointTotal $script:PlotPointTotal) `
        -TreePointTotal (Normalize-HooterPointTotal $script:TreePointTotal) `
        -RegenPointTotal (Normalize-HooterPointTotal $script:RegenPointTotal) `
        -ExecutionPointTotal (Normalize-HooterPointTotal $script:ExecutionPointTotal) `
        -Status "Setting"))

    foreach ($fieldKey in @($script:ExcludedFieldKeys.Keys | Where-Object { -not (Test-HooterHiddenFieldKey -FieldKey $_) } | Sort-Object)) {
        $parts = $fieldKey -split "\|", 3
        $scope = if ($parts.Count -ge 1) { $parts[0] } else { "" }
        $tableName = if ($parts.Count -ge 2) { $parts[1] } else { "" }
        $fieldName = if ($parts.Count -ge 3) { $parts[2] } else { "" }
        [void]$rows.Add((New-HooterExportRow `
            -RecordType "Setting" `
            -SessionID "SETUP" `
            -SavedAt (Get-Date).ToString("s") `
            -Database $script:DatabasePath `
            -ProjectName (Get-HooterCurrentProjectName) `
            -InventoryTotalPlots $inventoryProgress.TotalPlots `
            -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
            -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
            -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
            -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
            -Scope "FieldUse" `
            -TableName $tableName `
            -FieldName $fieldName `
            -FieldLabel "Excluded field" `
            -FieldKey $fieldKey `
            -RuleMode "ExcludedField" `
            -ToleranceValue "False" `
            -Status "Setting"))
    }

    foreach ($band in @(Get-HooterStemToleranceBands)) {
        [void]$rows.Add((New-HooterExportRow `
            -RecordType "Setting" `
            -SessionID "SETUP" `
            -SavedAt (Get-Date).ToString("s") `
            -Database $script:DatabasePath `
            -ProjectName (Get-HooterCurrentProjectName) `
            -InventoryTotalPlots $inventoryProgress.TotalPlots `
            -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
            -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
            -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
            -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
            -Scope "StemTolerance" `
            -TableName "PlotHoot QA" `
            -FieldName (ConvertTo-HooterText $band.Key) `
            -FieldLabel (ConvertTo-HooterText $band.Label) `
            -RuleMode "StemCountBand" `
            -ToleranceValue (Format-HooterStemToleranceValue $band.Tolerance) `
            -ScorePassPercent (Normalize-HooterScorePassPercent $script:ScorePassPercent) `
            -MaxPointLoss (Normalize-HooterMaxPointLoss $script:MaxPointLoss) `
            -PlotPointTotal (Normalize-HooterPointTotal $script:PlotPointTotal) `
            -TreePointTotal (Normalize-HooterPointTotal $script:TreePointTotal) `
            -RegenPointTotal (Normalize-HooterPointTotal $script:RegenPointTotal) `
            -ExecutionPointTotal (Normalize-HooterPointTotal $script:ExecutionPointTotal) `
            -Status "Setting"))
    }

    return @($rows.ToArray())
}

function Get-HooterFlatSessionRows {
    $flatRows = New-Object System.Collections.Generic.List[object]
    foreach ($setupRow in (Get-HooterSetupExportRows)) {
        [void]$flatRows.Add($setupRow)
    }
    $inventoryProgress = Get-HooterInventoryQaProgress
    foreach ($session in @($script:QaSessions.ToArray())) {
        foreach ($row in @($session.Rows)) {
            $missedTreeDbh = if ($null -ne $row.PSObject.Properties["MissedTreeDBH"]) { ConvertTo-HooterText $row.MissedTreeDBH } else { "" }
            $missedTreeDistance = if ($null -ne $row.PSObject.Properties["MissedTreeDistance"]) { ConvertTo-HooterText $row.MissedTreeDistance } else { "" }
            $missedTreeAzimuth = if ($null -ne $row.PSObject.Properties["MissedTreeAzimuth"]) { ConvertTo-HooterText $row.MissedTreeAzimuth } else { "" }
            [void]$flatRows.Add((New-HooterExportRow `
                -RecordType "QA" `
                -SessionID $session.SessionID `
                -SavedAt $session.SavedAt `
                -Database $session.Database `
                -ProjectName $session.ProjectName `
                -CheckCruiserName $(if ($null -ne $session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $session.CheckCruiserName } else { "" }) `
                -CheckCruiseDate $(if ($null -ne $session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $session.CheckCruiseDate } else { "" }) `
                -InventoryTotalPlots $inventoryProgress.TotalPlots `
                -InventoryCheckedPlots $inventoryProgress.CheckedPlots `
                -InventoryTargetPercent (Format-HooterScoreNumber $inventoryProgress.TargetPercent) `
                -InventoryRequiredPlots $inventoryProgress.RequiredPlots `
                -InventoryCompletionPercent (Format-HooterScoreNumber $inventoryProgress.CompletionPercent) `
                -PlotNumber $session.PlotNumber `
                -UTMEasting $session.UTMEasting `
                -UTMNorthing $session.UTMNorthing `
                -UTMZone $session.UTMZone `
                -OverallStatus $session.OverallStatus `
                -ScoreEarned $session.ScoreEarned `
                -ScorePossible $session.ScorePossible `
                -ScoreLost $session.ScoreLost `
                -ScorePercent $session.ScorePercent `
                -ScorePassPercent $session.ScorePassPercent `
                -MaxPointLoss $session.MaxPointLoss `
                -PlotPointTotal $session.PlotPointTotal `
                -TreePointTotal $session.TreePointTotal `
                -RegenPointTotal $session.RegenPointTotal `
                -ExecutionPointTotal $session.ExecutionPointTotal `
                -PlotPointLoss $session.PlotPointLoss `
                -TreePointLoss $session.TreePointLoss `
                -RegenPointLoss $session.RegenPointLoss `
                -ExecutionPointLoss $session.ExecutionPointLoss `
                -CriticalFailCount $session.CriticalFailCount `
                -FailedCount $session.FailedCount `
                -UncheckedCount $session.UncheckedCount `
                -MissedTreeCount $session.MissedTreeCount `
                -OverallNotes $session.OverallNotes `
                -Scope $row.Scope `
                -EntryNumber $row.EntryNumber `
                -CrewRecord $row.CrewRecord `
                -TableName $row.TableName `
                -FieldName $row.FieldName `
                -FieldLabel $row.FieldLabel `
                -FieldKey $row.FieldKey `
                -FieldOrder $(if ($null -ne $row.PSObject.Properties["FieldOrder"]) { ConvertTo-HooterText $row.FieldOrder } else { "" }) `
                -CrewValue $row.CrewValue `
                -QaValue $row.QaValue `
                -RuleMode $row.RuleMode `
                -ToleranceValue $row.ToleranceValue `
                -PointValue $row.PointValue `
                -CriticalFail $row.CriticalFail `
                -Status $row.Status `
                -EarnedPoints $row.EarnedPoints `
                -PossiblePoints $row.PossiblePoints `
                -LostPoints $row.LostPoints `
                -CriticalFailure $row.CriticalFailure `
                -FieldNotes $row.FieldNotes `
                -MissedTreeDBH $missedTreeDbh `
                -MissedTreeDistance $missedTreeDistance `
                -MissedTreeAzimuth $missedTreeAzimuth))
        }
    }
    return @($flatRows.ToArray())
}

function Export-HooterQaCsv {
    try {
        if ($null -ne $script:CurrentPlot) { Save-HooterCurrentSession }
        if ($script:QaSessions.Count -eq 0) { throw "There are no saved plot QA checks to export." }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $dialog.FileName = [System.IO.Path]::GetFileName((Get-HooterDefaultExportPath -Suffix "QA_All" -Extension "csv"))
        $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Get-HooterFlatSessionRows | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
        Set-HooterStatus "Exported QA CSV: $($dialog.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export QA CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Get-HooterImportRecordType {
    param([object]$Row)

    $recordType = ConvertTo-HooterText $Row.RecordType
    if ([string]::IsNullOrWhiteSpace($recordType)) { return "QA" }
    return $recordType
}

function Import-HooterFieldOrderRows {
    param([object[]]$Rows)

    $applied = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row.PSObject.Properties["FieldOrder"]) { continue }
        $fieldOrderText = ConvertTo-HooterText $row.FieldOrder
        if ([string]::IsNullOrWhiteSpace($fieldOrderText)) { continue }
        $fieldKey = ConvertTo-HooterText $row.FieldKey
        if ([string]::IsNullOrWhiteSpace($fieldKey) -or (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey)) { continue }
        $fieldOrder = 0
        if (-not [int]::TryParse($fieldOrderText, [ref]$fieldOrder)) { continue }
        $script:FieldOrder[$fieldKey] = $fieldOrder
        $applied[$fieldKey] = $true
    }
    return $applied.Count
}

function Import-HooterSetupRows {
    param([object[]]$Rows)

    $count = 0
    $obsoleteExecutionSettingFound = $false
    foreach ($row in @($Rows)) {
        if ($null -ne $row.PSObject.Properties["ProjectName"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.ProjectName))) {
            $script:ProjectName = Normalize-HooterProjectName -Value $row.ProjectName
            Update-HooterProjectDisplay
        }
        if ($null -ne $row.PSObject.Properties["ScorePassPercent"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.ScorePassPercent))) {
            $script:ScorePassPercent = Normalize-HooterScorePassPercent $row.ScorePassPercent
        }
        if ($null -ne $row.PSObject.Properties["MaxPointLoss"]) {
            $script:MaxPointLoss = Normalize-HooterMaxPointLoss $row.MaxPointLoss
        }
        if ($null -ne $row.PSObject.Properties["PlotPointTotal"]) {
            $script:PlotPointTotal = Normalize-HooterPointTotal $row.PlotPointTotal
        }
        if ($null -ne $row.PSObject.Properties["TreePointTotal"]) {
            $script:TreePointTotal = Normalize-HooterPointTotal $row.TreePointTotal
        }
        if ($null -ne $row.PSObject.Properties["RegenPointTotal"]) {
            $script:RegenPointTotal = Normalize-HooterPointTotal $row.RegenPointTotal
        }
        if ($null -ne $row.PSObject.Properties["ExecutionPointTotal"]) {
            $script:ExecutionPointTotal = Normalize-HooterPointTotal $row.ExecutionPointTotal
        }
        $scopeForSpecialSetting = ConvertTo-HooterText $row.Scope
        $ruleModeForSpecialSetting = ConvertTo-HooterText $row.RuleMode
        if ($scopeForSpecialSetting.Equals("Scoring", [System.StringComparison]::OrdinalIgnoreCase) -and
            $ruleModeForSpecialSetting.Equals("StemPercentToggle", [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:UseStemCountPercentageForScoring = ConvertTo-HooterBool $row.ToleranceValue
            $count++
            continue
        }
        if ($scopeForSpecialSetting.Equals("FieldUse", [System.StringComparison]::OrdinalIgnoreCase) -or
            $ruleModeForSpecialSetting.Equals("ExcludedField", [System.StringComparison]::OrdinalIgnoreCase)) {
            $fieldKey = ConvertTo-HooterText $row.FieldKey
            if (-not [string]::IsNullOrWhiteSpace($fieldKey) -and -not (Test-HooterHiddenFieldKey -FieldKey $fieldKey)) {
                if ((ConvertTo-HooterText $row.ToleranceValue).Equals("False", [System.StringComparison]::OrdinalIgnoreCase) -or
                    (ConvertTo-HooterText $row.Status).Equals("Disabled", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $script:ExcludedFieldKeys[$fieldKey] = $true
                }
                else {
                    if ($script:ExcludedFieldKeys.ContainsKey($fieldKey)) { $script:ExcludedFieldKeys.Remove($fieldKey) }
                }
                $count++
            }
            continue
        }
        if ($scopeForSpecialSetting.Equals("StemTolerance", [System.StringComparison]::OrdinalIgnoreCase) -or
            $ruleModeForSpecialSetting.Equals("StemCountBand", [System.StringComparison]::OrdinalIgnoreCase)) {
            $stemKey = ConvertTo-HooterText $row.FieldName
            if ([string]::IsNullOrWhiteSpace($stemKey)) {
                $stemKey = ConvertTo-HooterText $row.FieldLabel
            }
            if (Set-HooterStemToleranceBand -Key $stemKey -Tolerance $row.ToleranceValue) {
                $count++
            }
            continue
        }
        $fieldKey = ConvertTo-HooterText $row.FieldKey
        if ([string]::IsNullOrWhiteSpace($fieldKey)) { continue }
        if (Test-HooterHiddenFieldKey -FieldKey $fieldKey) { continue }
        if (Test-HooterObsoleteExecutionFieldKey -FieldKey $fieldKey) {
            $obsoleteExecutionSettingFound = $true
            continue
        }
        [void](Import-HooterFieldOrderRows -Rows @($row))
        $parts = $fieldKey -split "\|", 3
        $scope = ConvertTo-HooterText $row.Scope
        $tableName = ConvertTo-HooterText $row.TableName
        $fieldName = ConvertTo-HooterText $row.FieldName
        if ([string]::IsNullOrWhiteSpace($scope) -and $parts.Count -ge 1) { $scope = $parts[0] }
        if ([string]::IsNullOrWhiteSpace($tableName) -and $parts.Count -ge 2) { $tableName = $parts[1] }
        if ([string]::IsNullOrWhiteSpace($fieldName) -and $parts.Count -ge 3) { $fieldName = $parts[2] }
        $mode = ConvertTo-HooterText $row.RuleMode
        if ([string]::IsNullOrWhiteSpace($mode)) { $mode = ConvertTo-HooterText $row.Mode }
        $mode = Normalize-HooterToleranceMode $mode
        $default = Get-HooterTolerance -Field $fieldKey
        $script:Tolerances[$fieldKey] = [pscustomobject]@{
            FieldKey = $fieldKey
            Group = $scope
            TableName = $tableName
            FieldName = $fieldName
            Label = ConvertTo-HooterText $row.FieldLabel
            Mode = $mode
            Value = ConvertTo-HooterText $row.ToleranceValue
            PointValue = if ($null -ne $row.PSObject.Properties["PointValue"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.PointValue))) { ConvertTo-HooterText $row.PointValue } else { ConvertTo-HooterText $default.PointValue }
            CriticalFail = if ($null -ne $row.PSObject.Properties["CriticalFail"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.CriticalFail))) { ConvertTo-HooterBool $row.CriticalFail } else { Get-HooterCriticalFail -Tolerance $default }
        }
        $count++
    }

    $executionTotalNumber = 0.0
    if ((ConvertTo-HooterNumber -Value $script:ExecutionPointTotal -Number ([ref]$executionTotalNumber)) -and
        [Math]::Abs($executionTotalNumber - 12.0) -lt 0.000001) {
        $script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
    }

    if ($count -gt 0) {
        Apply-HooterSavedFieldOrderToCatalog
        Save-HooterSettings
        if ($script:Ui.ContainsKey("SettingsGrid")) { Populate-HooterSettingsGrid }
        if ($script:Ui.ContainsKey("StemToleranceGrid")) { Populate-HooterStemToleranceGrid }
        Update-HooterStemScoringControls
        if ($script:Ui.ContainsKey("PlotGrid") -or $script:Ui.ContainsKey("TreeGrid") -or $script:Ui.ContainsKey("RegenGrid")) {
            Apply-HooterFieldOrderToLoadedPlot
            Refresh-HooterAllStatuses
        }
    }
    return $count
}

function Get-HooterSavedRowsScoreStats {
    param([object[]]$Rows)

    $total = 0
    $checked = 0
    $failed = 0
    $earnedPoints = 0.0
    $possiblePoints = 0.0
    $lostPoints = 0.0
    $criticalFailures = 0

    foreach ($row in @($Rows)) {
        $total++
        $status = ConvertTo-HooterText $row.Status
        $qaValue = ConvertTo-HooterText $row.QaValue
        if (-not [string]::IsNullOrWhiteSpace($qaValue)) { $checked++ }
        if ($status -eq "Fail") { $failed++ }

        $rowPossible = 0.0
        $rowEarned = 0.0
        $rowLost = 0.0
        $hasPossible = ($null -ne $row.PSObject.Properties["PossiblePoints"] -and (ConvertTo-HooterNumber -Value $row.PossiblePoints -Number ([ref]$rowPossible)))
        if (-not $hasPossible) {
            if ($null -ne $row.PSObject.Properties["PointValue"]) {
                [void](ConvertTo-HooterNumber -Value $row.PointValue -Number ([ref]$rowPossible))
            }
            if ($rowPossible -le 0) { $rowPossible = 1.0 }
        }
        if ($null -ne $row.PSObject.Properties["EarnedPoints"] -and (ConvertTo-HooterNumber -Value $row.EarnedPoints -Number ([ref]$rowEarned))) {
        }
        elseif ($status -eq "Pass") {
            $rowEarned = $rowPossible
        }
        if ($null -ne $row.PSObject.Properties["LostPoints"] -and (ConvertTo-HooterNumber -Value $row.LostPoints -Number ([ref]$rowLost))) {
        }
        elseif ($status -eq "Fail") {
            $rowLost = $rowPossible
        }

        $criticalSetting = $false
        if ($null -ne $row.PSObject.Properties["CriticalFail"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.CriticalFail))) {
            $criticalSetting = ConvertTo-HooterBool $row.CriticalFail
        }
        else {
            $criticalSetting = ((ConvertTo-HooterText $row.RuleMode) -eq "PassFail")
        }

        $criticalFailure = $false
        if ($null -ne $row.PSObject.Properties["CriticalFailure"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $row.CriticalFailure))) {
            $criticalFailure = ConvertTo-HooterBool $row.CriticalFailure
        }
        else {
            $criticalFailure = ($status -eq "Fail" -and $criticalSetting)
        }
        if ($criticalFailure) { $criticalFailures++ }

        $possiblePoints += $rowPossible
        $earnedPoints += $rowEarned
        $lostPoints += $rowLost
    }

    return Get-HooterOverallStatusFromStats -Total $total -Checked $checked -Failed $failed -EarnedPoints $earnedPoints -PossiblePoints $possiblePoints -LostPoints $lostPoints -CriticalFailures $criticalFailures
}

function Import-HooterQaCsv {
    try {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $rows = @(Import-Csv -LiteralPath $dialog.FileName)
        if ($rows.Count -eq 0) { throw "The selected CSV did not contain any QA rows." }
        $setupRows = @($rows | Where-Object { (Get-HooterImportRecordType -Row $_) -in @("Setting", "Settings", "Setup", "Tolerance") })
        $qaRows = @($rows | Where-Object { (Get-HooterImportRecordType -Row $_) -notin @("Setting", "Settings", "Setup", "Tolerance") })
        $setupCount = Import-HooterSetupRows -Rows $setupRows
        $qaFieldOrderCount = Import-HooterFieldOrderRows -Rows $qaRows
        if ($qaFieldOrderCount -gt 0) {
            Apply-HooterSavedFieldOrderToCatalog
            Save-HooterSettings
            if ($script:Ui.ContainsKey("SettingsGrid")) { Populate-HooterSettingsGrid }
            if ($null -ne $script:CurrentPlot) { Apply-HooterFieldOrderToLoadedPlot }
        }
        if ($qaRows.Count -eq 0) {
            Update-HooterReviewGrid
            Set-HooterStatus "Imported $setupCount setup setting(s) and $qaFieldOrderCount field order value(s); no QA rows were found in the file."
            return
        }

        $setupMessage = if (($setupCount + $qaFieldOrderCount) -gt 0) { " Imported setup settings and field order have already been applied." } else { "" }
        $answer = [System.Windows.Forms.MessageBox]::Show("Replace the current saved QA list with the imported file?$setupMessage", "Import QA CSV", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { $script:QaSessions.Clear() }

        foreach ($group in ($qaRows | Group-Object { if (-not [string]::IsNullOrWhiteSpace($_.SessionID)) { $_.SessionID } else { $_.PlotNumber } })) {
            $first = $group.Group | Select-Object -First 1
            $sessionRows = @($group.Group | ForEach-Object {
                [pscustomobject]@{
                    Scope = ConvertTo-HooterText $_.Scope
                    EntryNumber = ConvertTo-HooterText $_.EntryNumber
                    CrewRecord = ConvertTo-HooterText $_.CrewRecord
                    TableName = ConvertTo-HooterText $_.TableName
                    FieldName = ConvertTo-HooterText $_.FieldName
                    FieldLabel = ConvertTo-HooterText $_.FieldLabel
                    FieldKey = ConvertTo-HooterText $_.FieldKey
                    FieldOrder = if ($null -ne $_.PSObject.Properties["FieldOrder"]) { ConvertTo-HooterText $_.FieldOrder } else { "" }
                    CrewValue = ConvertTo-HooterText $_.CrewValue
                    QaValue = ConvertTo-HooterText $_.QaValue
                    RuleMode = ConvertTo-HooterText $_.RuleMode
                    ToleranceValue = ConvertTo-HooterText $_.ToleranceValue
                    PointValue = ConvertTo-HooterText $_.PointValue
                    CriticalFail = ConvertTo-HooterText $_.CriticalFail
                    Status = ConvertTo-HooterText $_.Status
                    EarnedPoints = ConvertTo-HooterText $_.EarnedPoints
                    PossiblePoints = ConvertTo-HooterText $_.PossiblePoints
                    LostPoints = ConvertTo-HooterText $_.LostPoints
                    CriticalFailure = ConvertTo-HooterText $_.CriticalFailure
                    FieldNotes = ConvertTo-HooterText $_.FieldNotes
                    MissedTreeDBH = if ($null -ne $_.PSObject.Properties["MissedTreeDBH"]) { ConvertTo-HooterText $_.MissedTreeDBH } else { "" }
                    MissedTreeDistance = if ($null -ne $_.PSObject.Properties["MissedTreeDistance"]) { ConvertTo-HooterText $_.MissedTreeDistance } else { "" }
                    MissedTreeAzimuth = if ($null -ne $_.PSObject.Properties["MissedTreeAzimuth"]) { ConvertTo-HooterText $_.MissedTreeAzimuth } else { "" }
                }
            })
            $scoreStats = Get-HooterSavedRowsScoreStats -Rows $sessionRows
            $overall = ConvertTo-HooterText $first.OverallStatus
            if ([string]::IsNullOrWhiteSpace($overall)) { $overall = $scoreStats.Status }
            [void]$script:QaSessions.Add([pscustomobject]@{
                SessionID = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.SessionID))) { ([guid]::NewGuid()).ToString() } else { ConvertTo-HooterText $first.SessionID }
                SavedAt = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.SavedAt))) { (Get-Date).ToString("s") } else { ConvertTo-HooterText $first.SavedAt }
                Database = ConvertTo-HooterText $first.Database
                ProjectName = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.ProjectName))) { Get-HooterCurrentProjectName } else { Normalize-HooterProjectName -Value $first.ProjectName }
                CheckCruiserName = if ($null -ne $first.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $first.CheckCruiserName } else { "" }
                CheckCruiseDate = if ($null -ne $first.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $first.CheckCruiseDate } else { "" }
                PlotNumber = ConvertTo-HooterText $first.PlotNumber
                UTMEasting = ConvertTo-HooterText $first.UTMEasting
                UTMNorthing = ConvertTo-HooterText $first.UTMNorthing
                UTMZone = ConvertTo-HooterText $first.UTMZone
                OverallStatus = $overall
                CheckedCount = $scoreStats.Checked
                FailedCount = $scoreStats.Failed
                UncheckedCount = if ($null -ne $first.PSObject.Properties["UncheckedCount"] -and -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.UncheckedCount))) { ConvertTo-HooterText $first.UncheckedCount } else { [Math]::Max(0, $scoreStats.Total - $scoreStats.Checked) }
                MissedTreeCount = @($sessionRows | Where-Object { $_.Scope -eq "MissedTree" -and $_.FieldName -eq "CrewMissedTree" }).Count
                ScoreEarned = Format-HooterScoreNumber $scoreStats.EarnedPoints
                ScorePossible = Format-HooterScoreNumber $scoreStats.PossiblePoints
                ScoreLost = Format-HooterScoreNumber $scoreStats.LostPoints
                ScorePercent = Format-HooterScoreNumber $scoreStats.ScorePercent
                ScorePassPercent = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.ScorePassPercent))) { Format-HooterScoreNumber $scoreStats.ScorePassPercent } else { ConvertTo-HooterText $first.ScorePassPercent }
                MaxPointLoss = if ($null -ne $first.PSObject.Properties["MaxPointLoss"]) { ConvertTo-HooterText $first.MaxPointLoss } else { ConvertTo-HooterText $scoreStats.MaxPointLoss }
                PlotPointTotal = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.PlotPointTotal))) { Normalize-HooterPointTotal $script:PlotPointTotal } else { ConvertTo-HooterText $first.PlotPointTotal }
                TreePointTotal = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.TreePointTotal))) { Normalize-HooterPointTotal $script:TreePointTotal } else { ConvertTo-HooterText $first.TreePointTotal }
                RegenPointTotal = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.RegenPointTotal))) { Normalize-HooterPointTotal $script:RegenPointTotal } else { ConvertTo-HooterText $first.RegenPointTotal }
                ExecutionPointTotal = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.ExecutionPointTotal))) { Normalize-HooterPointTotal $script:ExecutionPointTotal } else { ConvertTo-HooterText $first.ExecutionPointTotal }
                PlotPointLoss = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.PlotPointLoss))) { "" } else { ConvertTo-HooterText $first.PlotPointLoss }
                TreePointLoss = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.TreePointLoss))) { "" } else { ConvertTo-HooterText $first.TreePointLoss }
                RegenPointLoss = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.RegenPointLoss))) { "" } else { ConvertTo-HooterText $first.RegenPointLoss }
                ExecutionPointLoss = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $first.ExecutionPointLoss))) { "" } else { ConvertTo-HooterText $first.ExecutionPointLoss }
                CriticalFailCount = $scoreStats.CriticalFailures
                OverallNotes = ConvertTo-HooterText $first.OverallNotes
                Rows = $sessionRows
            })
        }
        $importedProject = ConvertTo-HooterText (@($script:QaSessions.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $_.ProjectName)) } | Select-Object -First 1).ProjectName)
        if (-not [string]::IsNullOrWhiteSpace($importedProject)) {
            $script:ProjectName = Normalize-HooterProjectName -Value $importedProject
            Update-HooterProjectDisplay
        }
        Update-HooterReviewGrid
        Set-HooterStatus "Imported $($script:QaSessions.Count) saved plot QA check(s), $setupCount setup setting(s), and $qaFieldOrderCount field order value(s)."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Import QA CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Get-HooterFailedSummaries {
    $summaries = New-Object System.Collections.Generic.List[object]
    $inventoryProgress = Get-HooterInventoryQaProgress
    foreach ($session in @($script:QaSessions.ToArray())) {
        $failedRows = @($session.Rows | Where-Object { $_.Status -eq "Fail" })
        if ($session.OverallStatus -ne "Fail") { continue }
        $failedText = [string]::Join("; ", [string[]]@($failedRows | ForEach-Object {
            $valuePart = "crew='" + (ConvertTo-HooterText $_.CrewValue) + "', QA='" + (ConvertTo-HooterText $_.QaValue) + "'"
            "$($_.Scope) $($_.EntryNumber) $($_.FieldLabel) ($valuePart)"
        }))
        $checkedCount = 0
        [void][int]::TryParse((ConvertTo-HooterText $session.CheckedCount), [ref]$checkedCount)
        $uncheckedCount = if ($null -ne $session.PSObject.Properties["UncheckedCount"]) { ConvertTo-HooterText $session.UncheckedCount } else { [Math]::Max(0, @($session.Rows).Count - $checkedCount) }
        [void]$summaries.Add([pscustomobject]@{
            ProjectName = if ([string]::IsNullOrWhiteSpace((ConvertTo-HooterText $session.ProjectName))) { Get-HooterCurrentProjectName } else { Normalize-HooterProjectName -Value $session.ProjectName }
            InventoryTotalPlots = $inventoryProgress.TotalPlots
            InventoryCheckedPlots = $inventoryProgress.CheckedPlots
            InventoryTargetPercent = Format-HooterScoreNumber $inventoryProgress.TargetPercent
            InventoryRequiredPlots = $inventoryProgress.RequiredPlots
            InventoryCompletionPercent = Format-HooterScoreNumber $inventoryProgress.CompletionPercent
            PlotNumber = $session.PlotNumber
            CheckCruiserName = if ($null -ne $session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $session.CheckCruiserName } else { "" }
            CheckCruiseDate = if ($null -ne $session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $session.CheckCruiseDate } else { "" }
            OverallStatus = $session.OverallStatus
            Score = "total error $(Format-HooterScoreNumber $session.ScoreLost); $(Get-HooterMaxPointLossSummaryText -Value $session.MaxPointLoss); max possible $(Format-HooterScoreNumber $session.ScorePossible)"
            ScorePercent = Format-HooterScoreNumber $session.ScorePercent
            ScorePassPercent = Format-HooterScoreNumber $session.ScorePassPercent
            ScoreLost = Format-HooterScoreNumber $session.ScoreLost
            MaxPointLoss = Get-HooterMaxPointLossDisplay -Value $session.MaxPointLoss
            PlotPointTotal = Format-HooterScoreNumber $session.PlotPointTotal
            TreePointTotal = Format-HooterScoreNumber $session.TreePointTotal
            RegenPointTotal = Format-HooterScoreNumber $session.RegenPointTotal
            ExecutionPointTotal = Format-HooterScoreNumber $session.ExecutionPointTotal
            PlotPointLoss = Format-HooterScoreNumber $session.PlotPointLoss
            TreePointLoss = Format-HooterScoreNumber $session.TreePointLoss
            RegenPointLoss = Format-HooterScoreNumber $session.RegenPointLoss
            ExecutionPointLoss = Format-HooterScoreNumber $session.ExecutionPointLoss
            CriticalFailCount = $session.CriticalFailCount
            FailedCount = $failedRows.Count
            UncheckedCount = $uncheckedCount
            MissedTreeCount = if ($null -ne $session.PSObject.Properties["MissedTreeCount"]) { $session.MissedTreeCount } else { @($session.Rows | Where-Object { $_.Scope -eq "MissedTree" -and $_.FieldName -eq "CrewMissedTree" }).Count }
            UTMEasting = $session.UTMEasting
            UTMNorthing = $session.UTMNorthing
            UTMZone = $session.UTMZone
            QA_Remarks = $session.OverallNotes
            FailedFields = $failedText
            SavedAt = $session.SavedAt
        })
    }
    return @($summaries.ToArray())
}

function Get-HooterSessionFailedCount {
    param([object]$Session)

    if ($null -eq $Session) { return 0 }
    $failedRows = @($Session.Rows | Where-Object { (ConvertTo-HooterText $_.Status) -eq "Fail" })
    if ($failedRows.Count -gt 0) { return $failedRows.Count }
    $failedCount = 0
    if ($null -ne $Session.PSObject.Properties["FailedCount"] -and [int]::TryParse((ConvertTo-HooterText $Session.FailedCount), [ref]$failedCount)) {
        return $failedCount
    }
    return 0
}

function Get-HooterSessionUncheckedCount {
    param([object]$Session)

    if ($null -eq $Session) { return 0 }
    $uncheckedCount = 0
    if ($null -ne $Session.PSObject.Properties["UncheckedCount"] -and [int]::TryParse((ConvertTo-HooterText $Session.UncheckedCount), [ref]$uncheckedCount)) {
        return [Math]::Max(0, $uncheckedCount)
    }
    $checkedCount = 0
    [void][int]::TryParse((ConvertTo-HooterText $Session.CheckedCount), [ref]$checkedCount)
    return [Math]::Max(0, @($Session.Rows).Count - $checkedCount)
}

function Get-HooterSessionMissedTreeCount {
    param([object]$Session)

    if ($null -eq $Session) { return 0 }
    $missedCount = 0
    if ($null -ne $Session.PSObject.Properties["MissedTreeCount"] -and [int]::TryParse((ConvertTo-HooterText $Session.MissedTreeCount), [ref]$missedCount)) {
        return [Math]::Max(0, $missedCount)
    }
    return @($Session.Rows | Where-Object { (ConvertTo-HooterText $_.Scope) -eq "MissedTree" -and (ConvertTo-HooterText $_.FieldName) -eq "CrewMissedTree" }).Count
}

function Get-HooterSessionUtmText {
    param([object]$Session)

    if ($null -eq $Session) { return "" }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($Session.UTMEasting, $Session.UTMNorthing, $Session.UTMZone)) {
        $text = ConvertTo-HooterText $value
        if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$parts.Add($text) }
    }
    return [string]::Join(" / ", [string[]]$parts.ToArray())
}

function Test-HooterNumberGreater {
    param(
        [object]$Left,
        [object]$Right
    )

    $leftNumber = 0.0
    $rightNumber = 0.0
    if (-not (ConvertTo-HooterNumber -Value $Left -Number ([ref]$leftNumber))) { return $false }
    if (-not (ConvertTo-HooterNumber -Value $Right -Number ([ref]$rightNumber))) { return $false }
    return ($leftNumber -gt $rightNumber)
}

function Add-HooterUniqueReportReason {
    param(
        [System.Collections.Generic.List[string]]$Reasons,
        [string]$Text
    )

    $clean = ConvertTo-HooterText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) { return }
    foreach ($existing in @($Reasons.ToArray())) {
        if ($existing.Equals($clean, [System.StringComparison]::OrdinalIgnoreCase)) { return }
    }
    [void]$Reasons.Add($clean)
}

function Get-HooterSessionFailureReasons {
    param([object]$Session)

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Session) { return @() }

    if (Test-HooterNumberGreater -Left $Session.ScoreLost -Right $Session.MaxPointLoss) {
        Add-HooterUniqueReportReason -Reasons $reasons -Text ("Total error {0} is greater than the failure threshold {1}" -f (Format-HooterScoreNumber $Session.ScoreLost), (Format-HooterScoreNumber $Session.MaxPointLoss))
    }

    foreach ($section in @(
        [pscustomobject]@{ Name = "Plot"; Loss = $Session.PlotPointLoss; Max = $Session.PlotPointTotal },
        [pscustomobject]@{ Name = "Tree"; Loss = $Session.TreePointLoss; Max = $Session.TreePointTotal },
        [pscustomobject]@{ Name = "Regen"; Loss = $Session.RegenPointLoss; Max = $Session.RegenPointTotal },
        [pscustomobject]@{ Name = "Location / Execution"; Loss = $Session.ExecutionPointLoss; Max = $Session.ExecutionPointTotal }
    )) {
        if (Test-HooterNumberGreater -Left $section.Loss -Right $section.Max) {
            Add-HooterUniqueReportReason -Reasons $reasons -Text ("{0} point loss {1} is greater than max {2}" -f $section.Name, (Format-HooterScoreNumber $section.Loss), (Format-HooterScoreNumber $section.Max))
        }
    }

    $missedCount = Get-HooterSessionMissedTreeCount -Session $Session
    if ($missedCount -gt 0) {
        $missedLabel = if ($missedCount -eq 1) { "Crew missed tree found by QA" } else { "Crew missed trees found by QA" }
        Add-HooterUniqueReportReason -Reasons $reasons -Text ("{0}: {1}" -f $missedLabel, $missedCount)
    }

    foreach ($row in @($Session.Rows | Where-Object { (ConvertTo-HooterText $_.Status) -eq "Fail" })) {
        $scope = ConvertTo-HooterText $row.Scope
        $entry = ConvertTo-HooterText $row.EntryNumber
        $field = ConvertTo-HooterText $row.FieldLabel
        $loss = Format-HooterScoreNumber $row.LostPoints
        if ($scope -eq "MissedTree") {
            $dbh = ConvertTo-HooterText $row.MissedTreeDBH
            $distance = ConvertTo-HooterText $row.MissedTreeDistance
            $azimuth = ConvertTo-HooterText $row.MissedTreeAzimuth
            Add-HooterUniqueReportReason -Reasons $reasons -Text ("Missed tree {0}: DBH {1}, distance {2}, azimuth {3}" -f $entry, $dbh, $distance, $azimuth)
        }
        elseif ($scope -eq "Execution") {
            Add-HooterUniqueReportReason -Reasons $reasons -Text ("Location / Execution - {0} marked {1}, point loss {2}" -f $field, (ConvertTo-HooterText $row.QaValue), $loss)
        }
        else {
            Add-HooterUniqueReportReason -Reasons $reasons -Text ("{0} {1} - {2} failed; crew '{3}', QA '{4}', point loss {5}" -f $scope, $entry, $field, (ConvertTo-HooterText $row.CrewValue), (ConvertTo-HooterText $row.QaValue), $loss)
        }
    }

    $criticalCount = 0
    [void][int]::TryParse((ConvertTo-HooterText $Session.CriticalFailCount), [ref]$criticalCount)
    if ($criticalCount -gt 0 -and $reasons.Count -eq 0) {
        Add-HooterUniqueReportReason -Reasons $reasons -Text ("Critical failure count: {0}" -f $criticalCount)
    }
    if ((ConvertTo-HooterText $Session.OverallStatus) -eq "Fail" -and $reasons.Count -eq 0) {
        Add-HooterUniqueReportReason -Reasons $reasons -Text "Plot status saved as Fail"
    }

    return @($reasons.ToArray())
}

function Get-HooterReportProjectName {
    param([object[]]$Sessions)

    foreach ($session in @($Sessions)) {
        $candidate = Normalize-HooterProjectName -Value $session.ProjectName
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
    }
    $candidate = Get-HooterCurrentProjectName
    if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
    return "Unknown project"
}

function Get-HooterReportPlotSummaries {
    param([object[]]$Sessions)

    $summaries = New-Object System.Collections.Generic.List[object]
    foreach ($session in @($Sessions | Sort-Object PlotNumber)) {
        $failedRows = @($session.Rows | Where-Object { (ConvertTo-HooterText $_.Status) -eq "Fail" })
        $reasons = @(Get-HooterSessionFailureReasons -Session $session)
        [void]$summaries.Add([pscustomobject]@{
            PlotNumber = ConvertTo-HooterText $session.PlotNumber
            CheckCruiserName = if ($null -ne $session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $session.CheckCruiserName } else { "" }
            CheckCruiseDate = if ($null -ne $session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $session.CheckCruiseDate } else { "" }
            OverallStatus = ConvertTo-HooterText $session.OverallStatus
            ScoreLost = Format-HooterScoreNumber $session.ScoreLost
            MaxPointLoss = Get-HooterMaxPointLossDisplay -Value $session.MaxPointLoss
            ScorePossible = Format-HooterScoreNumber $session.ScorePossible
            PlotPointLoss = Format-HooterScoreNumber $session.PlotPointLoss
            PlotPointTotal = Format-HooterScoreNumber $session.PlotPointTotal
            TreePointLoss = Format-HooterScoreNumber $session.TreePointLoss
            TreePointTotal = Format-HooterScoreNumber $session.TreePointTotal
            RegenPointLoss = Format-HooterScoreNumber $session.RegenPointLoss
            RegenPointTotal = Format-HooterScoreNumber $session.RegenPointTotal
            ExecutionPointLoss = Format-HooterScoreNumber $session.ExecutionPointLoss
            ExecutionPointTotal = Format-HooterScoreNumber $session.ExecutionPointTotal
            CriticalFailCount = ConvertTo-HooterText $session.CriticalFailCount
            FailedCount = Get-HooterSessionFailedCount -Session $session
            UncheckedCount = Get-HooterSessionUncheckedCount -Session $session
            MissedTreeCount = Get-HooterSessionMissedTreeCount -Session $session
            ErrorCount = $failedRows.Count
            UTMEasting = ConvertTo-HooterText $session.UTMEasting
            UTMNorthing = ConvertTo-HooterText $session.UTMNorthing
            UTMZone = ConvertTo-HooterText $session.UTMZone
            UTM = Get-HooterSessionUtmText -Session $session
            QA_Remarks = ConvertTo-HooterText $session.OverallNotes
            Reasons = [string]::Join("; ", [string[]]$reasons)
            SavedAt = ConvertTo-HooterText $session.SavedAt
        })
    }
    return @($summaries.ToArray())
}

function Get-HooterReportRowsForScope {
    param(
        [object[]]$Sessions,
        [string[]]$Scopes
    )

    $scopeLookup = New-InsensitiveHashtable
    foreach ($scope in @($Scopes)) { $scopeLookup[$scope] = $true }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($session in @($Sessions | Sort-Object PlotNumber)) {
        foreach ($row in @($session.Rows)) {
            $scope = ConvertTo-HooterText $row.Scope
            if (-not $scopeLookup.ContainsKey($scope)) { continue }
            $displayScope = switch ($scope) {
                "Execution" { "Location / Execution" }
                "MissedTree" { "Missed tree" }
                default { $scope }
            }
            [void]$rows.Add([pscustomobject]@{
                PlotNumber = ConvertTo-HooterText $session.PlotNumber
                CheckCruiserName = if ($null -ne $session.PSObject.Properties["CheckCruiserName"]) { ConvertTo-HooterText $session.CheckCruiserName } else { "" }
                CheckCruiseDate = if ($null -ne $session.PSObject.Properties["CheckCruiseDate"]) { ConvertTo-HooterText $session.CheckCruiseDate } else { "" }
                OverallStatus = ConvertTo-HooterText $session.OverallStatus
                UTM = Get-HooterSessionUtmText -Session $session
                Scope = $displayScope
                EntryNumber = ConvertTo-HooterText $row.EntryNumber
                CrewRecord = ConvertTo-HooterText $row.CrewRecord
                TableName = ConvertTo-HooterText $row.TableName
                FieldName = ConvertTo-HooterText $row.FieldName
                FieldLabel = ConvertTo-HooterText $row.FieldLabel
                CrewValue = ConvertTo-HooterText $row.CrewValue
                QaValue = ConvertTo-HooterText $row.QaValue
                Rule = Get-HooterRuleText ([pscustomobject]@{ Mode = $row.RuleMode; Value = $row.ToleranceValue; PointValue = $row.PointValue; CriticalFail = $row.CriticalFail })
                RuleMode = ConvertTo-HooterText $row.RuleMode
                ToleranceValue = ConvertTo-HooterText $row.ToleranceValue
                PointValue = Format-HooterScoreNumber $row.PointValue
                Status = ConvertTo-HooterText $row.Status
                LostPoints = Format-HooterScoreNumber $row.LostPoints
                PossiblePoints = Format-HooterScoreNumber $row.PossiblePoints
                CriticalFailure = ConvertTo-HooterText $row.CriticalFailure
                FieldNotes = ConvertTo-HooterText $row.FieldNotes
                MissedTreeDBH = if ($null -ne $row.PSObject.Properties["MissedTreeDBH"]) { ConvertTo-HooterText $row.MissedTreeDBH } else { "" }
                MissedTreeDistance = if ($null -ne $row.PSObject.Properties["MissedTreeDistance"]) { ConvertTo-HooterText $row.MissedTreeDistance } else { "" }
                MissedTreeAzimuth = if ($null -ne $row.PSObject.Properties["MissedTreeAzimuth"]) { ConvertTo-HooterText $row.MissedTreeAzimuth } else { "" }
                QA_Remarks = ConvertTo-HooterText $session.OverallNotes
                SavedAt = ConvertTo-HooterText $session.SavedAt
            })
        }
    }
    return @($rows.ToArray())
}

function Get-HooterDistinctAuditEntryCount {
    param(
        [object[]]$Rows,
        [string]$Scope
    )

    $keys = @{}
    foreach ($row in @($Rows)) {
        if (-not (ConvertTo-HooterText $row.Scope).Equals($Scope, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $key = "{0}|{1}|{2}" -f (ConvertTo-HooterText $row.PlotNumber), (ConvertTo-HooterText $row.EntryNumber), (ConvertTo-HooterText $row.CrewRecord)
        $keys[$key] = $true
    }
    return $keys.Count
}

function New-HooterReportColumn {
    param(
        [string]$Header,
        [string]$Property,
        [string]$Class = ""
    )

    return [pscustomobject]@{
        Header = $Header
        Property = $Property
        Class = $Class
    }
}

function New-HooterHtmlTable {
    param(
        [string]$TableId,
        [string]$Title,
        [object[]]$Rows,
        [object[]]$Columns,
        [string]$EmptyText = "No records."
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<section class='report-table-block'>")
    [void]$sb.AppendLine("<div class='table-heading'><h2>$(Escape-Html $Title)</h2><span>$(($Rows).Count) row(s)</span></div>")
    [void]$sb.AppendLine("<div class='table-tools'><input class='table-search' type='search' data-table='$(Escape-Html $TableId)' placeholder='Filter this tab'><span class='hint'>Tap a column heading to sort.</span></div>")
    [void]$sb.AppendLine("<div class='table-wrap'>")
    [void]$sb.AppendLine("<table id='$(Escape-Html $TableId)' class='data-table'>")
    [void]$sb.Append("<thead><tr>")
    foreach ($column in @($Columns)) {
        [void]$sb.Append("<th data-sort='text'>$(Escape-Html $column.Header)</th>")
    }
    [void]$sb.AppendLine("</tr></thead><tbody>")
    if (@($Rows).Count -eq 0) {
        [void]$sb.AppendLine("<tr class='empty-row'><td colspan='$(@($Columns).Count)'>$(Escape-Html $EmptyText)</td></tr>")
    }
    else {
        foreach ($row in @($Rows)) {
            $status = ConvertTo-HooterText $row.Status
            if ([string]::IsNullOrWhiteSpace($status)) { $status = ConvertTo-HooterText $row.OverallStatus }
            $rowClass = switch ($status) {
                "Fail" { " class='row-fail'" }
                "Pass" { " class='row-pass'" }
                "Incomplete" { " class='row-incomplete'" }
                default { "" }
            }
            [void]$sb.Append("<tr$rowClass>")
            foreach ($column in @($Columns)) {
                $value = ""
                if ($null -ne $row.PSObject.Properties[$column.Property]) {
                    $value = ConvertTo-HooterText $row.($column.Property)
                }
                $cellClass = ConvertTo-HooterText $column.Class
                if ((ConvertTo-HooterText $column.Property) -eq "Status" -or (ConvertTo-HooterText $column.Property) -eq "OverallStatus") {
                    if ($value -eq "Fail") { $cellClass = "$cellClass status-fail".Trim() }
                    elseif ($value -eq "Pass") { $cellClass = "$cellClass status-pass".Trim() }
                    elseif ($value -eq "Incomplete") { $cellClass = "$cellClass status-incomplete".Trim() }
                }
                $classText = if ([string]::IsNullOrWhiteSpace($cellClass)) { "" } else { " class='$cellClass'" }
                [void]$sb.Append("<td$classText>$(Escape-Html $value)</td>")
            }
            [void]$sb.AppendLine("</tr>")
        }
    }
    [void]$sb.AppendLine("</tbody></table></div></section>")
    return $sb.ToString()
}

function New-HooterQaReportHtml {
    param([object[]]$Sessions)

    $sessions = @($Sessions)
    $projectName = Get-HooterReportProjectName -Sessions $sessions
    $inventoryProgress = Get-HooterInventoryQaProgress
    $plotSummaries = @(Get-HooterReportPlotSummaries -Sessions $sessions)
    $passedPlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Pass" })
    $failedPlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Fail" })
    $incompletePlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Incomplete" })
    $plotRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Plot"))
    $treeRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Tree", "MissedTree"))
    $regenRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Regen"))
    $executionRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Execution"))
    $treeAudited = Get-HooterDistinctAuditEntryCount -Rows $treeRows -Scope "Tree"
    $regenAudited = Get-HooterDistinctAuditEntryCount -Rows $regenRows -Scope "Regen"
    $totalErrors = 0
    foreach ($summary in @($plotSummaries)) {
        $count = 0
        [void][int]::TryParse((ConvertTo-HooterText $summary.ErrorCount), [ref]$count)
        $totalErrors += $count
    }
    $missedTrees = 0
    foreach ($summary in @($plotSummaries)) {
        $count = 0
        [void][int]::TryParse((ConvertTo-HooterText $summary.MissedTreeCount), [ref]$count)
        $missedTrees += $count
    }

    $progressText = if ($inventoryProgress.TotalPlots -gt 0) {
        "$($inventoryProgress.CheckedPlots)/$($inventoryProgress.TotalPlots) plots saved ($($inventoryProgress.CompletionPercent)% complete). BIA $(Format-HooterScoreNumber $inventoryProgress.TargetPercent)% target = $($inventoryProgress.RequiredPlots) plot(s)."
    }
    else {
        "Inventory plot count unavailable."
    }
    $progressClass = if ($inventoryProgress.TargetMet) { "good" } else { "warn" }
    $progressValue = if ($inventoryProgress.RequiredPlots -gt 0) { [Math]::Min(100, [Math]::Round(($inventoryProgress.CheckedPlots / [double]$inventoryProgress.RequiredPlots) * 100.0, 1)) } else { 0 }

    $plotSummaryColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "OverallStatus"
        New-HooterReportColumn "Errors" "ErrorCount" "error-text"
        New-HooterReportColumn "Total error" "ScoreLost"
        New-HooterReportColumn "Loss threshold" "MaxPointLoss"
        New-HooterReportColumn "Critical" "CriticalFailCount"
        New-HooterReportColumn "Unchecked" "UncheckedCount"
        New-HooterReportColumn "Missed trees" "MissedTreeCount"
        New-HooterReportColumn "UTM" "UTM"
        New-HooterReportColumn "Reasons" "Reasons" "notes"
        New-HooterReportColumn "QA remarks" "QA_Remarks" "notes"
    )
    $auditColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Entry" "EntryNumber"
        New-HooterReportColumn "Crew record" "CrewRecord"
        New-HooterReportColumn "Field" "FieldLabel"
        New-HooterReportColumn "Crew value" "CrewValue"
        New-HooterReportColumn "QA value" "QaValue"
        New-HooterReportColumn "Rule" "Rule"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes" "notes"
    )
    $treeColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Type" "Scope"
        New-HooterReportColumn "Entry" "EntryNumber"
        New-HooterReportColumn "Crew record" "CrewRecord"
        New-HooterReportColumn "Field" "FieldLabel"
        New-HooterReportColumn "Crew value" "CrewValue"
        New-HooterReportColumn "QA value" "QaValue"
        New-HooterReportColumn "Missed DBH" "MissedTreeDBH"
        New-HooterReportColumn "Missed distance" "MissedTreeDistance"
        New-HooterReportColumn "Missed azimuth" "MissedTreeAzimuth"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes" "notes"
    )
    $executionColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Item" "FieldLabel"
        New-HooterReportColumn "Rating" "QaValue"
        New-HooterReportColumn "Rule" "Rule"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes" "notes"
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!doctype html>")
    [void]$sb.AppendLine("<html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>")
    [void]$sb.AppendLine("<title>$(Escape-Html $projectName) - PlotHoot QA Report</title>")
    [void]$sb.AppendLine("<style>")
    [void]$sb.AppendLine(":root{--ink:#16312d;--muted:#5f6f6a;--line:#d8e1dc;--bg:#f5f7f4;--panel:#fff;--brand:#1d514d;--brand2:#2f6b5f;--good:#23744d;--warn:#8a6119;--fail:#b22d21;--failbg:#fff1ed;--passbg:#eefaf3;--warnbg:#fff7df}")
    [void]$sb.AppendLine("*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink)}header{background:linear-gradient(135deg,var(--brand),var(--brand2));color:white;padding:20px 24px}.brand{font-size:26px;font-weight:750}.meta{color:#dcece6;margin-top:4px}.footer{font-size:12px;color:var(--muted);padding:18px 24px 28px}main{padding:18px 24px 28px}.tabs{position:sticky;top:0;z-index:10;background:var(--bg);display:flex;gap:8px;overflow:auto;padding:10px 0 12px}.tab-button{border:1px solid var(--line);background:white;color:var(--ink);border-radius:7px;padding:9px 12px;font-weight:650;white-space:nowrap;cursor:pointer}.tab-button.active{background:var(--brand);border-color:var(--brand);color:white}.tab-panel{display:none}.tab-panel.active{display:block}.metric-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:14px 0 18px}.metric{background:white;border:1px solid var(--line);border-radius:8px;padding:13px}.metric span{display:block;color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em}.metric strong{font-size:28px;display:block;margin-top:3px}.metric.error strong,.error-text{color:var(--fail);font-weight:800}.progress-card{background:white;border:1px solid var(--line);border-radius:8px;padding:14px;margin:0 0 16px}.bar{height:12px;border-radius:999px;background:#e2e9e5;overflow:hidden;margin-top:9px}.fill{height:100%;background:var(--warn);width:0}.fill.good{background:var(--good)}.report-table-block{margin:0 0 20px}.table-heading{display:flex;align-items:end;justify-content:space-between;gap:10px}.table-heading h2{font-size:18px;margin:8px 0}.table-heading span,.hint{color:var(--muted);font-size:12px}.table-tools{display:flex;align-items:center;gap:12px;margin:0 0 8px}.table-search{width:min(420px,100%);border:1px solid var(--line);border-radius:7px;padding:9px 10px;font:inherit;background:white}.table-wrap{overflow:auto;border:1px solid var(--line);background:white;border-radius:8px;max-height:70vh}table{width:100%;border-collapse:separate;border-spacing:0;min-width:900px}th,td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);padding:8px 9px;text-align:left;vertical-align:top;font-size:13px}th{position:sticky;top:0;background:#214f4a;color:white;font-weight:700;cursor:pointer;z-index:2}td:last-child,th:last-child{border-right:0}tbody tr:nth-child(even){background:#fbfcfb}.row-fail{background:var(--failbg)!important}.row-pass{background:var(--passbg)!important}.row-incomplete{background:var(--warnbg)!important}.status-fail{color:var(--fail);font-weight:800}.status-pass{color:var(--good);font-weight:800}.status-incomplete{color:var(--warn);font-weight:800}.notes{white-space:pre-wrap;min-width:240px}.empty-row td{text-align:center;color:var(--muted);padding:18px}.summary-note{color:var(--muted);margin:0 0 10px}@media print{body{background:white}main{padding:8px}.tabs,.table-tools{display:none}.tab-panel{display:block}.table-wrap{max-height:none;overflow:visible;border:0}th{position:static}.metric{break-inside:avoid}}")
    [void]$sb.AppendLine(".tab-passed{border-color:#7e57c2;background:#f2edfb}.tab-passed.active{background:#7e57c2;border-color:#7e57c2;color:white}.tab-failed{border-color:#c62828;background:#fdebea}.tab-failed.active{background:#c62828;border-color:#c62828;color:white}.tab-plotdata{border-color:#c89516;background:#fff8d8}.tab-plotdata.active{background:#f2c94c;border-color:#c89516;color:#3b2d00}.tab-treedata{border-color:#2e7d32;background:#e9f6ec}.tab-treedata.active{background:#2e7d32;border-color:#2e7d32;color:white}.tab-regendata{border-color:#1e88e5;background:#e8f3ff}.tab-regendata.active{background:#1e88e5;border-color:#1e88e5;color:white}.tab-locationdata{border-color:#ef8a24;background:#fff0df}.tab-locationdata.active{background:#ef8a24;border-color:#ef8a24;color:#221200}")
    [void]$sb.AppendLine("</style></head><body>")
    [void]$sb.AppendLine("<header><div class='brand'>PlotHoot QA Report</div><div class='meta'>Project: $(Escape-Html $projectName)</div></header>")
    [void]$sb.AppendLine("<main>")
    [void]$sb.AppendLine("<nav class='tabs' aria-label='Report tabs'>")
    [void]$sb.AppendLine("<button class='tab-button active' data-tab='summary'>Summary</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-passed' data-tab='passed'>Passed plots ($($passedPlots.Count))</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-failed' data-tab='failed'>Failed plots ($($failedPlots.Count))</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-plotdata' data-tab='plotdata'>Audited plot data</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-treedata' data-tab='treedata'>Audited tree data</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-regendata' data-tab='regendata'>Audited regen data</button>")
    [void]$sb.AppendLine("<button class='tab-button tab-locationdata' data-tab='locationdata'>Audited location data</button>")
    [void]$sb.AppendLine("</nav>")

    [void]$sb.AppendLine("<section id='summary' class='tab-panel active'>")
    [void]$sb.AppendLine("<p class='summary-note'>Use the tabs to move through the report. Each table has a filter box and sortable column headings.</p>")
    [void]$sb.AppendLine("<div class='metric-grid'>")
    [void]$sb.AppendLine("<div class='metric'><span>Plots QA saved</span><strong>$($plotSummaries.Count)</strong></div>")
    [void]$sb.AppendLine("<div class='metric'><span>Plots passed</span><strong>$($passedPlots.Count)</strong></div>")
    [void]$sb.AppendLine("<div class='metric'><span>Plots failed</span><strong class='error-text'>$($failedPlots.Count)</strong></div>")
    [void]$sb.AppendLine("<div class='metric'><span>Incomplete plots</span><strong>$($incompletePlots.Count)</strong></div>")
    [void]$sb.AppendLine("<div class='metric'><span>Trees audited</span><strong>$treeAudited</strong></div>")
    [void]$sb.AppendLine("<div class='metric'><span>Regen records audited</span><strong>$regenAudited</strong></div>")
    [void]$sb.AppendLine("<div class='metric error'><span>Total errors</span><strong>$totalErrors</strong></div>")
    [void]$sb.AppendLine("<div class='metric error'><span>Missed trees</span><strong>$missedTrees</strong></div>")
    [void]$sb.AppendLine("</div>")
    [void]$sb.AppendLine("<div class='progress-card'><strong>BIA inventory QA progress:</strong> $(Escape-Html $progressText)<div class='bar'><div class='fill $progressClass' style='width:$progressValue%'></div></div></div>")
    [void]$sb.AppendLine((New-HooterHtmlTable -TableId "summary-results" -Title "Plot Result Summary" -Rows $plotSummaries -Columns $plotSummaryColumns -EmptyText "No saved plot QA checks."))
    [void]$sb.AppendLine("</section>")
    [void]$sb.AppendLine("<section id='passed' class='tab-panel'>$(New-HooterHtmlTable -TableId "passed-plots" -Title "Passed Plots" -Rows $passedPlots -Columns $plotSummaryColumns -EmptyText "No passed plots are saved.")</section>")
    [void]$sb.AppendLine("<section id='failed' class='tab-panel'>$(New-HooterHtmlTable -TableId "failed-plots" -Title "Failed Plots and Reasons" -Rows $failedPlots -Columns $plotSummaryColumns -EmptyText "No failed plots are saved.")</section>")
    [void]$sb.AppendLine("<section id='plotdata' class='tab-panel'>$(New-HooterHtmlTable -TableId "audited-plot-data" -Title "Audited Plot Data" -Rows $plotRows -Columns $auditColumns -EmptyText "No audited plot rows are saved.")</section>")
    [void]$sb.AppendLine("<section id='treedata' class='tab-panel'>$(New-HooterHtmlTable -TableId "audited-tree-data" -Title "Audited Tree Data" -Rows $treeRows -Columns $treeColumns -EmptyText "No audited tree rows are saved.")</section>")
    [void]$sb.AppendLine("<section id='regendata' class='tab-panel'>$(New-HooterHtmlTable -TableId "audited-regen-data" -Title "Audited Regen Data" -Rows $regenRows -Columns $auditColumns -EmptyText "No audited regen rows are saved.")</section>")
    [void]$sb.AppendLine("<section id='locationdata' class='tab-panel'>$(New-HooterHtmlTable -TableId "audited-location-data" -Title "Audited Location Data" -Rows $executionRows -Columns $executionColumns -EmptyText "No audited location/execution rows are saved.")</section>")
    [void]$sb.AppendLine("</main>")
    [void]$sb.AppendLine("<div class='footer'>Developed by BIA Division of Forestry, Branch of Inventory and planning. Report is offline HTML and can be opened on a tablet or desktop.</div>")
    [void]$sb.AppendLine("<script>")
    [void]$sb.AppendLine('(function(){')
    [void]$sb.AppendLine('function showTab(id){document.querySelectorAll(".tab-panel").forEach(function(panel){panel.classList.toggle("active",panel.id===id);});document.querySelectorAll(".tab-button").forEach(function(button){button.classList.toggle("active",button.getAttribute("data-tab")===id);});}')
    [void]$sb.AppendLine('document.querySelectorAll(".tab-button").forEach(function(button){button.addEventListener("click",function(){showTab(button.getAttribute("data-tab"));});});')
    [void]$sb.AppendLine('document.querySelectorAll(".table-search").forEach(function(input){input.addEventListener("input",function(){var table=document.getElementById(input.getAttribute("data-table"));if(!table||!table.tBodies.length){return;}var term=input.value.toLowerCase();Array.prototype.forEach.call(table.tBodies[0].rows,function(row){if(row.classList.contains("empty-row")){return;}row.style.display=row.textContent.toLowerCase().indexOf(term)>=0?"":"none";});});});')
    [void]$sb.AppendLine('document.querySelectorAll("th[data-sort]").forEach(function(th){th.addEventListener("click",function(){var table=th.closest("table");if(!table||!table.tBodies.length){return;}var index=Array.prototype.indexOf.call(th.parentNode.children,th);var asc=th.getAttribute("data-dir")!=="asc";Array.prototype.forEach.call(th.parentNode.children,function(other){other.removeAttribute("data-dir");});th.setAttribute("data-dir",asc?"asc":"desc");var rows=Array.prototype.slice.call(table.tBodies[0].rows).filter(function(row){return !row.classList.contains("empty-row");});rows.sort(function(a,b){var av=a.cells[index]?a.cells[index].textContent.trim():"";var bv=b.cells[index]?b.cells[index].textContent.trim():"";var an=parseFloat(av.replace(/,/g,""));var bn=parseFloat(bv.replace(/,/g,""));if(!isNaN(an)&&!isNaN(bn)){return asc?an-bn:bn-an;}return asc?av.localeCompare(bv):bv.localeCompare(av);});rows.forEach(function(row){table.tBodies[0].appendChild(row);});});});')
    [void]$sb.AppendLine('})();')
    [void]$sb.AppendLine("</script>")
    [void]$sb.AppendLine("</body></html>")
    return $sb.ToString()
}

function Export-HooterQaReportHtml {
    try {
        if ($null -ne $script:CurrentPlot) { Save-HooterCurrentSession }
        if ($script:QaSessions.Count -eq 0) { throw "There are no saved plot QA checks to export." }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
        $dialog.FileName = [System.IO.Path]::GetFileName((Get-HooterDefaultExportPath -Suffix "QA_Report" -Extension "html"))
        $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        New-HooterQaReportHtml -Sessions @($script:QaSessions.ToArray()) | Set-Content -LiteralPath $dialog.FileName -Encoding UTF8
        Set-HooterStatus "Exported QA report HTML: $($dialog.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export QA report HTML", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Escape-HooterXml {
    param([object]$Value)

    $text = ConvertTo-HooterText $Value
    if ([string]::IsNullOrEmpty($text)) { return "" }
    return [System.Security.SecurityElement]::Escape($text)
}

function Get-HooterExcelColumnName {
    param([int]$Index)

    $name = ""
    $value = $Index
    while ($value -gt 0) {
        $value--
        $name = [char](65 + ($value % 26)) + $name
        $value = [math]::Floor($value / 26)
    }
    return $name
}

function New-HooterXlsxCell {
    param(
        [object]$Value,
        [int]$StyleId = 0
    )

    return [pscustomobject]@{
        Value = ConvertTo-HooterText $Value
        StyleId = $StyleId
    }
}

function Add-HooterXlsxRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [object[]]$Values,
        [int]$StyleId = 0,
        [hashtable]$StyleByIndex = $null
    )

    $cells = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt @($Values).Count; $index++) {
        $cellStyle = $StyleId
        if ($null -ne $StyleByIndex -and $StyleByIndex.ContainsKey($index)) {
            $cellStyle = [int]$StyleByIndex[$index]
        }
        [void]$cells.Add((New-HooterXlsxCell -Value $Values[$index] -StyleId $cellStyle))
    }
    [void]$Rows.Add([object[]]$cells.ToArray())
}

function Add-HooterXlsxObjectTableRows {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [object[]]$DataRows,
        [object[]]$Columns
    )

    Add-HooterXlsxRow -Rows $Rows -Values @($Columns | ForEach-Object { $_.Header }) -StyleId 1
    foreach ($row in @($DataRows)) {
        $values = New-Object System.Collections.Generic.List[object]
        $styleByIndex = @{}
        for ($index = 0; $index -lt @($Columns).Count; $index++) {
            $column = $Columns[$index]
            $value = ""
            if ($null -ne $row.PSObject.Properties[$column.Property]) {
                $value = ConvertTo-HooterText $row.($column.Property)
            }
            [void]$values.Add($value)
            if ($column.Property -in @("Status", "OverallStatus")) {
                if ($value -eq "Fail") { $styleByIndex[$index] = 3 }
                elseif ($value -eq "Pass") { $styleByIndex[$index] = 4 }
                elseif ($value -eq "Incomplete") { $styleByIndex[$index] = 5 }
            }
            elseif ($column.Property -in @("ErrorCount", "FailedCount", "MissedTreeCount", "CriticalFailCount")) {
                $number = 0.0
                if ((ConvertTo-HooterNumber -Value $value -Number ([ref]$number)) -and $number -gt 0) {
                    $styleByIndex[$index] = 6
                }
            }
        }
        Add-HooterXlsxRow -Rows $Rows -Values @($values.ToArray()) -StyleByIndex $styleByIndex
    }
}

function Get-HooterXlsxCellXml {
    param(
        [object]$Cell,
        [string]$Reference
    )

    $style = ""
    if ($null -ne $Cell -and $Cell.PSObject.Properties["StyleId"] -and [int]$Cell.StyleId -gt 0) {
        $style = " s=""$([int]$Cell.StyleId)"""
    }
    $value = if ($null -eq $Cell) { "" } else { Escape-HooterXml $Cell.Value }
    if ([string]::IsNullOrWhiteSpace($value) -and [string]::IsNullOrWhiteSpace($style)) {
        return "<c r=""$Reference""/>"
    }
    return "<c r=""$Reference""$style t=""inlineStr""><is><t>$value</t></is></c>"
}

function Get-HooterXlsxWorksheetXml {
    param(
        [object]$Sheet
    )

    $rows = @($Sheet.Rows)
    $rowCount = [Math]::Max(1, $rows.Count)
    $maxCols = 1
    foreach ($row in $rows) {
        $maxCols = [Math]::Max($maxCols, @($row).Count)
    }
    $lastCell = "$(Get-HooterExcelColumnName $maxCols)$rowCount"
    $freezeRow = 1
    if ($null -ne $Sheet.PSObject.Properties["FreezeRow"]) {
        [void][int]::TryParse((ConvertTo-HooterText $Sheet.FreezeRow), [ref]$freezeRow)
    }
    $freezeXml = ""
    if ($freezeRow -gt 0) {
        $topLeft = "A$($freezeRow + 1)"
        $freezeXml = "<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""$freezeRow"" topLeftCell=""$topLeft"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>"
    }
    else {
        $freezeXml = "<sheetViews><sheetView workbookViewId=""0""/></sheetViews>"
    }

    $widths = @{}
    for ($col = 1; $col -le $maxCols; $col++) { $widths[$col] = 10 }
    foreach ($row in $rows) {
        for ($col = 1; $col -le @($row).Count; $col++) {
            $textLength = (ConvertTo-HooterText $row[$col - 1].Value).Length
            $widths[$col] = [Math]::Min(48, [Math]::Max([int]$widths[$col], [Math]::Max(10, $textLength + 2)))
        }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
    [void]$sb.AppendLine("<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">")
    $tabColor = ""
    if ($null -ne $Sheet.PSObject.Properties["TabColor"]) {
        $tabColor = (ConvertTo-HooterText $Sheet.TabColor).TrimStart("#")
    }
    if (-not [string]::IsNullOrWhiteSpace($tabColor)) {
        if ($tabColor.Length -eq 6) { $tabColor = "FF$tabColor" }
        [void]$sb.AppendLine("<sheetPr><tabColor rgb=""$(Escape-HooterXml $tabColor)""/></sheetPr>")
    }
    [void]$sb.AppendLine("<dimension ref=""A1:$lastCell""/>")
    [void]$sb.AppendLine($freezeXml)
    [void]$sb.AppendLine("<cols>")
    for ($col = 1; $col -le $maxCols; $col++) {
        $width = ([double]$widths[$col]).ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
        [void]$sb.AppendLine("<col min=""$col"" max=""$col"" width=""$width"" customWidth=""1""/>")
    }
    [void]$sb.AppendLine("</cols><sheetData>")
    for ($rowIndex = 1; $rowIndex -le $rows.Count; $rowIndex++) {
        $row = @($rows[$rowIndex - 1])
        [void]$sb.Append("<row r=""$rowIndex"">")
        for ($col = 1; $col -le $row.Count; $col++) {
            $cellRef = "$(Get-HooterExcelColumnName $col)$rowIndex"
            [void]$sb.Append((Get-HooterXlsxCellXml -Cell $row[$col - 1] -Reference $cellRef))
        }
        [void]$sb.AppendLine("</row>")
    }
    [void]$sb.AppendLine("</sheetData>")
    $autoFilterRow = 0
    if ($null -ne $Sheet.PSObject.Properties["AutoFilterRow"]) {
        [void][int]::TryParse((ConvertTo-HooterText $Sheet.AutoFilterRow), [ref]$autoFilterRow)
    }
    if ($autoFilterRow -gt 0 -and $rows.Count -gt $autoFilterRow) {
        $filterRef = "A${autoFilterRow}:$(Get-HooterExcelColumnName $maxCols)$($rows.Count)"
        [void]$sb.AppendLine("<autoFilter ref=""$filterRef""/>")
    }
    [void]$sb.AppendLine("<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>")
    [void]$sb.AppendLine("</worksheet>")
    return $sb.ToString()
}

function Get-HooterXlsxStylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="6">
    <font><sz val="11"/><color theme="1"/><name val="Segoe UI"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/><family val="2"/></font>
    <font><b/><sz val="16"/><color rgb="FF16312D"/><name val="Segoe UI"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFB22D21"/><name val="Segoe UI"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF23744D"/><name val="Segoe UI"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF8A6119"/><name val="Segoe UI"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF214F4A"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF1ED"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEEFAF3"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF7DF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD8E1DC"/></left><right style="thin"><color rgb="FFD8E1DC"/></right><top style="thin"><color rgb="FFD8E1DC"/></top><bottom style="thin"><color rgb="FFD8E1DC"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@
}

function Get-HooterSafeWorksheetName {
    param(
        [string]$Name,
        [hashtable]$UsedNames
    )

    $clean = ConvertTo-HooterText $Name
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "Sheet" }
    $clean = [regex]::Replace($clean, "[\\/\?\*\[\]:]", " ")
    $clean = [regex]::Replace($clean, "\s+", " ").Trim()
    if ($clean.Length -gt 31) { $clean = $clean.Substring(0, 31).Trim() }
    $base = if ([string]::IsNullOrWhiteSpace($clean)) { "Sheet" } else { $clean }
    $candidate = $base
    $counter = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $suffix = " $counter"
        $maxBase = [Math]::Max(1, 31 - $suffix.Length)
        $candidate = $base
        if ($candidate.Length -gt $maxBase) { $candidate = $candidate.Substring(0, $maxBase).Trim() }
        $candidate = "$candidate$suffix"
        $counter++
    }
    $UsedNames[$candidate] = $true
    return $candidate
}

function Write-HooterXlsxPart {
    param(
        [string]$Path,
        [string]$Text
    )

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) {
        [void][System.IO.Directory]::CreateDirectory($folder)
    }
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function New-HooterXlsxPackage {
    param(
        [string]$Path,
        [object[]]$Sheets,
        [string]$ProjectName
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("PlotHoot_Xlsx_{0}" -f ([guid]::NewGuid().ToString("N")))
    [void][System.IO.Directory]::CreateDirectory($tempRoot)
    try {
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "_rels"))
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "docProps"))
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl"))
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl\_rels"))
        [void][System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl\worksheets"))

        $contentTypes = New-Object System.Text.StringBuilder
        [void]$contentTypes.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$contentTypes.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        [void]$contentTypes.AppendLine('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        [void]$contentTypes.AppendLine('<Default Extension="xml" ContentType="application/xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        for ($i = 1; $i -le @($Sheets).Count; $i++) {
            [void]$contentTypes.AppendLine("<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
        }
        [void]$contentTypes.AppendLine('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
        [void]$contentTypes.AppendLine('</Types>')
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "[Content_Types].xml") -Text $contentTypes.ToString()

        Write-HooterXlsxPart -Path (Join-Path $tempRoot "_rels\.rels") -Text @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

        $workbook = New-Object System.Text.StringBuilder
        [void]$workbook.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbook.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
        for ($i = 1; $i -le @($Sheets).Count; $i++) {
            $sheetName = Escape-HooterXml $Sheets[$i - 1].Name
            [void]$workbook.AppendLine("<sheet name=""$sheetName"" sheetId=""$i"" r:id=""rId$i""/>")
        }
        [void]$workbook.AppendLine('</sheets></workbook>')
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "xl\workbook.xml") -Text $workbook.ToString()

        $rels = New-Object System.Text.StringBuilder
        [void]$rels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$rels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 1; $i -le @($Sheets).Count; $i++) {
            [void]$rels.AppendLine("<Relationship Id=""rId$i"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$i.xml""/>")
        }
        $stylesId = @($Sheets).Count + 1
        [void]$rels.AppendLine("<Relationship Id=""rId$stylesId"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>")
        [void]$rels.AppendLine('</Relationships>')
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "xl\_rels\workbook.xml.rels") -Text $rels.ToString()
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "xl\styles.xml") -Text (Get-HooterXlsxStylesXml)

        for ($i = 1; $i -le @($Sheets).Count; $i++) {
            Write-HooterXlsxPart -Path (Join-Path $tempRoot "xl\worksheets\sheet$i.xml") -Text (Get-HooterXlsxWorksheetXml -Sheet $Sheets[$i - 1])
        }

        $created = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "docProps\core.xml") -Text "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>$(Escape-HooterXml $ProjectName) PlotHoot QA Results</dc:title><dc:creator>PlotHoot</dc:creator><cp:lastModifiedBy>PlotHoot</cp:lastModifiedBy><dcterms:created xsi:type=""dcterms:W3CDTF"">$created</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">$created</dcterms:modified></cp:coreProperties>"
        $sheetNamesXml = [string]::Join("", [string[]]@($Sheets | ForEach-Object { "<vt:lpstr>$(Escape-HooterXml $_.Name)</vt:lpstr>" }))
        Write-HooterXlsxPart -Path (Join-Path $tempRoot "docProps\app.xml") -Text "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>PlotHoot</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>$(@($Sheets).Count)</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=""$(@($Sheets).Count)"" baseType=""lpstr"">$sheetNamesXml</vt:vector></TitlesOfParts><Company>BIA Division of Forestry, Branch of Inventory and planning</Company></Properties>"

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        $zipStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            $archive = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($file in Get-ChildItem -LiteralPath $tempRoot -Recurse -File) {
                    $relative = $file.FullName.Substring($tempRoot.Length).TrimStart([char[]]@("\", "/"))
                    $entryName = $relative.Replace("\", "/")
                    $entry = $archive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
                    $entryStream = $entry.Open()
                    $fileStream = [System.IO.File]::OpenRead($file.FullName)
                    try {
                        $fileStream.CopyTo($entryStream)
                    }
                    finally {
                        $fileStream.Dispose()
                        $entryStream.Dispose()
                    }
                }
            }
            finally {
                $archive.Dispose()
            }
        }
        finally {
            $zipStream.Dispose()
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

function New-HooterQaWorkbookXlsx {
    param(
        [object[]]$Sessions,
        [string]$Path
    )

    $sessions = @($Sessions)
    $projectName = Get-HooterReportProjectName -Sessions $sessions
    $inventoryProgress = Get-HooterInventoryQaProgress
    $plotSummaries = @(Get-HooterReportPlotSummaries -Sessions $sessions)
    $passedPlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Pass" })
    $failedPlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Fail" })
    $incompletePlots = @($plotSummaries | Where-Object { (ConvertTo-HooterText $_.OverallStatus) -eq "Incomplete" })
    $plotRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Plot"))
    $treeRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Tree", "MissedTree"))
    $regenRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Regen"))
    $executionRows = @(Get-HooterReportRowsForScope -Sessions $sessions -Scopes @("Execution"))
    $setupRows = @(Get-HooterSetupExportRows)
    $treeAudited = Get-HooterDistinctAuditEntryCount -Rows $treeRows -Scope "Tree"
    $regenAudited = Get-HooterDistinctAuditEntryCount -Rows $regenRows -Scope "Regen"
    $totalErrors = 0
    $missedTrees = 0
    foreach ($summary in @($plotSummaries)) {
        $count = 0
        [void][int]::TryParse((ConvertTo-HooterText $summary.ErrorCount), [ref]$count)
        $totalErrors += $count
        $missed = 0
        [void][int]::TryParse((ConvertTo-HooterText $summary.MissedTreeCount), [ref]$missed)
        $missedTrees += $missed
    }
    $progressText = if ($inventoryProgress.TotalPlots -gt 0) {
        "$($inventoryProgress.CheckedPlots)/$($inventoryProgress.TotalPlots) plots saved ($($inventoryProgress.CompletionPercent)% complete). BIA $(Format-HooterScoreNumber $inventoryProgress.TargetPercent)% target = $($inventoryProgress.RequiredPlots) plot(s)."
    }
    else {
        "Inventory plot count unavailable."
    }

    $plotSummaryColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "OverallStatus"
        New-HooterReportColumn "Errors" "ErrorCount"
        New-HooterReportColumn "Total error" "ScoreLost"
        New-HooterReportColumn "Loss threshold" "MaxPointLoss"
        New-HooterReportColumn "Critical" "CriticalFailCount"
        New-HooterReportColumn "Unchecked" "UncheckedCount"
        New-HooterReportColumn "Missed trees" "MissedTreeCount"
        New-HooterReportColumn "UTM" "UTM"
        New-HooterReportColumn "Reasons" "Reasons"
        New-HooterReportColumn "QA remarks" "QA_Remarks"
    )
    $auditColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Entry" "EntryNumber"
        New-HooterReportColumn "Crew record" "CrewRecord"
        New-HooterReportColumn "Field" "FieldLabel"
        New-HooterReportColumn "Crew value" "CrewValue"
        New-HooterReportColumn "QA value" "QaValue"
        New-HooterReportColumn "Rule" "Rule"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes"
    )
    $treeColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Type" "Scope"
        New-HooterReportColumn "Entry" "EntryNumber"
        New-HooterReportColumn "Crew record" "CrewRecord"
        New-HooterReportColumn "Field" "FieldLabel"
        New-HooterReportColumn "Crew value" "CrewValue"
        New-HooterReportColumn "QA value" "QaValue"
        New-HooterReportColumn "Missed DBH" "MissedTreeDBH"
        New-HooterReportColumn "Missed distance" "MissedTreeDistance"
        New-HooterReportColumn "Missed azimuth" "MissedTreeAzimuth"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes"
    )
    $executionColumns = @(
        New-HooterReportColumn "Plot" "PlotNumber"
        New-HooterReportColumn "Check cruiser" "CheckCruiserName"
        New-HooterReportColumn "Check date" "CheckCruiseDate"
        New-HooterReportColumn "Status" "Status"
        New-HooterReportColumn "Item" "FieldLabel"
        New-HooterReportColumn "Rating" "QaValue"
        New-HooterReportColumn "Rule" "Rule"
        New-HooterReportColumn "Point loss" "LostPoints"
        New-HooterReportColumn "Critical" "CriticalFailure"
        New-HooterReportColumn "Notes" "FieldNotes"
    )
    $setupColumns = @(
        New-HooterReportColumn "Scope" "Scope"
        New-HooterReportColumn "Field" "FieldLabel"
        New-HooterReportColumn "Table" "TableName"
        New-HooterReportColumn "Field name" "FieldName"
        New-HooterReportColumn "Tolerance" "RuleMode"
        New-HooterReportColumn "Value" "ToleranceValue"
        New-HooterReportColumn "Points" "PointValue"
        New-HooterReportColumn "Critical fail" "CriticalFail"
        New-HooterReportColumn "Field order" "FieldOrder"
        New-HooterReportColumn "Loss threshold" "MaxPointLoss"
        New-HooterReportColumn "Plot max" "PlotPointTotal"
        New-HooterReportColumn "Tree max/tree" "TreePointTotal"
        New-HooterReportColumn "Regen max" "RegenPointTotal"
        New-HooterReportColumn "Table D max" "ExecutionPointTotal"
    )

    $usedNames = @{}
    $sheets = New-Object System.Collections.Generic.List[object]
    $summaryRows = New-Object System.Collections.Generic.List[object]
    Add-HooterXlsxRow -Rows $summaryRows -Values @("PlotHoot QA Results") -StyleId 2
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Project", $projectName)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("BIA inventory QA progress", $progressText)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("")
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Metric", "Value") -StyleId 1
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Plots QA saved", $plotSummaries.Count)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Plots passed", $passedPlots.Count)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Plots failed", $failedPlots.Count) -StyleByIndex @{ 1 = 6 }
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Incomplete plots", $incompletePlots.Count)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Trees audited", $treeAudited)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Regen records audited", $regenAudited)
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Total errors", $totalErrors) -StyleByIndex @{ 1 = 6 }
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Missed trees", $missedTrees) -StyleByIndex @{ 1 = 6 }
    Add-HooterXlsxRow -Rows $summaryRows -Values @("")
    Add-HooterXlsxRow -Rows $summaryRows -Values @("Plot Result Summary") -StyleId 2
    $summaryFilterRow = $summaryRows.Count + 1
    Add-HooterXlsxObjectTableRows -Rows $summaryRows -DataRows $plotSummaries -Columns $plotSummaryColumns
    [void]$sheets.Add([pscustomobject]@{ Name = Get-HooterSafeWorksheetName -Name "Summary" -UsedNames $usedNames; Rows = @($summaryRows.ToArray()); AutoFilterRow = $summaryFilterRow; FreezeRow = 1 })

    foreach ($sheetInfo in @(
        [pscustomobject]@{ Name = "Passed Plots"; Rows = $passedPlots; Columns = $plotSummaryColumns; Empty = "No passed plots are saved."; TabColor = "7E57C2" },
        [pscustomobject]@{ Name = "Failed Plots"; Rows = $failedPlots; Columns = $plotSummaryColumns; Empty = "No failed plots are saved."; TabColor = "C62828" },
        [pscustomobject]@{ Name = "Audited Plot Data"; Rows = $plotRows; Columns = $auditColumns; Empty = "No audited plot rows are saved."; TabColor = "F2C94C" },
        [pscustomobject]@{ Name = "Audited Tree Data"; Rows = $treeRows; Columns = $treeColumns; Empty = "No audited tree rows are saved."; TabColor = "2E7D32" },
        [pscustomobject]@{ Name = "Audited Regen Data"; Rows = $regenRows; Columns = $auditColumns; Empty = "No audited regen rows are saved."; TabColor = "1E88E5" },
        [pscustomobject]@{ Name = "Audited Location Data"; Rows = $executionRows; Columns = $executionColumns; Empty = "No audited location rows are saved."; TabColor = "EF8A24" },
        [pscustomobject]@{ Name = "Setup Settings"; Rows = $setupRows; Columns = $setupColumns; Empty = "No setup rows are available."; TabColor = "" }
    )) {
        $rows = New-Object System.Collections.Generic.List[object]
        Add-HooterXlsxObjectTableRows -Rows $rows -DataRows @($sheetInfo.Rows) -Columns @($sheetInfo.Columns)
        if (@($sheetInfo.Rows).Count -eq 0) {
            Add-HooterXlsxRow -Rows $rows -Values @($sheetInfo.Empty)
        }
        [void]$sheets.Add([pscustomobject]@{ Name = Get-HooterSafeWorksheetName -Name $sheetInfo.Name -UsedNames $usedNames; Rows = @($rows.ToArray()); AutoFilterRow = 1; FreezeRow = 1; TabColor = $sheetInfo.TabColor })
    }

    New-HooterXlsxPackage -Path $Path -Sheets @($sheets.ToArray()) -ProjectName $projectName
}

function Export-HooterQaWorkbookXlsx {
    try {
        if ($null -ne $script:CurrentPlot) { Save-HooterCurrentSession }
        if ($script:QaSessions.Count -eq 0) { throw "There are no saved plot QA checks to export." }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        $dialog.FileName = [System.IO.Path]::GetFileName((Get-HooterDefaultExportPath -Suffix "QA_Results" -Extension "xlsx"))
        $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        New-HooterQaWorkbookXlsx -Sessions @($script:QaSessions.ToArray()) -Path $dialog.FileName
        Set-HooterStatus "Exported QA Excel workbook: $($dialog.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export QA Excel workbook", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Export-HooterFailedCsv {
    try {
        if ($null -ne $script:CurrentPlot) { Save-HooterCurrentSession }
        $summaries = @(Get-HooterFailedSummaries)
        if ($summaries.Count -eq 0) { throw "No failed plots are saved right now." }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $dialog.FileName = [System.IO.Path]::GetFileName((Get-HooterDefaultExportPath -Suffix "Failed_Plots" -Extension "csv"))
        $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $summaries | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
        Set-HooterStatus "Exported failed plots CSV: $($dialog.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export failed plots CSV", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function New-HooterFailedHtml {
    param([object[]]$Summaries)

    $projectName = ""
    if ($Summaries.Count -gt 0) {
        $firstSummary = $Summaries | Select-Object -First 1
        $projectName = Normalize-HooterProjectName -Value $firstSummary.ProjectName
    }
    if ([string]::IsNullOrWhiteSpace($projectName)) {
        $projectName = Get-HooterCurrentProjectName
    }
    if ([string]::IsNullOrWhiteSpace($projectName)) { $projectName = "Unknown project" }
    $inventoryProgress = Get-HooterInventoryQaProgress
    if ($Summaries.Count -gt 0) {
        $firstProgressSummary = $Summaries | Select-Object -First 1
        $summaryTotal = 0
        if ($null -ne $firstProgressSummary.PSObject.Properties["InventoryTotalPlots"] -and (ConvertTo-HooterNumber -Value $firstProgressSummary.InventoryTotalPlots -Number ([ref]$summaryTotal)) -and $summaryTotal -gt 0) {
            $summaryChecked = 0
            [void][int]::TryParse((ConvertTo-HooterText $firstProgressSummary.InventoryCheckedPlots), [ref]$summaryChecked)
            $targetPercent = Get-HooterInventoryQaTargetPercent
            $requiredPlots = [int][Math]::Ceiling($summaryTotal * ($targetPercent / 100.0))
            $completionPercent = [Math]::Round(($summaryChecked / [double]$summaryTotal) * 100.0, 1)
            $inventoryProgress = [pscustomobject]@{
                TotalPlots = [int]$summaryTotal
                CheckedPlots = $summaryChecked
                TargetPercent = $targetPercent
                RequiredPlots = $requiredPlots
                CompletionPercent = $completionPercent
            }
        }
    }
    $progressText = if ($inventoryProgress.TotalPlots -gt 0) {
        "$($inventoryProgress.CheckedPlots)/$($inventoryProgress.TotalPlots) plots saved ($($inventoryProgress.CompletionPercent)% complete); BIA $(Format-HooterScoreNumber $inventoryProgress.TargetPercent)% target = $($inventoryProgress.RequiredPlots) plots"
    }
    else {
        "Inventory plot count unavailable"
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!doctype html>")
    [void]$sb.AppendLine("<html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>")
    [void]$sb.AppendLine("<title>$(Escape-Html $projectName) - PlotHoot Failed Plots</title>")
    [void]$sb.AppendLine("<style>")
    [void]$sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f7f4;color:#162b28}header{background:#1e4744;color:white;padding:22px 28px}main{padding:24px}.brand{font-size:26px;font-weight:700}.meta{color:#d7e7df;margin-top:4px}.panel{background:white;border:1px solid #d9e1dc;border-radius:8px;padding:18px;margin-bottom:18px}table{width:100%;border-collapse:collapse;background:white}th{background:#1e4744;color:white;text-align:left}th,td{border:1px solid #d9e1dc;padding:9px;vertical-align:top}.fail{color:#aa3623;font-weight:700}.notes{white-space:pre-wrap}.footer{font-size:12px;color:#52615d;margin-top:24px}@media print{body{background:white}main{padding:0}.panel{border:0}}")
    [void]$sb.AppendLine("</style></head><body>")
    [void]$sb.AppendLine("<header><div class='brand'>PlotHoot Failed Plot Report</div><div class='meta'>Project: $(Escape-Html $projectName)</div></header><main>")
    [void]$sb.AppendLine("<div class='panel'><strong>Project:</strong> $(Escape-Html $projectName)<br><strong>Failed plots:</strong> $($Summaries.Count)<br><strong>BIA inventory QA progress:</strong> $(Escape-Html $progressText)</div>")
    [void]$sb.AppendLine("<table><thead><tr><th>Plot</th><th>Check cruiser</th><th>Check date</th><th>Status</th><th>Score</th><th>Critical</th><th>Missed trees</th><th>Failed checks</th><th>Unchecked</th><th>UTM</th><th>QA remarks</th><th>Failed fields</th></tr></thead><tbody>")
    foreach ($summary in @($Summaries)) {
        $utm = [string]::Join(" / ", [string[]]@($summary.UTMEasting, $summary.UTMNorthing, $summary.UTMZone | Where-Object { -not [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $_)) }))
        $score = "$(ConvertTo-HooterText $summary.Score), Plot max $(ConvertTo-HooterText $summary.PlotPointTotal), Tree max effective $(ConvertTo-HooterText $summary.TreePointTotal), Regen max $(ConvertTo-HooterText $summary.RegenPointTotal), Table D max $(ConvertTo-HooterText $summary.ExecutionPointTotal)"
        [void]$sb.AppendLine("<tr><td>$(Escape-Html $summary.PlotNumber)</td><td>$(Escape-Html $summary.CheckCruiserName)</td><td>$(Escape-Html $summary.CheckCruiseDate)</td><td class='fail'>$(Escape-Html $summary.OverallStatus)</td><td>$(Escape-Html $score)</td><td>$(Escape-Html $summary.CriticalFailCount)</td><td>$(Escape-Html $summary.MissedTreeCount)</td><td>$(Escape-Html $summary.FailedCount)</td><td>$(Escape-Html $summary.UncheckedCount)</td><td>$(Escape-Html $utm)</td><td class='notes'>$(Escape-Html $summary.QA_Remarks)</td><td>$(Escape-Html $summary.FailedFields)</td></tr>")
    }
    [void]$sb.AppendLine("</tbody></table>")
    [void]$sb.AppendLine("<div class='footer'>Developed by BIA Division of Forestry, Branch of Inventory and planning. Report is offline HTML and can be opened on a tablet or desktop.</div>")
    [void]$sb.AppendLine("</main></body></html>")
    return $sb.ToString()
}

function Export-HooterFailedHtml {
    try {
        if ($null -ne $script:CurrentPlot) { Save-HooterCurrentSession }
        $summaries = @(Get-HooterFailedSummaries)
        if ($summaries.Count -eq 0) { throw "No failed plots are saved right now." }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
        $dialog.FileName = [System.IO.Path]::GetFileName((Get-HooterDefaultExportPath -Suffix "Failed_Plots" -Extension "html"))
        $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        New-HooterFailedHtml -Summaries $summaries | Set-Content -LiteralPath $dialog.FileName -Encoding UTF8
        Set-HooterStatus "Exported failed plots HTML: $($dialog.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export failed plots HTML", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}

function Get-HooterInAppGuideText {
    return @"
PlotHoot User Guide

Purpose
PlotHoot is an offline QA field checker. It reads the selected Access database, shows the crew/database values, and gives the QA cruiser blank cells to enter independent QA values. PlotHoot does not write QA data back to the database.

Platform Support
PlotHoot runs on Windows desktops and Windows tablets such as Microsoft Surface. The launchers do not run on Android or iOS. Android and iOS devices can open exported CSV files and offline HTML reports, but they cannot run the PlotHoot app or connect directly to Access .mdb/.accdb databases.

Launcher Options
For normal field use, start PlotHoot with PlotHoot.lnk from the PlotHoot folder. The shortcut starts PlotHoot hidden through Windows PowerShell and uses the included PlotHoot icon.

When moving PlotHoot to a tablet, copy the entire PlotHoot version folder. Do not copy only PlotHoot.lnk, PlotHoot-Clean.cmd, or Create PlotHoot Desktop Shortcut.cmd; those launchers need the PlotHoot App Files folder beside them.

An owl loading screen appears while PlotHoot prepares the app window.

If the folder was copied to a tablet and the shortcut has no icon or does not open, run Create PlotHoot Desktop Shortcut.cmd once while signed into that tablet account. It repairs the parent-folder shortcut and creates a desktop shortcut with the PlotHoot icon for that Windows user. Run it once for the guest account and once for the admin account if both need shortcuts.

For security review or visible troubleshooting, use PlotHoot-Clean.cmd. This launcher does not use VBS, encoded payloads, or dynamic script blocks. It uses a per-launch execution-policy bypass so locked-down tablet or guest accounts can start PlotHoot without changing Windows policy or needing admin rights. It opens a command window and may sit for a few seconds while PlotHoot starts; that is expected for this troubleshooting launcher.

The app files also include PlotHoot.vbs and the backup PlotHoot-Backup.vbs launcher. Keep those files in PlotHoot App Files with PlotHoot.ps1.

Quick Start
1. Click Browse and choose the master or project database.
2. Click Load Completed Plots to load only plots with InventoryAssignmentPlots StatusID = 4 into the dropdown. Use Load All Plots only when you need a plot that has not been marked completed.
3. Confirm the project name shown in the header.
4. Type the plot number directly, or choose it from the optional dropdown.
5. Press Enter in the plot box or click Load plot.
6. Enter QA values on Plot Data, Tree Data, and Regen Data.
7. Add the required check cruiser name, check cruise date, and per-plot QA remarks on Review / Export.
8. Click Save QA in Review / Export.
9. Export QA Results, the QA CSV, failed plot CSV, or QA Report HTML.

Project Name
After connecting, PlotHoot shows the project name in the header. CSV and HTML exports include the same project name. If the database does not provide a project-name field, PlotHoot uses the database filename and removes trailing date, version, and database text.

Plot Data
The Plot Data tab loads one set of plot-level QA rows. Crew, measurement day/date, and plot number use a simple filled-in checkbox check that starts unchecked until the QA user marks it. Plot, tree, and regen remarks are not tolerance items. Double-click a Crew value cell to open a larger read-only box for long database text. Plot, tree, and regen fields follow saved custom field order first. If no custom order is saved, they follow the database measurement query column order when available: GetPlotMeasurementsForPeriod, GetTreeMeasurementsForPeriodByKey, and GetRegenMeasurements. If Access cannot expose a query order, PlotHoot falls back to AppColumns order.

Tree Data
All field-collected trees for the loaded plot are loaded automatically. Use the tree dropdown, Previous tree, and Next tree to move through the plot. Leave Show selected tree only checked for a clean one-tree-at-a-time workflow on large plots.

Use Crew-recorded tree found by QA to confirm that a tree already recorded by the crew was found by the QA cruiser. If the QA cruiser finds a tally tree the crew missed, check Crew missed tree(s) found by QA and open the Missed Trees tab. Add one row per missed tree with DBH, distance from the tree center to plot center, azimuth from the tree to plot center, and notes. Any crew-missed tree is an automatic plot failure.

Regen Data
All regen records for the loaded plot are loaded automatically and shown together by default. Use the regen dropdown, Previous regen, and Next regen to jump through the list. Check Show selected regen only when you want to focus on one regen record.

Location / Execution
The Location / Execution tab captures Table D from the check form, excluding Plot Tally Sheets Neatness/Legibility. These rows use the GoodFairPoor tolerance in Tolerance Setup. Good is always 0 point loss. Fair and Poor use the point loss written in the Tolerance Setup Value cell, such as Good=0; Fair=1; Poor=2. If Fair should still pass for an item, set Fair=0. A Poor rating becomes an automatic plot failure when Critical fail is checked for that row. GPS defaults to Good=0; Fair=0; Poor=0 with Critical fail checked, so only Poor forces the plot to fail.

Sorting Columns
Tap any grid column header to sort A-Z or small-to-large. Tap the same header again to reverse the sort.

Tooltips
Hover over a control, or press-and-hold with touch, to see short explanations for settings, scoring, and export fields.

Tablet Keyboard
On Windows tablets, numeric QA fields, species/code fields, tolerance values, points, and max-loss boxes request the number keyboard. Text, notes, remarks, and pass/fail-style fields request the normal text keyboard.

Auto-Advance
Tap a typed QA value cell once to start editing. After entering a QA value, press Enter, Next, or Tab on the keyboard to save that value and move straight down to the next QA value field, ready for typing. Checkbox checks are marked by tapping the checkbox. Tree and regen tabs advance to the next selected tree or regen record when the last visible field is complete.

Tolerance Setup
The Tolerance Setup tab controls the tolerance rule for each visible QA field. Field rows come from checked active AppColumns fields when AppColumns is available. PlotHoot skips office/system fields such as IDs, GMP, Calc fields, per-acre expansion, FLC commercial, management unit, and plot/tree/regen remarks. If AppColumns is not available, PlotHoot falls back to readable database fields.

Use the Show filter to view only Plot, Tree, Regen, or Location / Execution rules. Uncheck Use to remove a field from the loaded QA rows for this setup, then click Save setup.

To customize field order, drag the Move handle at the far left of a row on Plot Data, Tree Data, or Regen Data. You can also drag the Field cell. Moving a tree or regen field applies that field order to every loaded tree or regen entry. PlotHoot saves the new order automatically for future loaded plots. Reset defaults clears the saved custom order and returns to the database query/AppColumns order.

Exact: QA value must match the crew value exactly. The Value cell is disabled because no tolerance amount is needed.
Range: fixed plus/minus range. Example: DBH 0.20, UTM 30 ft, or elevation 100.
Percent: height-style fields use relative percent around the crew value. Fields already recorded as percentages or ratios, such as Slope Percent and Crown Ratio, use plus/minus percentage points.
Class: allowed class difference. Example: crown class within 1 class.
Use Stem Count Percentage For Scoring: project-wide regen stem count switch. When checked, regen stem count rows use percentage-basis scoring and the Regen StemCount table turns off. When unchecked, regen stem count rows use the editable Regen StemCount table.
StemCount: uses the editable Regen stem count tolerances table in Tolerance Setup. The QA cruiser count is presumed correct and chooses the count range. Tolerance cells are blank by default; blank means exact match is required until you enter an allowed plus/minus number.
StemPercent: uses the percentage-basis regen method when Use Stem Count Percentage For Scoring is checked. The default Value is Allowance=10; Step=5, meaning the first 10 percent error is allowed, then each started 5 percent above the allowance adds the row Points as point loss. Example: crew count 5 and QA count 7 is about 29 percent error; 29 - 10 = 19; 19 / 5 rounds up to 4 steps; with Points = 2, point loss is 8.
PassFail: treats pass/fail, yes/no, true/false style values as equivalent.
Filled: checkbox check that the crew/database value was supplied. Filled checkboxes start unchecked. Plot number, crew, measurement day/date, and tree number default to this.
GoodFairPoor: Location / Execution ratings. Use Good=0; Fair=1; Poor=2, or Fair=0 when Fair should pass.

Tree IDBH values stored as whole-number tenths are compared as inches. For example, crew 60 and QA 62 are treated as 6.0 and 6.2 inches, so they pass with Range 0.20. Regen IDBH defaults to Exact.

Already-percent fields are compared by percentage points. For example, Slope Percent crew 40 and QA 45 pass with Percent 5.

Example
To allow elevation within plus/minus 100 feet:
Tolerance type = Range
Value = 100

Scoring
Most field Points start from AppColumns.QApoints. You can edit Points in Tolerance Setup or directly in a loaded Plot, Tree, or Regen QA table for the current QA work. Location / Execution Points come from the GoodFairPoor Value, such as Good=0; Fair=1; Poor=2. A failed non-critical row adds its points to total point loss. A failed critical row forces the overall plot status to Fail.

Fail cutoff is optional. It is the total point-loss cutoff for the whole plot. When it is blank, PlotHoot does not use a total point-loss cutoff; critical failures and any section maxes you enter can still fail the plot. If you enter 25, total error greater than 25 fails the plot. Older notes may call this loss cutoff.

Plot max loss, Tree max loss, Regen max loss, and Table D max loss are optional section caps. Leave a section max blank to let PlotHoot use the loaded row point sum for that section.

Regen max loss is optional whether you use the stem-count table or the stem-count percentage method. The percentage checkbox only changes how regen stem count row errors are calculated; it does not make Regen max loss required.

When Tree max loss is filled in, it is multiplied by the number of loaded trees on the plot. Example: 22 max points per tree and 2 trees gives a Tree total max loss of 44; more than 44 tree point loss fails the plot.

The Summary of Check Cruise total max point loss is the section max sum. Blank section maxes show the loaded row point sum, and a blank failure threshold shows Not set.

Crew-recorded tree found by QA is added automatically to each loaded crew tree. It uses PassFail and Critical fail by default. Enter Pass/Yes when the crew-recorded tree is found, or Fail/No when that crew-recorded tree is missing.

BIA Inventory QA Progress
The BIA inventory QA progress tracker counts saved plots against all plots in the connected inventory. PlotHoot uses a 10 percent target and rounds up to the next whole plot.

Export / Import
Save setup stores tolerance, scoring, Use-field choices, stem-count scoring choice, stem table values, and moved field order locally in Data\PlotHootSettings.json beside the app when possible. If Windows blocks writing beside the app, PlotHoot uses the current user's local app data folder.

Export QA CSV saves both QA rows and setup/tolerance/scoring/Use-field/field-order rows. Rows marked QA are field checks. Rows marked Setting preserve the setup, so importing that CSV can restore the same tolerance, scoring, removed-field choices, and moved field order setup later. QA rows also carry FieldOrder as a fallback. The CSV includes ProjectName, check cruiser, and check cruise date.

Export QA Results saves an Excel .xlsx file with real tabs: Summary, Passed Plots, Failed Plots, Audited Plot Data, Audited Tree Data, Audited Regen Data, Audited Location Data, and Setup Settings. Use this when you want the QA information split into workbook tabs.

Failed CSV is a compact failed-plot list with project name, check cruiser, check cruise date, BIA inventory QA progress, remarks, and UTM values.
QA Report HTML is an offline formatted report that opens on a tablet or desktop. It starts with a summary page, shows total trees audited and errors, lists passed plots and failed plots with UTMs and failure reasons, and includes filterable tabs for audited plot, tree, regen, and location data.

Overall Status
Not checked: no QA values have been entered.
Incomplete: some QA values are entered, but not all required loaded check rows are complete.
Pass: all loaded check rows have QA values, no critical failures occurred, no section is over its max, and total error is within any filled-in failure threshold.
Fail: a critical row failed, a section is over its max, or all rows are filled and total error is greater than a filled-in failure threshold.

The Overall QA status line shows total error, the optional fail cutoff, failed check rows, and unchecked rows. Failed check rows are individual field/check rows with Fail status, not failed plots.
Save QA stores the current plot check even when it is incomplete. The check cruiser name, check cruise date, and QA remarks are saved per plot and included in exports. To finish a plot, fill every required QA row on Plot Data, Tree Data, Regen Data, and Location / Execution. Once there are no unchecked rows, the status changes to Pass or Fail; PlotHoot does not use a separate Complete status.

Safety
The selected database is read-only to PlotHoot. QA work is saved only through CSV or HTML exports. No internet connection or admin rights are required.

Tree Section Field Selection
Plot Number is intentionally skipped in Tree Data and Regen Data. Real DBH is intentionally skipped in Tree Data. Management Unit and plot/tree/regen remarks are intentionally skipped in tolerance checks.

Created by Christopher LaCroix.

Developed by BIA Division of Forestry, Branch of Inventory and planning.
"@.Trim()
}

function New-HooterUserGuideHtml {
    $guideText = Escape-Html (Get-HooterInAppGuideText)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!doctype html>")
    [void]$sb.AppendLine("<html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>")
    [void]$sb.AppendLine("<title>PlotHoot User Guide</title>")
    [void]$sb.AppendLine("<style>")
    [void]$sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f7f4;color:#162b28}header{background:#1e4744;color:white;padding:22px 28px}main{max-width:900px;margin:0 auto;padding:24px;white-space:pre-wrap;line-height:1.45}.brand{font-size:26px;font-weight:700}.footer{font-size:12px;color:#52615d;margin-top:24px}@media print{body{background:white}main{max-width:none;padding:0}header{color:#162b28;background:white;border-bottom:2px solid #1e4744}}")
    [void]$sb.AppendLine("</style></head><body>")
    [void]$sb.AppendLine("<header><div class='brand'>PlotHoot User Guide</div></header>")
    [void]$sb.AppendLine("<main>$guideText</main>")
    [void]$sb.AppendLine("</body></html>")
    return $sb.ToString()
}

function Show-HooterMainForm {
    Add-HooterAssemblies
    $splash = Show-HooterSplashScreen
    try {
    [void](Repair-HooterLaunchShortcuts -Quiet)
    Set-HooterAppIdentity
    Load-HooterSettings

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$script:AppName $script:AppVersion"
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(720, 720)
    $form.Size = New-Object System.Drawing.Size(1280, 830)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    $form.Font = New-HooterFont 9.2
    if (Test-Path -LiteralPath $script:AppIconPath) {
        try { $form.Icon = New-Object System.Drawing.Icon($script:AppIconPath) } catch {}
    }
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 18000
    $toolTip.InitialDelay = 350
    $toolTip.ReshowDelay = 120
    $toolTip.ShowAlways = $true
    $setToolTip = {
        param([System.Windows.Forms.Control]$Control, [string]$Text)
        if ($null -ne $Control -and -not [string]::IsNullOrWhiteSpace($Text)) {
            $toolTip.SetToolTip($Control, $Text)
        }
    }.GetNewClosure()
    $setGridColumnTip = {
        param([System.Windows.Forms.DataGridView]$Grid, [string]$ColumnName, [string]$Text)
        if ($null -ne $Grid -and $Grid.Columns.Contains($ColumnName)) {
            $Grid.Columns[$ColumnName].ToolTipText = $Text
            $Grid.Columns[$ColumnName].HeaderCell.ToolTipText = $Text
        }
    }

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = "Top"
    $header.Height = 104
    $header.BackColor = [System.Drawing.Color]::FromArgb(30, 71, 68)
    $form.Controls.Add($header)

    if (Test-Path -LiteralPath $script:OwlImagePath) {
        $picture = New-Object System.Windows.Forms.PictureBox
        $picture.Location = New-Object System.Drawing.Point(18, 16)
        $picture.Size = New-Object System.Drawing.Size(70, 70)
        $picture.SizeMode = "Zoom"
        $picture.Image = [System.Drawing.Image]::FromFile($script:OwlImagePath)
        $header.Controls.Add($picture)
        & $setToolTip $picture "PlotHoot offline QA field-check app."
    }

    $title = New-HooterLabel "PlotHoot" 104 18 260 34 18 ([System.Drawing.FontStyle]::Bold)
    $title.ForeColor = [System.Drawing.Color]::White
    $header.Controls.Add($title)
    $subtitle = New-HooterLabel "Offline field QA for plot, tree, and regen checks" 106 54 460 24 9.6
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(214, 231, 223)
    $header.Controls.Add($subtitle)
    $projectLabel = New-HooterLabel "Project: Not connected" 106 76 460 22 9.0
    $projectLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 231, 223)
    $header.Controls.Add($projectLabel)
    & $setToolTip $title "PlotHoot reads the selected Access database and helps compare crew values against QA cruiser values."
    & $setToolTip $subtitle "Offline QA for plot, tree, regen, missed-tree, and location/execution checks."
    & $setToolTip $projectLabel "Project name read from the connected database. PlotHoot cleans trailing date/version text from filename fallbacks."

    $dbLabel = New-HooterLabel "Master database" 590 16 120 22 9.0
    $dbLabel.ForeColor = [System.Drawing.Color]::White
    $header.Controls.Add($dbLabel)
    $databaseBox = New-HooterTextBox 710 13 240 (Get-HooterDatabaseDisplayName -Path $Database)
    $databaseBox.Anchor = "Top,Left"
    $databaseBox.Tag = ConvertTo-HooterText $Database
    $header.Controls.Add($databaseBox)
    $browseDbButton = New-HooterButton "Browse" 990 12 82 30
    $header.Controls.Add($browseDbButton)
    $connectButton = New-HooterButton "Load Completed Plots" 858 53 158 30
    $connectButton.Font = New-HooterFont 9.2 ([System.Drawing.FontStyle]::Bold)
    $connectButton.BackColor = [System.Drawing.Color]::FromArgb(255, 224, 130)
    $connectButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(120, 76, 20)
    $connectButton.ForeColor = [System.Drawing.Color]::FromArgb(53, 35, 10)
    $header.Controls.Add($connectButton)
    $loadAllPlotsButton = New-HooterButton "Load All Plots" 1024 53 112 30
    $header.Controls.Add($loadAllPlotsButton)
    & $setToolTip $databaseBox "Selected master or project Access database. PlotHoot shows the file name here but remembers the full path internally."
    & $setToolTip $dbLabel "Database currently selected for read-only QA checks."
    & $setToolTip $browseDbButton "Choose the Access database file."
    & $setToolTip $connectButton "Recommended: load only completed plots where InventoryAssignmentPlots StatusID = 4 into the plot dropdown."
    & $setToolTip $loadAllPlotsButton "Load every plot into the dropdown. Use this only when you need to check a plot that is not marked completed yet."

    $plotLabel = New-HooterLabel "Plot #" 590 58 54 22 9.0
    $plotLabel.ForeColor = [System.Drawing.Color]::White
    $header.Controls.Add($plotLabel)
    $plotBox = New-Object System.Windows.Forms.ComboBox
    $plotBox.Location = New-Object System.Drawing.Point(650, 55)
    $plotBox.Size = New-Object System.Drawing.Size(200, 28)
    $plotBox.Font = New-HooterFont 9.4
    $plotBox.DropDownStyle = "DropDown"
    $plotBox.AutoCompleteMode = "SuggestAppend"
    $plotBox.AutoCompleteSource = "CustomSource"
    $plotBox.IntegralHeight = $false
    $plotBox.MaxDropDownItems = 18
    $plotBox.DropDownHeight = 360
    $header.Controls.Add($plotBox)
    $loadPlotButton = New-HooterButton "Load plot" 1144 53 92 30
    $header.Controls.Add($loadPlotButton)
    & $setToolTip $plotLabel "Typed or selected plot number to load for QA."
    & $setToolTip $plotBox "Type a plot number directly, then press Enter or tap Load plot. The dropdown normally lists completed plots only; tap Load All Plots if needed."
    & $setToolTip $loadPlotButton "Load plot, tree, regen, and Table D checks for the typed or selected plot number."

    $statusLabel = New-HooterLabel "Choose a database, connect, then load a plot." 18 110 900 26 9.2
    $statusLabel.Anchor = "Top,Left,Right"
    $statusLabel.AutoEllipsis = $false
    $statusLabel.UseCompatibleTextRendering = $true
    $form.Controls.Add($statusLabel)
    & $setToolTip $statusLabel "Current PlotHoot status, including connection, loaded plot, save, and export messages."

    $autoSaveNoticePanel = New-Object System.Windows.Forms.Panel
    $autoSaveNoticePanel.Location = New-Object System.Drawing.Point(1036, 108)
    $autoSaveNoticePanel.Size = New-Object System.Drawing.Size(190, 30)
    $autoSaveNoticePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 246, 214)
    $autoSaveNoticePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $autoSaveNoticePanel.Visible = $false
    $autoSaveNoticePanel.Anchor = "Top,Right"
    $autoSaveNoticeLabel = New-HooterLabel "Auto saving..." 10 5 168 20 9.0 ([System.Drawing.FontStyle]::Bold)
    $autoSaveNoticeLabel.ForeColor = [System.Drawing.Color]::FromArgb(86, 57, 6)
    $autoSaveNoticePanel.Controls.Add($autoSaveNoticeLabel)
    $form.Controls.Add($autoSaveNoticePanel)
    & $setToolTip $autoSaveNoticePanel "Temporary message shown while PlotHoot is updating, saving, or refreshing the screen."
    & $setToolTip $autoSaveNoticeLabel "PlotHoot is working. This message clears automatically."

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(18, 140)
    $tabControl.Size = New-Object System.Drawing.Size(1226, 590)
    $tabControl.Anchor = "Top,Bottom,Left,Right"
    $tabControl.Font = New-HooterFont 9.4
    $tabControl.ShowToolTips = $true
    Register-HooterColoredTabHeaders -TabControl $tabControl
    $form.Controls.Add($tabControl)
    & $setToolTip $tabControl "Use the colored tabs to move between plot, tree, regen, scoring setup, review/export, and help."

    $plotPage = New-Object System.Windows.Forms.TabPage
    $plotPage.Text = "Plot Data"
    $plotPage.ToolTipText = "Plot-level database fields with crew values beside QA values."
    $plotPage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $plotPage -Color ([System.Drawing.Color]::FromArgb(242, 201, 76)) -TextColor ([System.Drawing.Color]::FromArgb(32, 32, 24))
    Set-HooterScrollableTab -Page $plotPage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($plotPage)
    $plotGrid = New-HooterGrid 12 16 1186 510
    $plotGrid.Anchor = "Top,Bottom,Left,Right"
    Enable-HooterVisibleGridScrollBars -Grid $plotGrid
    Add-HooterCheckColumns -Grid $plotGrid -WithEntry $false
    Register-HooterDataEntryGridDragReorder -Grid $plotGrid
    Enable-HooterGridTapSort -Grid $plotGrid -AfterSort ({ Refresh-HooterGridStatuses -Grid $plotGrid; Update-HooterOverallLabel }.GetNewClosure())
    $plotPage.Controls.Add($plotGrid)
    & $setToolTip $plotGrid "Enter QA values beside the crew values. Press Enter or Tab to move down to the next QA value."

    $treePage = New-Object System.Windows.Forms.TabPage
    $treePage.Text = "Tree Data"
    $treePage.ToolTipText = "All crew-recorded trees for the loaded plot, checked one tree at a time by default."
    $treePage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $treePage -Color ([System.Drawing.Color]::FromArgb(46, 125, 50))
    Set-HooterScrollableTab -Page $treePage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($treePage)
    $treeRecordBox = New-Object System.Windows.Forms.ComboBox
    $treeRecordBox.Location = New-Object System.Drawing.Point(16, 18)
    $treeRecordBox.Size = New-Object System.Drawing.Size(300, 28)
    $treeRecordBox.DropDownStyle = "DropDownList"
    $treeRecordBox.Font = New-HooterFont 9.2
    $treeRecordBox.IntegralHeight = $false
    $treeRecordBox.MaxDropDownItems = 18
    $treeRecordBox.DropDownHeight = 360
    $treePage.Controls.Add($treeRecordBox)
    $addTreeButton = New-HooterButton "Previous tree" 326 16 118 30
    $treePage.Controls.Add($addTreeButton)
    $removeTreeButton = New-HooterButton "Next tree" 452 16 96 30
    $treePage.Controls.Add($removeTreeButton)
    $treeSelectedOnlyCheck = New-Object System.Windows.Forms.CheckBox
    $treeSelectedOnlyCheck.Text = "Show selected tree only"
    $treeSelectedOnlyCheck.Location = New-Object System.Drawing.Point(565, 20)
    $treeSelectedOnlyCheck.Size = New-Object System.Drawing.Size(170, 24)
    $treeSelectedOnlyCheck.Checked = $true
    $treeSelectedOnlyCheck.Font = New-HooterFont 9.0
    $treeSelectedOnlyCheck.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 42)
    $treePage.Controls.Add($treeSelectedOnlyCheck)
    & $setToolTip $treeRecordBox "Choose which loaded tree to QA. All trees remain saved and exported even when only one is visible."
    & $setToolTip $addTreeButton "Move to the previous loaded tree."
    & $setToolTip $removeTreeButton "Move to the next loaded tree."
    & $setToolTip $treeSelectedOnlyCheck "Keep this checked for a one-tree-at-a-time workflow on large plots."
    $treeProgressLabel = New-HooterLabel "No tree records loaded." 750 20 430 24 9.0
    $treeProgressLabel.Anchor = "Top,Left,Right"
    $treePage.Controls.Add($treeProgressLabel)
    & $setToolTip $treeProgressLabel "Progress for the selected loaded tree and total tree-check status."
    $crewMissedTreeCheck = New-Object System.Windows.Forms.CheckBox
    $crewMissedTreeCheck.Text = "Crew missed tree(s) found by QA - automatic plot fail"
    $crewMissedTreeCheck.Location = New-Object System.Drawing.Point(16, 50)
    $crewMissedTreeCheck.Size = New-Object System.Drawing.Size(380, 24)
    $crewMissedTreeCheck.Checked = $false
    $crewMissedTreeCheck.Font = New-HooterFont 9.0 ([System.Drawing.FontStyle]::Bold)
    $crewMissedTreeCheck.ForeColor = [System.Drawing.Color]::FromArgb(130, 36, 24)
    $treePage.Controls.Add($crewMissedTreeCheck)
    & $setToolTip $crewMissedTreeCheck "Check this when the QA cruiser finds a tally tree the crew missed. This forces the plot to fail. Use the Missed Trees tab to log DBH, distance, azimuth, and notes."
    $treeGrid = New-HooterGrid 12 82 1186 444
    $treeGrid.Anchor = "Top,Bottom,Left,Right"
    Enable-HooterVisibleGridScrollBars -Grid $treeGrid
    Add-HooterCheckColumns -Grid $treeGrid -WithEntry $true
    if ($treeGrid.Columns.Contains("EntryNumber")) { $treeGrid.Columns["EntryNumber"].Visible = $false }
    if ($treeGrid.Columns.Contains("CrewRecord")) {
        $treeGrid.Columns["CrewRecord"].HeaderText = "Tree record"
        $treeGrid.Columns["CrewRecord"].Width = 190
        $treeGrid.Columns["CrewRecord"].ToolTipText = "Crew/database tree record that this QA row is checking."
        $treeGrid.Columns["CrewRecord"].HeaderCell.ToolTipText = "Crew/database tree record that this QA row is checking."
    }
    Register-HooterDataEntryGridDragReorder -Grid $treeGrid
    Enable-HooterGridTapSort -Grid $treeGrid -AfterSort ({ Set-HooterTreeVisibleRows; Refresh-HooterGridStatuses -Grid $treeGrid; Update-HooterOverallLabel }.GetNewClosure())
    $treePage.Controls.Add($treeGrid)
    & $setToolTip $treeGrid "Enter QA values for the selected crew tree. Use Previous/Next tree for large plots."

    $missedTreePage = New-Object System.Windows.Forms.TabPage
    $missedTreePage.Text = "Missed Trees"
    $missedTreePage.ToolTipText = "Trees found by QA that were missed by the crew; these force an automatic plot failure."
    $missedTreePage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $missedTreePage -Color ([System.Drawing.Color]::FromArgb(198, 40, 40))
    Set-HooterScrollableTab -Page $missedTreePage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($missedTreePage)
    $missedTreeHint = New-HooterLabel "Log trees found by the QA cruiser that were missed by the crew. Any missed tree is an automatic plot failure." 16 18 980 24 9.2
    $missedTreeHint.Anchor = "Top,Left,Right"
    $missedTreePage.Controls.Add($missedTreeHint)
    $addMissedTreeButton = New-HooterButton "Add missed tree" 16 50 130 30
    $missedTreePage.Controls.Add($addMissedTreeButton)
    $removeMissedTreeButton = New-HooterButton "Remove selected" 156 50 130 30
    $missedTreePage.Controls.Add($removeMissedTreeButton)
    $missedTreeStatusLabel = New-HooterLabel "No crew-missed trees logged." 304 55 780 22 9.0 ([System.Drawing.FontStyle]::Bold)
    $missedTreeStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 75, 71)
    $missedTreeStatusLabel.Anchor = "Top,Left,Right"
    $missedTreePage.Controls.Add($missedTreeStatusLabel)
    & $setToolTip $missedTreeHint "These are extra tally trees found by QA that do not exist in the crew/database tree list."
    & $setToolTip $addMissedTreeButton "Add a missed tree row for DBH, distance, azimuth, and notes."
    & $setToolTip $removeMissedTreeButton "Remove the selected missed tree row."
    $missedTreeGrid = New-HooterGrid 12 90 1186 436
    $missedTreeGrid.Anchor = "Top,Bottom,Left,Right"
    Enable-HooterVisibleGridScrollBars -Grid $missedTreeGrid
    Add-HooterMissedTreeColumns -Grid $missedTreeGrid
    Enable-HooterGridTapSort -Grid $missedTreeGrid -AfterSort ({ Show-HooterAutoSaveNotice -Text "Updating missed tree check..."; Request-HooterMissedTreeRefresh }.GetNewClosure())
    $missedTreePage.Controls.Add($missedTreeGrid)
    & $setToolTip $missedTreeStatusLabel "Shows whether missed-tree rows have been logged and whether they force a critical failure."
    & $setToolTip $missedTreeGrid "Add one row for each tree found by QA that was missed by the crew."

    $regenPage = New-Object System.Windows.Forms.TabPage
    $regenPage.Text = "Regen Data"
    $regenPage.ToolTipText = "All regen records loaded for the plot. QA values are treated as the correct count/value."
    $regenPage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $regenPage -Color ([System.Drawing.Color]::FromArgb(30, 136, 229))
    Set-HooterScrollableTab -Page $regenPage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($regenPage)
    $regenRecordBox = New-Object System.Windows.Forms.ComboBox
    $regenRecordBox.Location = New-Object System.Drawing.Point(16, 18)
    $regenRecordBox.Size = New-Object System.Drawing.Size(340, 28)
    $regenRecordBox.DropDownStyle = "DropDownList"
    $regenRecordBox.Font = New-HooterFont 9.2
    $regenRecordBox.IntegralHeight = $false
    $regenRecordBox.MaxDropDownItems = 18
    $regenRecordBox.DropDownHeight = 360
    $regenPage.Controls.Add($regenRecordBox)
    $addRegenButton = New-HooterButton "Previous regen" 366 16 124 30
    $regenPage.Controls.Add($addRegenButton)
    $removeRegenButton = New-HooterButton "Next regen" 500 16 104 30
    $regenPage.Controls.Add($removeRegenButton)
    $regenSelectedOnlyCheck = New-Object System.Windows.Forms.CheckBox
    $regenSelectedOnlyCheck.Text = "Show selected regen only"
    $regenSelectedOnlyCheck.Location = New-Object System.Drawing.Point(620, 20)
    $regenSelectedOnlyCheck.Size = New-Object System.Drawing.Size(180, 24)
    $regenSelectedOnlyCheck.Checked = $false
    $regenSelectedOnlyCheck.Font = New-HooterFont 9.0
    $regenSelectedOnlyCheck.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 42)
    $regenPage.Controls.Add($regenSelectedOnlyCheck)
    & $setToolTip $regenRecordBox "Jump to a loaded regen record. Regen records show together by default."
    & $setToolTip $addRegenButton "Move to the previous loaded regen record."
    & $setToolTip $removeRegenButton "Move to the next loaded regen record."
    & $setToolTip $regenSelectedOnlyCheck "Optional focus filter. Leave unchecked to show every regen record on the loaded plot at once."
    $regenProgressLabel = New-HooterLabel "No regen records loaded." 815 20 370 24 9.0
    $regenProgressLabel.Anchor = "Top,Left,Right"
    $regenPage.Controls.Add($regenProgressLabel)
    & $setToolTip $regenProgressLabel "Progress for loaded regen checks on the current plot."
    $regenGrid = New-HooterGrid 12 58 1186 468
    $regenGrid.Anchor = "Top,Bottom,Left,Right"
    Enable-HooterVisibleGridScrollBars -Grid $regenGrid
    Add-HooterCheckColumns -Grid $regenGrid -WithEntry $true
    if ($regenGrid.Columns.Contains("EntryNumber")) { $regenGrid.Columns["EntryNumber"].Width = 64 }
    if ($regenGrid.Columns.Contains("CrewRecord")) {
        $regenGrid.Columns["CrewRecord"].HeaderText = "Regen record"
        $regenGrid.Columns["CrewRecord"].Width = 190
        $regenGrid.Columns["CrewRecord"].ToolTipText = "Crew/database regen record that this QA row is checking."
        $regenGrid.Columns["CrewRecord"].HeaderCell.ToolTipText = "Crew/database regen record that this QA row is checking."
    }
    Set-HooterGridColumnToolTip -Grid $regenGrid -ColumnName "QaValue" -Text "Enter the QA cruiser value here. For regen checks, the QA cruiser value is treated as the correct answer."
    Register-HooterDataEntryGridDragReorder -Grid $regenGrid
    Enable-HooterGridTapSort -Grid $regenGrid -AfterSort ({ Set-HooterRegenVisibleRows; Refresh-HooterGridStatuses -Grid $regenGrid; Update-HooterOverallLabel }.GetNewClosure())
    $regenPage.Controls.Add($regenGrid)
    & $setToolTip $regenGrid "Enter QA cruiser regen values. For regen checks, the QA value is presumed correct."

    $executionPage = New-Object System.Windows.Forms.TabPage
    $executionPage.Text = "Location / Execution"
    $executionPage.ToolTipText = "Table D location and execution checks with Good/Fair/Poor ratings."
    $executionPage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $executionPage -Color ([System.Drawing.Color]::FromArgb(239, 138, 36)) -TextColor ([System.Drawing.Color]::FromArgb(34, 28, 20))
    Set-HooterScrollableTab -Page $executionPage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($executionPage)
    $executionHint = New-HooterLabel "Table D checklist: Good/Fair/Poor point loss is editable in Tolerance Setup. GPS defaults to Poor = critical plot failure." 16 18 980 24 9.2
    $executionHint.Anchor = "Top,Left,Right"
    $executionPage.Controls.Add($executionHint)
    & $setToolTip $executionHint "Settings uses GoodFairPoor point loss values such as Good=0; Fair=1; Poor=2. Set Fair=0 when Fair should pass."
    $executionGrid = New-HooterGrid 12 58 1186 468
    $executionGrid.Anchor = "Top,Bottom,Left,Right"
    Enable-HooterVisibleGridScrollBars -Grid $executionGrid
    Add-HooterExecutionColumns -Grid $executionGrid
    Enable-HooterGridTapSort -Grid $executionGrid -AfterSort ({ Refresh-HooterExecutionGridStatuses; Update-HooterOverallLabel }.GetNewClosure())
    $executionPage.Controls.Add($executionGrid)
    & $setToolTip $executionGrid "Rate each location/execution item as Good, Fair, or Poor. Point loss is controlled in Tolerance Setup."

    $settingsPage = New-Object System.Windows.Forms.TabPage
    $settingsPage.Text = "Tolerance Setup"
    $settingsPage.ToolTipText = "Tolerance rules, point values, field use, stem-count settings, and max-loss thresholds."
    $settingsPage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $settingsPage -Color ([System.Drawing.Color]::FromArgb(0, 105, 92))
    Set-HooterScrollableTab -Page $settingsPage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($settingsPage)
    $saveSettingsButton = New-HooterButton "Save setup" 16 16 120 30
    $settingsPage.Controls.Add($saveSettingsButton)
    $defaultSettingsButton = New-HooterButton "Reset defaults" 146 16 120 30
    $settingsPage.Controls.Add($defaultSettingsButton)
    $failThresholdLabel = New-HooterLabel "Fail cutoff" 276 21 88 22 9.0
    $settingsPage.Controls.Add($failThresholdLabel)
    $maxPointLossBox = New-HooterTextBox 366 17 54 $script:MaxPointLoss
    $settingsPage.Controls.Add($maxPointLossBox)
    $plotMaxLossLabel = New-HooterLabel "Plot max loss" 424 21 92 22 9.0
    $settingsPage.Controls.Add($plotMaxLossLabel)
    $plotPointTotalBox = New-HooterTextBox 516 17 54 $script:PlotPointTotal
    $settingsPage.Controls.Add($plotPointTotalBox)
    $treeMaxLossLabel = New-HooterLabel "Tree max loss" 588 21 92 22 9.0
    $settingsPage.Controls.Add($treeMaxLossLabel)
    $treePointTotalBox = New-HooterTextBox 680 17 54 $script:TreePointTotal
    $settingsPage.Controls.Add($treePointTotalBox)
    $treeMaxSummaryLabel = New-HooterLabel "Tree Count: 0    Tree total max loss: 0" 588 46 600 22 9.0 ([System.Drawing.FontStyle]::Bold)
    $treeMaxSummaryLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 75, 71)
    $settingsPage.Controls.Add($treeMaxSummaryLabel)
    $regenMaxLossLabel = New-HooterLabel "Regen max loss" 752 21 106 22 9.0
    $settingsPage.Controls.Add($regenMaxLossLabel)
    $regenPointTotalBox = New-HooterTextBox 858 17 54 $script:RegenPointTotal
    $settingsPage.Controls.Add($regenPointTotalBox)
    $executionMaxLossLabel = New-HooterLabel "Table D max loss" 918 21 116 22 9.0
    $settingsPage.Controls.Add($executionMaxLossLabel)
    $executionPointTotalBox = New-HooterTextBox 1034 17 54 $script:ExecutionPointTotal
    $settingsPage.Controls.Add($executionPointTotalBox)
    $moveFieldUpButton = New-HooterButton "Up" 1098 16 44 30
    $moveFieldUpButton.Visible = $false
    $settingsPage.Controls.Add($moveFieldUpButton)
    $moveFieldDownButton = New-HooterButton "Down" 1148 16 54 30
    $moveFieldDownButton.Visible = $false
    $settingsPage.Controls.Add($moveFieldDownButton)
    & $setToolTip $saveSettingsButton "Save tolerance, critical fail, max loss, and field order setup on this tablet or desktop."
    & $setToolTip $defaultSettingsButton "Reset visible tolerance and scoring settings, and clear saved custom field order so fields return to database query/AppColumns order."
    & $setToolTip $failThresholdLabel "Optional total point-loss cutoff. If total error is greater than this number, the plot fails. Leave blank when you only want critical failures and section maxes to fail the plot."
    & $setToolTip $maxPointLossBox "Optional overall failure threshold. Example: 25 means a total point loss of 26 or more fails the plot. Blank means no total loss threshold."
    & $setToolTip $plotMaxLossLabel "Maximum point loss for Table A / plot classification checks."
    & $setToolTip $plotPointTotalBox "Optional max point loss assigned to plot classification checks. Leave blank to calculate from loaded plot rows."
    & $setToolTip $treeMaxLossLabel "Maximum allowed Tree Classification point loss per loaded tree."
    & $setToolTip $treePointTotalBox "Optional per-tree max point loss. PlotHoot multiplies this by loaded tree count. Leave blank to calculate from loaded tree rows."
    & $setToolTip $treeMaxSummaryLabel "Calculated from Tree max loss multiplied by the number of loaded tree records. Example: 22 x 2 trees = 44 total tree max loss."
    & $setToolTip $regenMaxLossLabel "Maximum point loss for the regen / Table C section."
    & $setToolTip $regenPointTotalBox "Optional max point loss assigned to regen checks. Leave blank to calculate from loaded regen rows."
    & $setToolTip $executionMaxLossLabel "Maximum point loss for Table D location/execution rows. Good/Fair/Poor row loss is editable below."
    & $setToolTip $executionPointTotalBox "Optional Table D max loss. Leave blank to calculate from loaded Location / Execution rows."
    & $setToolTip $moveFieldUpButton "Move the selected Tolerance Setup row up in the setup list."
    & $setToolTip $moveFieldDownButton "Move the selected Tolerance Setup row down in the setup list."
    $useStemPercentCheck = New-Object System.Windows.Forms.CheckBox
    $useStemPercentCheck.Text = "Use Stem Count Percentage For Scoring"
    $useStemPercentCheck.Location = New-Object System.Drawing.Point(16, 76)
    $useStemPercentCheck.Size = New-Object System.Drawing.Size(318, 24)
    $useStemPercentCheck.Font = New-HooterFont 9.2
    $useStemPercentCheck.Checked = [bool]$script:UseStemCountPercentageForScoring
    $settingsPage.Controls.Add($useStemPercentCheck)
    $stemToleranceLabel = New-HooterLabel "Regen StemCount table" 16 108 260 22 9.2 ([System.Drawing.FontStyle]::Bold)
    $stemToleranceLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 75, 71)
    $settingsPage.Controls.Add($stemToleranceLabel)
    $stemToleranceGrid = New-HooterGrid 16 134 318 190
    $stemToleranceGrid.Anchor = "Top,Left"
    $stemToleranceGrid.MultiSelect = $false
    $stemToleranceGrid.AllowUserToResizeRows = $false
    $stemToleranceGrid.AllowUserToResizeColumns = $false
    $stemToleranceGrid.RowTemplate.Height = 26
    $stemToleranceGrid.ColumnHeadersHeight = 28
    $stemToleranceGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    [void](Add-GridTextColumn -Grid $stemToleranceGrid -Name "CountLabel" -Header "Count" -Width 180 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $stemToleranceGrid -Name "ToleranceValue" -Header "Tolerance" -Width 105)
    [void](Add-GridTextColumn -Grid $stemToleranceGrid -Name "StemBandKey" -Header "StemBandKey" -Width 80 -ReadOnly $true -Visible $false)
    Enable-HooterVisibleGridScrollBars -Grid $stemToleranceGrid
    Populate-HooterStemToleranceGrid -Grid $stemToleranceGrid
    $settingsPage.Controls.Add($stemToleranceGrid)
    & $setToolTip $useStemPercentCheck "Check this when the project scores regen stem counts by percent error. When checked, the stem-count table below is turned off."
    & $setToolTip $stemToleranceLabel "Sliding regen stem count tolerance table used only when Use Stem Count Percentage For Scoring is not checked."
    & $setToolTip $stemToleranceGrid "Edit the allowed plus/minus stem-count difference for each QA cruiser stem-count range. Blank means exact match is required. These values save and export with setup settings. Disabled when percentage scoring is checked."
    & $setGridColumnTip $stemToleranceGrid "CountLabel" "QA cruiser regen stem count range. The QA value entered on the Regen tab is presumed correct."
    & $setGridColumnTip $stemToleranceGrid "ToleranceValue" "Allowed plus/minus difference between the crew count and the QA cruiser count. Example: 20 means crew count may be within 20 stems of the QA count."

    $settingsFilterLabel = New-HooterLabel "Show" 352 48 42 22 9.0
    $settingsPage.Controls.Add($settingsFilterLabel)
    $settingsFilterBox = New-Object System.Windows.Forms.ComboBox
    $settingsFilterBox.Location = New-Object System.Drawing.Point(398, 44)
    $settingsFilterBox.Size = New-Object System.Drawing.Size(168, 28)
    $settingsFilterBox.Font = New-HooterFont 9.2
    $settingsFilterBox.DropDownStyle = "DropDownList"
    [void]$settingsFilterBox.Items.Add("All")
    [void]$settingsFilterBox.Items.Add("Plot")
    [void]$settingsFilterBox.Items.Add("Tree")
    [void]$settingsFilterBox.Items.Add("Regen")
    [void]$settingsFilterBox.Items.Add("Location / Execution")
    $settingsFilterBox.SelectedIndex = 0
    $settingsPage.Controls.Add($settingsFilterBox)
    & $setToolTip $settingsFilterLabel "Choose which tolerance section to display."
    & $setToolTip $settingsFilterBox "Filter Tolerance Setup to show only one section at a time."

    $settingsGrid = New-HooterGrid 352 78 846 448
    $settingsGrid.Anchor = "Top,Bottom,Left,Right"
    $useFieldColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $useFieldColumn.Name = "UseField"
    $useFieldColumn.HeaderText = "Use"
    $useFieldColumn.Width = 48
    $useFieldColumn.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    $useFieldColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    [void]$settingsGrid.Columns.Add($useFieldColumn)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "Group" -Header "Tab" -Width 80 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "FieldLabel" -Header "Field" -Width 250 -ReadOnly $true)
    $modeColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $modeColumn.Name = "Mode"
    $modeColumn.HeaderText = "Tolerance type"
    $modeColumn.Width = 150
    $modeColumn.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    [void]$modeColumn.Items.Add("Exact")
    [void]$modeColumn.Items.Add("Range")
    [void]$modeColumn.Items.Add("Percent")
    [void]$modeColumn.Items.Add("Class")
    [void]$modeColumn.Items.Add("StemCount")
    [void]$modeColumn.Items.Add("StemPercent")
    [void]$modeColumn.Items.Add("PassFail")
    [void]$modeColumn.Items.Add("Filled")
    [void]$modeColumn.Items.Add("GoodFairPoor")
    [void]$settingsGrid.Columns.Add($modeColumn)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "ToleranceValue" -Header "Value" -Width 120)
    $unitColumn = Add-GridTextColumn -Grid $settingsGrid -Name "UnitHint" -Header "Units" -Width 82 -ReadOnly $true
    $unitColumn.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(105, 112, 108)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "PointValue" -Header "Points" -Width 80 -ReadOnly $false)
    $criticalColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $criticalColumn.Name = "CriticalFail"
    $criticalColumn.HeaderText = "Critical fail"
    $criticalColumn.Width = 95
    $criticalColumn.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    $criticalColumn.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
    [void]$settingsGrid.Columns.Add($criticalColumn)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "FieldKey" -Header "FieldKey" -Width 80 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "TableName" -Header "Table" -Width 180 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "FieldName" -Header "Database field" -Width 220 -ReadOnly $true -Visible $false)
    [void](Add-GridTextColumn -Grid $settingsGrid -Name "FieldOrder" -Header "FieldOrder" -Width 80 -ReadOnly $true -Visible $false)
    Enable-HooterVisibleGridScrollBars -Grid $settingsGrid
    Enable-HooterGridTapSort -Grid $settingsGrid
    $settingsPage.Controls.Add($settingsGrid)
    & $setToolTip $settingsGrid "Set each field tolerance, point loss, and critical fail rule. Move field order by dragging Field cells on Plot Data, Tree Data, or Regen Data."
    & $setGridColumnTip $settingsGrid "UseField" "Uncheck to remove this field from loaded QA rows for this setup."
    & $setGridColumnTip $settingsGrid "Group" "Which PlotHoot tab or data section this field belongs to."
    & $setGridColumnTip $settingsGrid "FieldLabel" "Field shown to the QA cruiser."
    & $setGridColumnTip $settingsGrid "Mode" "Tolerance type: Exact, Range, Percent, Class, StemCount, StemPercent, PassFail, Filled, or GoodFairPoor."
    & $setGridColumnTip $settingsGrid "ToleranceValue" "Tolerance amount. Examples: 0.20 for DBH, 30 for UTM feet, 5 for height percent, 1 for one class, Allowance=10; Step=5 for StemPercent, or Good=0; Fair=1; Poor=2 for Location/Execution."
    & $setGridColumnTip $settingsGrid "UnitHint" "Suggested units for the Value cell."
    & $setGridColumnTip $settingsGrid "PointValue" "Point loss for this row. Most fields load from AppColumns.QApoints, but you can edit it here or directly in a loaded QA table."
    & $setGridColumnTip $settingsGrid "CriticalFail" "When checked, a failed value forces the whole plot to fail. For GoodFairPoor, this applies to Poor."
    & $setGridColumnTip $settingsGrid "TableName" "Source database table."
    & $setGridColumnTip $settingsGrid "FieldName" "Source database field."

    $reviewPage = New-Object System.Windows.Forms.TabPage
    $reviewPage.Text = "Review / Export"
    $reviewPage.ToolTipText = "Save the current plot QA record and export CSV, HTML, or Excel reports."
    $reviewPage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $reviewPage -Color ([System.Drawing.Color]::FromArgb(126, 87, 194))
    Set-HooterScrollableTab -Page $reviewPage -MinimumScrollHeight 650
    [void]$tabControl.TabPages.Add($reviewPage)
    $overallLabel = New-HooterLabel "Overall QA status: Not checked" 16 18 1040 28 11.0 ([System.Drawing.FontStyle]::Bold)
    $reviewPage.Controls.Add($overallLabel)
    $inventoryProgressLabel = New-HooterLabel "BIA inventory QA progress: connect to an inventory database to calculate the 10% plot target." 16 48 720 22 9.2 ([System.Drawing.FontStyle]::Bold)
    $reviewPage.Controls.Add($inventoryProgressLabel)
    $inventoryProgressBar = New-Object System.Windows.Forms.ProgressBar
    $inventoryProgressBar.Location = New-Object System.Drawing.Point(750, 50)
    $inventoryProgressBar.Size = New-Object System.Drawing.Size(448, 16)
    $inventoryProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $inventoryProgressBar.Minimum = 0
    $inventoryProgressBar.Maximum = 1000
    $inventoryProgressBar.Value = 0
    $reviewPage.Controls.Add($inventoryProgressBar)
    $checkCruiserLabel = New-HooterLabel "Check cruiser *" 16 72 120 22
    $reviewPage.Controls.Add($checkCruiserLabel)
    $checkCruiserBox = New-HooterTextBox 16 96 220 ""
    $reviewPage.Controls.Add($checkCruiserBox)
    $checkCruiseDateLabel = New-HooterLabel "Check date *" 252 72 100 22
    $reviewPage.Controls.Add($checkCruiseDateLabel)
    $checkCruiseDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $checkCruiseDatePicker.Location = New-Object System.Drawing.Point(252, 96)
    $checkCruiseDatePicker.Size = New-Object System.Drawing.Size(150, 28)
    $checkCruiseDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $checkCruiseDatePicker.ShowCheckBox = $true
    $checkCruiseDatePicker.Checked = $true
    $checkCruiseDatePicker.Font = New-HooterFont 9.4
    $reviewPage.Controls.Add($checkCruiseDatePicker)
    $qaRemarksLabel = New-HooterLabel "QA remarks" 16 130 100 22
    $reviewPage.Controls.Add($qaRemarksLabel)
    $notesBox = New-Object System.Windows.Forms.TextBox
    $notesBox.Location = New-Object System.Drawing.Point(16, 154)
    $notesBox.Size = New-Object System.Drawing.Size(470, 86)
    $notesBox.Multiline = $true
    $notesBox.ScrollBars = "Vertical"
    $notesBox.Font = New-HooterFont 9.4
    $reviewPage.Controls.Add($notesBox)
    $exportWorkbookButton = New-HooterButton "Export QA Results" 510 82 140 32
    $reviewPage.Controls.Add($exportWorkbookButton)
    $savePlotButton = New-HooterButton "Save QA" 656 82 76 32
    $reviewPage.Controls.Add($savePlotButton)
    $exportAllButton = New-HooterButton "Export QA CSV" 738 82 112 32
    $reviewPage.Controls.Add($exportAllButton)
    $importButton = New-HooterButton "Import QA CSV" 856 82 112 32
    $reviewPage.Controls.Add($importButton)
    $exportFailCsvButton = New-HooterButton "Failed CSV" 974 82 82 32
    $reviewPage.Controls.Add($exportFailCsvButton)
    $exportFailHtmlButton = New-HooterButton "QA Report HTML" 1062 82 136 32
    $reviewPage.Controls.Add($exportFailHtmlButton)
    & $setToolTip $overallLabel "Current pass/fail/incomplete status for the loaded plot QA check."
    & $setToolTip $inventoryProgressLabel "BIA policy target: save QA checks for at least 10% of all plots in the connected inventory."
    & $setToolTip $inventoryProgressBar "Progress toward the BIA 10% saved-plot QA target."
    & $setToolTip $checkCruiserLabel "Required field. The name is saved with this plot QA record."
    & $setToolTip $checkCruiserBox "Required before saving QA. Name of the check cruiser for the currently loaded plot. This is saved with the plot QA record and appears in exports."
    & $setToolTip $checkCruiseDateLabel "Required field. Defaults to today's date for this plot check."
    & $setToolTip $checkCruiseDatePicker "Required before saving QA. Defaults to today for the currently loaded plot and appears in exports."
    & $setToolTip $qaRemarksLabel "Optional remarks for this plot QA record."
    & $setToolTip $notesBox "Overall QA remarks for the selected plot. These remarks appear in saved QA and failed-plot exports."
    & $setToolTip $exportWorkbookButton "Export an Excel workbook with separate tabs for summary, passed plots, failed plots, audited plot/tree/regen/location data, and setup settings."
    & $setToolTip $savePlotButton "Save the current plot QA check in this app session."
    & $setToolTip $exportAllButton "Export all saved QA rows plus setup/tolerance/scoring settings."
    & $setToolTip $importButton "Import a PlotHoot QA CSV, including saved setup/tolerance/scoring settings."
    & $setToolTip $exportFailCsvButton "Export a compact CSV of failed plots, BIA inventory QA progress, remarks, and UTM values."
    & $setToolTip $exportFailHtmlButton "Export a formatted offline HTML QA report with summary, passed plots, failed plots, audited plot rows, tree rows, regen rows, and location rows."
    $summaryCheckLabel = New-HooterLabel "Summary of Check Cruise" 510 124 220 22 9.2 ([System.Drawing.FontStyle]::Bold)
    $reviewPage.Controls.Add($summaryCheckLabel)
    $errorSummaryGrid = New-HooterGrid 510 152 688 220
    $errorSummaryGrid.Anchor = "Top,Left,Right"
    $errorSummaryGrid.ReadOnly = $true
    $errorSummaryGrid.MultiSelect = $false
    $errorSummaryGrid.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $errorSummaryGrid.RowTemplate.Height = 28
    $errorSummaryGrid.ColumnHeadersHeight = 28
    $errorSummaryGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "Item" -Header "Item" -Width 430 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "TotalError" -Header "Total Error" -Width 100 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "MaxPointLoss" -Header "Max Point Loss" -Width 125 -ReadOnly $true)
    $reviewPage.Controls.Add($errorSummaryGrid)
    & $setToolTip $summaryCheckLabel "Summary totals for plot, tree, regen, and location/execution point loss."
    & $setToolTip $errorSummaryGrid "Current plot point-loss summary by check-form table."
    & $setGridColumnTip $errorSummaryGrid "Item" "Check-form table being summarized."
    & $setGridColumnTip $errorSummaryGrid "TotalError" "Point loss currently recorded for this section."
    & $setGridColumnTip $errorSummaryGrid "MaxPointLoss" "Configured max point loss for this section. The total row is the section max sum; the threshold row can be Not set."
    $reviewGrid = New-HooterGrid 12 390 1186 226
    $reviewGrid.Anchor = "Top,Bottom,Left,Right"
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "PlotNumber" -Header "Plot" -Width 95 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "OverallStatus" -Header "Status" -Width 100 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "Score" -Header "Error / fail rule" -Width 330 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "CriticalFailCount" -Header "Critical" -Width 70 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "FailedCount" -Header "Failed" -Width 70 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "UncheckedCount" -Header "Unchecked" -Width 85 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "UTM" -Header "UTM" -Width 150 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "CheckCruiserName" -Header "Check cruiser" -Width 135 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "CheckCruiseDate" -Header "Check date" -Width 95 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "Notes" -Header "QA remarks" -Width 170 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "SavedAt" -Header "Saved" -Width 140 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $reviewGrid -Name "SessionID" -Header "SessionID" -Width 80 -ReadOnly $true -Visible $false)
    Enable-HooterVisibleGridScrollBars -Grid $reviewGrid
    Enable-HooterGridTapSort -Grid $reviewGrid
    $reviewPage.Controls.Add($reviewGrid)
    & $setToolTip $reviewGrid "Saved plot QA sessions. Tap a column header to sort."
    & $setGridColumnTip $reviewGrid "PlotNumber" "Saved plot number."
    & $setGridColumnTip $reviewGrid "OverallStatus" "Overall saved QA result for this plot."
    & $setGridColumnTip $reviewGrid "Score" "Total point loss and configured max loss values."
    & $setGridColumnTip $reviewGrid "CriticalFailCount" "Number of critical failures in the saved QA session."
    & $setGridColumnTip $reviewGrid "FailedCount" "Number of failed check rows."
    & $setGridColumnTip $reviewGrid "UncheckedCount" "Number of required check rows that still need QA values. A plot remains Incomplete while this is greater than zero, unless a critical failure already forces Fail."
    & $setGridColumnTip $reviewGrid "UTM" "UTM values from the loaded plot, used for failed-plot navigation exports."
    & $setGridColumnTip $reviewGrid "CheckCruiserName" "Check cruiser saved for this plot."
    & $setGridColumnTip $reviewGrid "CheckCruiseDate" "Check cruise date saved for this plot."
    & $setGridColumnTip $reviewGrid "Notes" "QA remarks saved for this plot."

    $userGuidePage = New-Object System.Windows.Forms.TabPage
    $userGuidePage.Text = "User Guide"
    $userGuidePage.ToolTipText = "Built-in offline guide for PlotHoot workflow and scoring."
    $userGuidePage.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 244)
    Set-HooterTabHeaderColor -Page $userGuidePage -Color ([System.Drawing.Color]::FromArgb(96, 105, 116))
    Set-HooterScrollableTab -Page $userGuidePage -MinimumScrollHeight 700
    [void]$tabControl.TabPages.Add($userGuidePage)
    $openGuideButton = New-HooterButton "Open printable guide" 16 16 160 30
    $userGuidePage.Controls.Add($openGuideButton)
    $guideHintLabel = New-HooterLabel "This tab is offline. The printable guide opens as a local HTML file." 190 21 620 22 9.0
    $userGuidePage.Controls.Add($guideHintLabel)
    $guideBox = New-Object System.Windows.Forms.RichTextBox
    $guideBox.Location = New-Object System.Drawing.Point(16, 58)
    $guideBox.Size = New-Object System.Drawing.Size(1186, 560)
    $guideBox.Anchor = "Top,Bottom,Left,Right"
    $guideBox.ReadOnly = $true
    $guideBox.BorderStyle = "FixedSingle"
    $guideBox.BackColor = [System.Drawing.Color]::White
    $guideBox.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 42)
    $guideBox.Font = New-HooterFont 10.0
    $guideBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::ForcedVertical
    $guideBox.Text = Get-HooterInAppGuideText
    $userGuidePage.Controls.Add($guideBox)
    & $setToolTip $openGuideButton "Open the standalone HTML user guide for printing or easier reading."
    & $setToolTip $guideHintLabel "The guide is stored locally with PlotHoot and does not need internet."
    & $setToolTip $guideBox "Offline PlotHoot user guide. Use the scroll bar to read more."

    $footer = New-HooterLabel "Developed by BIA Division of Forestry, Branch of Inventory and planning." 18 742 820 24 8.8
    $footer.Anchor = "Bottom,Left"
    $footer.ForeColor = [System.Drawing.Color]::FromArgb(72, 87, 82)
    $form.Controls.Add($footer)
    & $setToolTip $footer "PlotHoot credit line shown on the app window."

    $applyResponsiveLayout = {
        try {
            $clientW = [Math]::Max(720, $form.ClientSize.Width)
            $clientH = [Math]::Max(700, $form.ClientSize.Height)
            $compact = ($form.ClientSize.Height -gt $form.ClientSize.Width) -or ($form.ClientSize.Width -lt 1240)
            $outerMargin = if ($compact) { 10 } else { 18 }

            if ($compact) {
                $header.Height = 168
                Set-HooterControlBounds -Control $picture -X 18 -Y 16 -W 58 -H 58
                Set-HooterControlBounds -Control $title -X 88 -Y 10 -W ([Math]::Max(220, $clientW - 104)) -H 30
                Set-HooterControlBounds -Control $subtitle -X 88 -Y 40 -W ([Math]::Max(220, $clientW - 104)) -H 22
                Set-HooterControlBounds -Control $projectLabel -X 88 -Y 64 -W ([Math]::Max(220, $clientW - 104)) -H 22

                $rightEdge = $clientW - $outerMargin
                $browseButtonW = 82
                $completedButtonW = 158
                $allPlotsButtonW = 112
                $buttonGap = 8
                $dbButtonGap = 28
                $allPlotsX = $rightEdge - $allPlotsButtonW
                $connectX = $allPlotsX - $buttonGap - $completedButtonW
                $browseX = $connectX - $buttonGap - $browseButtonW
                $dbX = 136
                $dbAvailableW = $browseX - $dbButtonGap - $dbX
                $dbW = [Math]::Max(1, $dbAvailableW)
                Set-HooterControlBounds -Control $dbLabel -X 16 -Y 96 -W 112 -H 22
                Set-HooterControlBounds -Control $databaseBox -X $dbX -Y 92 -W $dbW -H 28
                Set-HooterControlBounds -Control $browseDbButton -X $browseX -Y 91 -W $browseButtonW -H 30
                Set-HooterControlBounds -Control $connectButton -X $connectX -Y 91 -W $completedButtonW -H 30
                Set-HooterControlBounds -Control $loadAllPlotsButton -X $allPlotsX -Y 91 -W $allPlotsButtonW -H 30

                $loadW = 88
                $loadX = $rightEdge - $loadW
                $plotX = 78
                $plotW = [Math]::Max(190, $loadX - $buttonGap - $plotX)
                Set-HooterControlBounds -Control $plotLabel -X 16 -Y 132 -W 54 -H 22
                Set-HooterControlBounds -Control $plotBox -X $plotX -Y 128 -W $plotW -H 28
                Set-HooterControlBounds -Control $loadPlotButton -X $loadX -Y 127 -W $loadW -H 30
            }
            else {
                $header.Height = 104
                Set-HooterControlBounds -Control $picture -X 18 -Y 16 -W 70 -H 70

                $rightEdge = $clientW - 36
                $allPlotsW = 112
                $completedW = 158
                $browseW = 82
                $dbButtonGap = 28
                $allPlotsX = $rightEdge - $allPlotsW
                $connectX = $allPlotsX - 8 - $completedW
                $browseX = $connectX - 8 - $browseW
                $dbX = 710
                $dbLabelX = 590
                $dbAvailableW = $browseX - $dbButtonGap - $dbX
                $dbW = [Math]::Max(1, $dbAvailableW)
                Set-HooterControlBounds -Control $title -X 104 -Y 18 -W 260 -H 34
                Set-HooterControlBounds -Control $subtitle -X 106 -Y 54 -W ([Math]::Max(300, $dbLabelX - 122)) -H 24
                Set-HooterControlBounds -Control $projectLabel -X 106 -Y 76 -W ([Math]::Max(300, $dbLabelX - 122)) -H 22
                Set-HooterControlBounds -Control $dbLabel -X $dbLabelX -Y 16 -W 120 -H 22
                Set-HooterControlBounds -Control $databaseBox -X $dbX -Y 13 -W $dbW -H 28
                Set-HooterControlBounds -Control $browseDbButton -X $browseX -Y 12 -W $browseW -H 30
                Set-HooterControlBounds -Control $connectButton -X $connectX -Y 12 -W $completedW -H 30
                Set-HooterControlBounds -Control $loadAllPlotsButton -X $allPlotsX -Y 12 -W $allPlotsW -H 30

                $loadX = $rightEdge - 92
                $plotW = [Math]::Max(220, $loadX - 10 - $dbX)
                Set-HooterControlBounds -Control $plotLabel -X $dbLabelX -Y 58 -W 54 -H 22
                Set-HooterControlBounds -Control $plotBox -X $dbX -Y 55 -W $plotW -H 28
                Set-HooterControlBounds -Control $loadPlotButton -X $loadX -Y 53 -W 92 -H 30
            }

            $statusY = $header.Height + 6
            $statusH = if ($compact) { 42 } else { 30 }
            Set-HooterControlBounds -Control $statusLabel -X $outerMargin -Y $statusY -W ($clientW - ($outerMargin * 2)) -H $statusH
            Set-HooterControlBounds -Control $autoSaveNoticePanel -X ([Math]::Max($outerMargin, $clientW - $outerMargin - 190)) -Y $statusY -W 190 -H 30
            $tabY = $statusY + $statusH + 6
            $footerY = $clientH - 28
            Set-HooterControlBounds -Control $footer -X $outerMargin -Y $footerY -W ([Math]::Max(300, $clientW - ($outerMargin * 2))) -H 24
            Set-HooterControlBounds -Control $tabControl -X $outerMargin -Y $tabY -W ($clientW - ($outerMargin * 2)) -H ([Math]::Max(360, $footerY - $tabY - 8))

            $gridMargin = if ($compact) { 8 } else { 12 }
            $plotTop = if ($compact) { 8 } else { 16 }
            Set-HooterDataEntryGridBounds -Page $plotPage -Grid $plotGrid -Top $plotTop -Margin $gridMargin

            $treePageW = [Math]::Max(320, $treePage.ClientSize.Width)
            if ($compact) {
                $treeComboW = [Math]::Min(300, [Math]::Max(220, [int]($treePageW * 0.32)))
                $treePrevX = $gridMargin + $treeComboW + 8
                $treeNextX = $treePrevX + 118 + 8
                $treeCheckX = $treeNextX + 96 + 12
                Set-HooterControlBounds -Control $treeRecordBox -X $gridMargin -Y 10 -W $treeComboW -H 28
                Set-HooterControlBounds -Control $addTreeButton -X $treePrevX -Y 9 -W 118 -H 30
                Set-HooterControlBounds -Control $removeTreeButton -X $treeNextX -Y 9 -W 96 -H 30
                Set-HooterControlBounds -Control $treeSelectedOnlyCheck -X $treeCheckX -Y 13 -W ([Math]::Max(150, $treePageW - $treeCheckX - $gridMargin)) -H 24
                Set-HooterControlBounds -Control $crewMissedTreeCheck -X $gridMargin -Y 42 -W ($treePageW - ($gridMargin * 2)) -H 24
                Set-HooterControlBounds -Control $treeProgressLabel -X $gridMargin -Y 68 -W ($treePageW - ($gridMargin * 2)) -H 22
                Set-HooterDataEntryGridBounds -Page $treePage -Grid $treeGrid -Top 96 -Margin $gridMargin
            }
            else {
                Set-HooterControlBounds -Control $treeRecordBox -X 16 -Y 18 -W 300 -H 28
                Set-HooterControlBounds -Control $addTreeButton -X 326 -Y 16 -W 118 -H 30
                Set-HooterControlBounds -Control $removeTreeButton -X 452 -Y 16 -W 96 -H 30
                Set-HooterControlBounds -Control $treeSelectedOnlyCheck -X 565 -Y 20 -W 170 -H 24
                Set-HooterControlBounds -Control $treeProgressLabel -X 750 -Y 20 -W ([Math]::Max(260, $treePageW - 762)) -H 24
                Set-HooterControlBounds -Control $crewMissedTreeCheck -X 16 -Y 50 -W ([Math]::Max(360, $treePageW - 32)) -H 24
                Set-HooterDataEntryGridBounds -Page $treePage -Grid $treeGrid -Top 82 -Margin $gridMargin
            }

            $missedTreePageW = [Math]::Max(320, $missedTreePage.ClientSize.Width)
            if ($compact) {
                Set-HooterControlBounds -Control $missedTreeHint -X $gridMargin -Y 10 -W ($missedTreePageW - ($gridMargin * 2)) -H 42
                Set-HooterControlBounds -Control $addMissedTreeButton -X $gridMargin -Y 58 -W 130 -H 30
                Set-HooterControlBounds -Control $removeMissedTreeButton -X 148 -Y 58 -W 130 -H 30
                Set-HooterControlBounds -Control $missedTreeStatusLabel -X $gridMargin -Y 94 -W ($missedTreePageW - ($gridMargin * 2)) -H 22
                Set-HooterDataEntryGridBounds -Page $missedTreePage -Grid $missedTreeGrid -Top 122 -Margin $gridMargin
            }
            else {
                Set-HooterControlBounds -Control $missedTreeHint -X 16 -Y 18 -W ([Math]::Max(500, $missedTreePageW - 32)) -H 24
                Set-HooterControlBounds -Control $addMissedTreeButton -X 16 -Y 50 -W 130 -H 30
                Set-HooterControlBounds -Control $removeMissedTreeButton -X 156 -Y 50 -W 130 -H 30
                Set-HooterControlBounds -Control $missedTreeStatusLabel -X 304 -Y 55 -W ([Math]::Max(360, $missedTreePageW - 316)) -H 22
                Set-HooterDataEntryGridBounds -Page $missedTreePage -Grid $missedTreeGrid -Top 90 -Margin $gridMargin
            }

            $regenPageW = [Math]::Max(320, $regenPage.ClientSize.Width)
            if ($compact) {
                $regenComboW = [Math]::Min(320, [Math]::Max(230, [int]($regenPageW * 0.34)))
                $regenPrevX = $gridMargin + $regenComboW + 8
                $regenNextX = $regenPrevX + 124 + 8
                $regenCheckX = $regenNextX + 104 + 12
                Set-HooterControlBounds -Control $regenRecordBox -X $gridMargin -Y 10 -W $regenComboW -H 28
                Set-HooterControlBounds -Control $addRegenButton -X $regenPrevX -Y 9 -W 124 -H 30
                Set-HooterControlBounds -Control $removeRegenButton -X $regenNextX -Y 9 -W 104 -H 30
                Set-HooterControlBounds -Control $regenSelectedOnlyCheck -X $regenCheckX -Y 13 -W ([Math]::Max(150, $regenPageW - $regenCheckX - $gridMargin)) -H 24
                Set-HooterControlBounds -Control $regenProgressLabel -X $gridMargin -Y 42 -W ($regenPageW - ($gridMargin * 2)) -H 22
                Set-HooterDataEntryGridBounds -Page $regenPage -Grid $regenGrid -Top 68 -Margin $gridMargin
            }
            else {
                Set-HooterControlBounds -Control $regenRecordBox -X 16 -Y 18 -W 340 -H 28
                Set-HooterControlBounds -Control $addRegenButton -X 366 -Y 16 -W 124 -H 30
                Set-HooterControlBounds -Control $removeRegenButton -X 500 -Y 16 -W 104 -H 30
                Set-HooterControlBounds -Control $regenSelectedOnlyCheck -X 620 -Y 20 -W 180 -H 24
                Set-HooterControlBounds -Control $regenProgressLabel -X 815 -Y 20 -W ([Math]::Max(260, $regenPageW - 827)) -H 24
                Set-HooterDataEntryGridBounds -Page $regenPage -Grid $regenGrid -Top 58 -Margin $gridMargin
            }

            $executionPageW = [Math]::Max(320, $executionPage.ClientSize.Width)
            if ($compact) {
                Set-HooterControlBounds -Control $executionHint -X $gridMargin -Y 10 -W ($executionPageW - ($gridMargin * 2)) -H 42
                Set-HooterDataEntryGridBounds -Page $executionPage -Grid $executionGrid -Top 58 -Margin $gridMargin
            }
            else {
                Set-HooterControlBounds -Control $executionHint -X 16 -Y 18 -W ([Math]::Max(500, $executionPageW - 32)) -H 24
                Set-HooterDataEntryGridBounds -Page $executionPage -Grid $executionGrid -Top 58 -Margin $gridMargin
            }

            $reviewPageW = [Math]::Max(320, $reviewPage.ClientSize.Width)
            if ($compact) {
                $buttonGap = 8
                $buttonW = [Math]::Max(96, [int](($reviewPageW - ($gridMargin * 2) - ($buttonGap * 2)) / 3))
                $buttonX1 = $gridMargin
                $buttonX2 = $buttonX1 + $buttonW + $buttonGap
                $buttonX3 = $buttonX2 + $buttonW + $buttonGap
                Set-HooterControlBounds -Control $overallLabel -X $gridMargin -Y 10 -W ($reviewPageW - ($gridMargin * 2)) -H 28
                Set-HooterControlBounds -Control $inventoryProgressLabel -X $gridMargin -Y 40 -W ($reviewPageW - ($gridMargin * 2)) -H 22
                Set-HooterControlBounds -Control $inventoryProgressBar -X $gridMargin -Y 66 -W ($reviewPageW - ($gridMargin * 2)) -H 16
                $reviewContentW = $reviewPageW - ($gridMargin * 2)
                if ($reviewContentW -lt 460) {
                    Set-HooterControlBounds -Control $checkCruiserLabel -X $gridMargin -Y 88 -W $reviewContentW -H 22
                    Set-HooterControlBounds -Control $checkCruiserBox -X $gridMargin -Y 112 -W $reviewContentW -H 28
                    Set-HooterControlBounds -Control $checkCruiseDateLabel -X $gridMargin -Y 148 -W $reviewContentW -H 22
                    Set-HooterControlBounds -Control $checkCruiseDatePicker -X $gridMargin -Y 172 -W ([Math]::Min(180, $reviewContentW)) -H 28
                    Set-HooterControlBounds -Control $qaRemarksLabel -X $gridMargin -Y 206 -W 120 -H 22
                    Set-HooterControlBounds -Control $notesBox -X $gridMargin -Y 232 -W $reviewContentW -H 72
                    $buttonRowOneY = 318
                    $buttonRowTwoY = 358
                    $summaryY = 402
                    $errorSummaryY = 430
                    $errorSummaryH = 220
                    $reviewGridTop = $errorSummaryY + $errorSummaryH + 24
                }
                else {
                    $metadataGap = 12
                    $cruiserW = [Math]::Min(360, [int](($reviewContentW - $metadataGap) * 0.62))
                    $dateX = $gridMargin + $cruiserW + $metadataGap
                    $dateW = [Math]::Max(150, $reviewContentW - $cruiserW - $metadataGap)
                    Set-HooterControlBounds -Control $checkCruiserLabel -X $gridMargin -Y 88 -W $cruiserW -H 22
                    Set-HooterControlBounds -Control $checkCruiserBox -X $gridMargin -Y 112 -W $cruiserW -H 28
                    Set-HooterControlBounds -Control $checkCruiseDateLabel -X $dateX -Y 88 -W $dateW -H 22
                    Set-HooterControlBounds -Control $checkCruiseDatePicker -X $dateX -Y 112 -W $dateW -H 28
                    Set-HooterControlBounds -Control $qaRemarksLabel -X $gridMargin -Y 148 -W 120 -H 22
                    Set-HooterControlBounds -Control $notesBox -X $gridMargin -Y 174 -W $reviewContentW -H 72
                    $buttonRowOneY = 258
                    $buttonRowTwoY = 298
                    $summaryY = 342
                    $errorSummaryY = 370
                    $errorSummaryH = 220
                    $reviewGridTop = $errorSummaryY + $errorSummaryH + 24
                }
                Set-HooterControlBounds -Control $exportWorkbookButton -X $buttonX1 -Y $buttonRowOneY -W $buttonW -H 32
                Set-HooterControlBounds -Control $savePlotButton -X $buttonX2 -Y $buttonRowOneY -W $buttonW -H 32
                Set-HooterControlBounds -Control $exportAllButton -X $buttonX3 -Y $buttonRowOneY -W $buttonW -H 32
                Set-HooterControlBounds -Control $importButton -X $buttonX1 -Y $buttonRowTwoY -W $buttonW -H 32
                Set-HooterControlBounds -Control $exportFailCsvButton -X $buttonX2 -Y $buttonRowTwoY -W $buttonW -H 32
                Set-HooterControlBounds -Control $exportFailHtmlButton -X $buttonX3 -Y $buttonRowTwoY -W $buttonW -H 32
                Set-HooterControlBounds -Control $summaryCheckLabel -X $gridMargin -Y $summaryY -W ($reviewPageW - ($gridMargin * 2)) -H 22
                Set-HooterControlBounds -Control $errorSummaryGrid -X $gridMargin -Y $errorSummaryY -W ($reviewPageW - ($gridMargin * 2)) -H $errorSummaryH
                Set-HooterDataEntryGridBounds -Page $reviewPage -Grid $reviewGrid -Top $reviewGridTop -Margin $gridMargin
            }
            else {
                Set-HooterControlBounds -Control $overallLabel -X 16 -Y 18 -W ([Math]::Max(500, $reviewPageW - 32)) -H 28
                Set-HooterControlBounds -Control $inventoryProgressLabel -X 16 -Y 48 -W 720 -H 22
                Set-HooterControlBounds -Control $inventoryProgressBar -X 750 -Y 50 -W ([Math]::Max(260, $reviewPageW - 766)) -H 16
                Set-HooterControlBounds -Control $checkCruiserLabel -X 16 -Y 72 -W 110 -H 22
                Set-HooterControlBounds -Control $checkCruiserBox -X 16 -Y 96 -W 220 -H 28
                Set-HooterControlBounds -Control $checkCruiseDateLabel -X 252 -Y 72 -W 100 -H 22
                Set-HooterControlBounds -Control $checkCruiseDatePicker -X 252 -Y 96 -W 150 -H 28
                Set-HooterControlBounds -Control $qaRemarksLabel -X 16 -Y 130 -W 100 -H 22
                Set-HooterControlBounds -Control $notesBox -X 16 -Y 154 -W 470 -H 86
                Set-HooterControlBounds -Control $exportWorkbookButton -X 510 -Y 82 -W 140 -H 32
                Set-HooterControlBounds -Control $savePlotButton -X 656 -Y 82 -W 76 -H 32
                Set-HooterControlBounds -Control $exportAllButton -X 738 -Y 82 -W 112 -H 32
                Set-HooterControlBounds -Control $importButton -X 856 -Y 82 -W 112 -H 32
                Set-HooterControlBounds -Control $exportFailCsvButton -X 974 -Y 82 -W 82 -H 32
                Set-HooterControlBounds -Control $exportFailHtmlButton -X 1062 -Y 82 -W 136 -H 32
                Set-HooterControlBounds -Control $summaryCheckLabel -X 510 -Y 124 -W 220 -H 22
                Set-HooterControlBounds -Control $errorSummaryGrid -X 510 -Y 152 -W ([Math]::Max(360, $reviewPageW - 526)) -H 220
                Set-HooterDataEntryGridBounds -Page $reviewPage -Grid $reviewGrid -Top 390 -Margin 12
            }

            if ($settingsPage.ClientSize.Width -gt 0 -and $settingsGrid.ClientSize.Width -ge 0) {
                $settingsPageW = [Math]::Max(320, $settingsPage.ClientSize.Width)
                if ($compact) {
                    Set-HooterControlBounds -Control $saveSettingsButton -X $gridMargin -Y 10 -W 112 -H 30
                    Set-HooterControlBounds -Control $defaultSettingsButton -X 128 -Y 10 -W 120 -H 30
                    Set-HooterControlBounds -Control $failThresholdLabel -X 264 -Y 15 -W 88 -H 22
                    Set-HooterControlBounds -Control $maxPointLossBox -X 354 -Y 11 -W 54 -H 28
                    Set-HooterControlBounds -Control $plotMaxLossLabel -X 420 -Y 15 -W 92 -H 22
                    Set-HooterControlBounds -Control $plotPointTotalBox -X 512 -Y 11 -W 54 -H 28
                    Set-HooterControlBounds -Control $moveFieldUpButton -X ([Math]::Max(574, $settingsPageW - 112)) -Y 10 -W 44 -H 30
                    Set-HooterControlBounds -Control $moveFieldDownButton -X ([Math]::Max(624, $settingsPageW - 62)) -Y 10 -W 54 -H 30

                    Set-HooterControlBounds -Control $treeMaxLossLabel -X $gridMargin -Y 50 -W 92 -H 22
                    Set-HooterControlBounds -Control $treePointTotalBox -X 106 -Y 46 -W 54 -H 28
                    Set-HooterControlBounds -Control $treeMaxSummaryLabel -X 174 -Y 50 -W ([Math]::Max(260, $settingsPageW - 182)) -H 22

                    Set-HooterControlBounds -Control $regenMaxLossLabel -X $gridMargin -Y 84 -W 106 -H 22
                    Set-HooterControlBounds -Control $regenPointTotalBox -X 116 -Y 80 -W 54 -H 28
                    Set-HooterControlBounds -Control $executionMaxLossLabel -X 188 -Y 84 -W 116 -H 22
                    Set-HooterControlBounds -Control $executionPointTotalBox -X 306 -Y 80 -W 54 -H 28
                    Set-HooterControlBounds -Control $useStemPercentCheck -X $gridMargin -Y 118 -W ([Math]::Max(280, $settingsPageW - ($gridMargin * 2))) -H 24
                    Set-HooterControlBounds -Control $stemToleranceLabel -X $gridMargin -Y 150 -W ([Math]::Max(260, $settingsPageW - ($gridMargin * 2))) -H 22
                    Set-HooterControlBounds -Control $stemToleranceGrid -X $gridMargin -Y 176 -W ([Math]::Max(280, [Math]::Min(390, $settingsPageW - ($gridMargin * 2)))) -H 188
                    Set-HooterControlBounds -Control $settingsFilterLabel -X $gridMargin -Y 372 -W 42 -H 22
                    Set-HooterControlBounds -Control $settingsFilterBox -X 58 -Y 368 -W ([Math]::Min(180, [Math]::Max(150, $settingsPageW - 66))) -H 28
                    Set-HooterDataEntryGridBounds -Page $settingsPage -Grid $settingsGrid -Top 406 -Margin $gridMargin
                }
                else {
                    Set-HooterControlBounds -Control $saveSettingsButton -X 16 -Y 16 -W 120 -H 30
                    Set-HooterControlBounds -Control $defaultSettingsButton -X 146 -Y 16 -W 120 -H 30
                    Set-HooterControlBounds -Control $failThresholdLabel -X 276 -Y 21 -W 88 -H 22
                    Set-HooterControlBounds -Control $maxPointLossBox -X 366 -Y 17 -W 54 -H 28
                    Set-HooterControlBounds -Control $plotMaxLossLabel -X 424 -Y 21 -W 92 -H 22
                    Set-HooterControlBounds -Control $plotPointTotalBox -X 516 -Y 17 -W 54 -H 28
                    Set-HooterControlBounds -Control $treeMaxLossLabel -X 588 -Y 21 -W 92 -H 22
                    Set-HooterControlBounds -Control $treePointTotalBox -X 680 -Y 17 -W 54 -H 28
                    Set-HooterControlBounds -Control $regenMaxLossLabel -X 752 -Y 21 -W 106 -H 22
                    Set-HooterControlBounds -Control $regenPointTotalBox -X 858 -Y 17 -W 54 -H 28
                    Set-HooterControlBounds -Control $executionMaxLossLabel -X 918 -Y 21 -W 116 -H 22
                    Set-HooterControlBounds -Control $executionPointTotalBox -X 1034 -Y 17 -W 54 -H 28
                    Set-HooterControlBounds -Control $moveFieldUpButton -X 1098 -Y 16 -W 44 -H 30
                    Set-HooterControlBounds -Control $moveFieldDownButton -X 1148 -Y 16 -W 54 -H 30
                    Set-HooterControlBounds -Control $treeMaxSummaryLabel -X 588 -Y 46 -W ([Math]::Max(360, $settingsPageW - 600)) -H 22
                    Set-HooterControlBounds -Control $useStemPercentCheck -X 16 -Y 76 -W 318 -H 24
                    Set-HooterControlBounds -Control $stemToleranceLabel -X 16 -Y 108 -W 300 -H 22
                    Set-HooterControlBounds -Control $stemToleranceGrid -X 16 -Y 134 -W 318 -H 190
                    Set-HooterControlBounds -Control $settingsFilterLabel -X 352 -Y 48 -W 42 -H 22
                    Set-HooterControlBounds -Control $settingsFilterBox -X 398 -Y 44 -W 168 -H 28
                    $settingsGridX = 352
                    Set-HooterControlBounds -Control $settingsGrid -X $settingsGridX -Y 78 -W ([Math]::Max(420, $settingsPageW - $settingsGridX - 12)) -H ([Math]::Max(180, $settingsPage.ClientSize.Height - 90))
                }
            }
            if ($guideBox.ClientSize.Width -ge 0) {
                Set-HooterControlBounds -Control $guideBox -X $gridMargin -Y 58 -W ([Math]::Max(260, $userGuidePage.ClientSize.Width - ($gridMargin * 2))) -H ([Math]::Max(240, $userGuidePage.ClientSize.Height - 68))
            }
        }
        catch {}
    }.GetNewClosure()
    $form.Add_Resize({ & $applyResponsiveLayout }.GetNewClosure())
    $form.Add_Shown({ & $applyResponsiveLayout }.GetNewClosure())
    $form.Add_Shown({ Close-HooterSplashScreen -SplashForm $splash }.GetNewClosure())
    $tabControl.Add_Selecting({
        try {
            if ($script:Ui.ContainsKey("Form")) { $script:Ui.Form.UseWaitCursor = $true }
            Show-HooterAutoSaveNotice -Text "Opening tab..."
        }
        catch {
        }
    }.GetNewClosure())
    $tabControl.Add_SelectedIndexChanged({
        try {
            Show-HooterAutoSaveNotice -Text "Opening tab..."
            & $applyResponsiveLayout
        }
        finally {
            try {
                if ($script:Ui.ContainsKey("Form")) { $script:Ui.Form.UseWaitCursor = $false }
            }
            catch {
            }
            Hide-HooterAutoSaveNotice -DelayMs 650
        }
    }.GetNewClosure())
    & $applyResponsiveLayout

    foreach ($grid in @($plotGrid, $treeGrid, $missedTreeGrid, $regenGrid, $executionGrid, $stemToleranceGrid, $settingsGrid)) {
        Register-HooterGridInputScopes -Grid $grid
    }
    Register-HooterTextInputScope -Control $databaseBox -Scope "Text"
    Register-HooterTextInputScope -Control $checkCruiserBox -Scope "Text"
    Register-HooterTextInputScope -Control $notesBox -Scope "Text"
    foreach ($scoreBox in @($maxPointLossBox, $plotPointTotalBox, $treePointTotalBox, $regenPointTotalBox, $executionPointTotalBox)) {
        Register-HooterTextInputScope -Control $scoreBox -Scope "Number"
    }

    $script:Ui = @{
        Form = $form
        ToolTip = $toolTip
        TabControl = $tabControl
        MissedTreePage = $missedTreePage
        ReviewPage = $reviewPage
        ProjectLabel = $projectLabel
        DatabaseBox = $databaseBox
        PlotBox = $plotBox
        LoadAllPlotsButton = $loadAllPlotsButton
        StatusLabel = $statusLabel
        AutoSaveNoticePanel = $autoSaveNoticePanel
        AutoSaveNoticeLabel = $autoSaveNoticeLabel
        PlotGrid = $plotGrid
        TreeGrid = $treeGrid
        RegenGrid = $regenGrid
        ExecutionGrid = $executionGrid
        TreeRecordBox = $treeRecordBox
        TreeSelectedOnlyCheck = $treeSelectedOnlyCheck
        CrewMissedTreeCheck = $crewMissedTreeCheck
        TreeProgressLabel = $treeProgressLabel
        MissedTreeGrid = $missedTreeGrid
        MissedTreeStatusLabel = $missedTreeStatusLabel
        RegenRecordBox = $regenRecordBox
        RegenSelectedOnlyCheck = $regenSelectedOnlyCheck
        RegenProgressLabel = $regenProgressLabel
        UseStemPercentCheck = $useStemPercentCheck
        StemToleranceLabel = $stemToleranceLabel
        StemToleranceGrid = $stemToleranceGrid
        SettingsFilterBox = $settingsFilterBox
        SettingsGrid = $settingsGrid
        MaxPointLossBox = $maxPointLossBox
        PlotPointTotalBox = $plotPointTotalBox
        TreePointTotalBox = $treePointTotalBox
        TreeMaxSummaryLabel = $treeMaxSummaryLabel
        RegenPointTotalBox = $regenPointTotalBox
        ExecutionPointTotalBox = $executionPointTotalBox
        ErrorSummaryGrid = $errorSummaryGrid
        ReviewGrid = $reviewGrid
        CheckCruiserBox = $checkCruiserBox
        CheckCruiseDatePicker = $checkCruiseDatePicker
        NotesBox = $notesBox
        OverallLabel = $overallLabel
        InventoryProgressLabel = $inventoryProgressLabel
        InventoryProgressBar = $inventoryProgressBar
        ExportWorkbookButton = $exportWorkbookButton
    }
    Populate-HooterStemToleranceGrid
    Update-HooterStemScoringControls
    Populate-HooterExecutionGrid
    Update-HooterOverallLabel

    $browseDbButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Access databases (*.mdb;*.accdb)|*.mdb;*.accdb|All files (*.*)|*.*"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-HooterDatabaseBoxPath -TextBox $script:Ui.DatabaseBox -Path $dialog.FileName
        }
    })
    $connectButton.Add_Click({ Connect-HooterDatabase -CompletedOnly })
    $loadAllPlotsButton.Add_Click({ Connect-HooterDatabase })
    $loadPlotButton.Add_Click({ Load-HooterPlot })
    $plotBox.Add_KeyDown({
        param($Sender, $KeyEvent)
        if ($KeyEvent.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $KeyEvent.Handled = $true
            $KeyEvent.SuppressKeyPress = $true
            Load-HooterPlot
        }
    })
    $savePlotButton.Add_Click({ Save-HooterCurrentSession })
    $addTreeButton.Add_Click({ Select-HooterTreeOffset -Offset -1 })
    $removeTreeButton.Add_Click({ Select-HooterTreeOffset -Offset 1 })
    $treeRecordBox.Add_SelectedIndexChanged({ Set-HooterTreeVisibleRows })
    $treeSelectedOnlyCheck.Add_CheckedChanged({ Set-HooterTreeVisibleRows })
    $crewMissedTreeCheck.Add_CheckedChanged({
        if ($script:SuppressMissedTreeCheckEvent) { return }
        if (-not $crewMissedTreeCheck.Checked -and (Get-HooterMissedTreeRowCount) -gt 0) {
            try {
                $script:SuppressMissedTreeCheckEvent = $true
                $crewMissedTreeCheck.Checked = $true
            }
            finally {
                $script:SuppressMissedTreeCheckEvent = $false
            }
            Set-HooterStatus "Remove logged missed tree rows before clearing the missed-tree plot fail check."
            return
        }
        Show-HooterAutoSaveNotice -Text "Updating missed tree check..."
        if ($crewMissedTreeCheck.Checked -and $script:Ui.ContainsKey("TabControl") -and $script:Ui.ContainsKey("MissedTreePage")) {
            try { $script:Ui.TabControl.SelectedTab = $script:Ui.MissedTreePage } catch {}
        }
        Request-HooterMissedTreeRefresh
    })
    $addMissedTreeButton.Add_Click({ Add-HooterMissedTreeRow })
    $removeMissedTreeButton.Add_Click({ Remove-HooterSelectedMissedTrees })
    $addRegenButton.Add_Click({ Select-HooterRegenOffset -Offset -1 })
    $removeRegenButton.Add_Click({ Select-HooterRegenOffset -Offset 1 })
    $regenRecordBox.Add_SelectedIndexChanged({ Set-HooterRegenVisibleRows })
    $regenSelectedOnlyCheck.Add_CheckedChanged({ Set-HooterRegenVisibleRows })
    $useStemPercentCheck.Add_CheckedChanged({
        if ($script:SuppressStemScoringToggleEvent) { return }
        Set-HooterUseStemCountPercentageForScoring -Enabled ([bool]$useStemPercentCheck.Checked) -Refresh
        Set-HooterStatus $(if ($useStemPercentCheck.Checked) { "Using regen stem count percentage scoring. The stem-count table is off." } else { "Using regen stem count table scoring. The table is active." })
    })
    $moveFieldUpButton.Add_Click({ Move-HooterSettingsFieldRow -Offset -1 })
    $moveFieldDownButton.Add_Click({ Move-HooterSettingsFieldRow -Offset 1 })
    $saveSettingsButton.Add_Click({
        Update-HooterSettingsFromGrid
        Invoke-HooterAutoSaveOperation -Text "Saving setup..." -Action { Save-HooterSettings }
        Apply-HooterFieldOrderToLoadedPlot
        Refresh-HooterAllStatuses
        Set-HooterStatus "Saved PlotHoot tolerance setup."
    })
    foreach ($scoreBox in @($maxPointLossBox, $plotPointTotalBox, $treePointTotalBox, $regenPointTotalBox, $executionPointTotalBox)) {
        $scoreBox.Add_Leave({
            Update-HooterSettingsFromGrid
            Refresh-HooterAllStatuses
        })
    }
    $defaultSettingsButton.Add_Click({
        $script:Tolerances = @{}
        $script:FieldOrder = @{}
        $script:ScorePassPercent = $script:DefaultScorePassPercent
        $script:MaxPointLoss = $script:DefaultMaxPointLoss
        $script:PlotPointTotal = $script:DefaultPlotPointTotal
        $script:TreePointTotal = $script:DefaultTreePointTotal
        $script:RegenPointTotal = $script:DefaultRegenPointTotal
        $script:ExecutionPointTotal = $script:DefaultExecutionPointTotal
        $script:UseStemCountPercentageForScoring = $script:DefaultUseStemCountPercentageForScoring
        Reset-HooterStemToleranceBands
        if ($script:Ui.ContainsKey("MaxPointLossBox")) { $script:Ui.MaxPointLossBox.Text = $script:MaxPointLoss }
        if ($script:Ui.ContainsKey("PlotPointTotalBox")) { $script:Ui.PlotPointTotalBox.Text = $script:PlotPointTotal }
        if ($script:Ui.ContainsKey("TreePointTotalBox")) { $script:Ui.TreePointTotalBox.Text = $script:TreePointTotal }
        if ($script:Ui.ContainsKey("RegenPointTotalBox")) { $script:Ui.RegenPointTotalBox.Text = $script:RegenPointTotal }
        if ($script:Ui.ContainsKey("ExecutionPointTotalBox")) { $script:Ui.ExecutionPointTotalBox.Text = $script:ExecutionPointTotal }
        if ($script:Ui.ContainsKey("StemToleranceGrid")) { Populate-HooterStemToleranceGrid }
        Update-HooterStemScoringControls
        Apply-HooterSavedFieldOrderToCatalog
        Populate-HooterSettingsGrid
        Apply-HooterFieldOrderToLoadedPlot
        Refresh-HooterAllStatuses
        Set-HooterStatus "Reset tolerance and scoring defaults. Saved custom field order and unchecked field choices were cleared, so fields now use database query/AppColumns order."
    })
    $importButton.Add_Click({ Import-HooterQaCsv })
    $exportAllButton.Add_Click({ Export-HooterQaCsv })
    $exportFailCsvButton.Add_Click({ Export-HooterFailedCsv })
    $exportFailHtmlButton.Add_Click({ Export-HooterQaReportHtml })
    $exportWorkbookButton.Add_Click({ Export-HooterQaWorkbookXlsx })
    $settingsFilterBox.Add_SelectedIndexChanged({ Apply-HooterSettingsFilter })
    $openGuideButton.Add_Click({
        $rootFolder = Split-Path -Parent $script:AppRoot
        $guidePath = Join-Path $rootFolder "PlotHoot User Guide.html"
        try {
            if (-not (Test-Path -LiteralPath $guidePath)) {
                $guidePath = Join-Path $script:AppRoot "PlotHoot User Guide.html"
            }
            if (-not (Test-Path -LiteralPath $guidePath)) {
                Ensure-HooterDataRoot
                $guidePath = Join-Path $script:DataRoot "PlotHoot User Guide.html"
                New-HooterUserGuideHtml | Set-Content -LiteralPath $guidePath -Encoding UTF8
            }
            if (Test-Path -LiteralPath $guidePath) {
                Start-Process -FilePath $guidePath
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Open guide", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    foreach ($grid in @($plotGrid, $treeGrid, $regenGrid)) {
        $grid.Add_CellClick({
            param($Sender, $EventArgs)
            if ($EventArgs.RowIndex -lt 0 -or $EventArgs.ColumnIndex -lt 0) { return }
            $name = $Sender.Columns[$EventArgs.ColumnIndex].Name
            if ($name -eq "QaValue" -and -not ($Sender.Rows[$EventArgs.RowIndex].Cells["QaValue"] -is [System.Windows.Forms.DataGridViewCheckBoxCell])) {
                [void](Start-HooterGridCellEdit -Grid $Sender -RowIndex $EventArgs.RowIndex -ColumnName "QaValue" -SelectAll)
            }
            elseif ($name -eq "PointValue") {
                [void](Start-HooterGridCellEdit -Grid $Sender -RowIndex $EventArgs.RowIndex -ColumnName "PointValue" -SelectAll)
            }
        })
        $grid.Add_CellDoubleClick({
            param($Sender, $EventArgs)
            if ($EventArgs.RowIndex -lt 0 -or $EventArgs.ColumnIndex -lt 0) { return }
            if ($Sender.Columns[$EventArgs.ColumnIndex].Name -eq "CrewValue") {
                Show-HooterCrewValuePopup -Grid $Sender -RowIndex $EventArgs.RowIndex
            }
        })
        $grid.Add_CellValueChanged({
            param($Sender, $EventArgs)
            if ($script:SuppressGridEvents) { return }
            if ($EventArgs.RowIndex -lt 0) { return }
            if ($EventArgs.ColumnIndex -lt 0) { return }
            $name = $Sender.Columns[$EventArgs.ColumnIndex].Name
            if ($name -eq "PointValue") {
                Update-HooterWorkingPointOverrideFromRow -Row $Sender.Rows[$EventArgs.RowIndex]
                return
            }
            if ($name -in @("QaValue", "Notes")) {
                Update-HooterRowStatus -Row $Sender.Rows[$EventArgs.RowIndex]
                Request-HooterOverallRefresh
                if ($Sender -eq $script:Ui.TreeGrid) { Request-HooterTreeProgressRefresh }
                if ($Sender -eq $script:Ui.RegenGrid) { Request-HooterRegenProgressRefresh }
            }
            if ($name -eq "QaValue" -and $Sender.ContainsFocus -and ($Sender.Rows[$EventArgs.RowIndex].Cells["QaValue"] -is [System.Windows.Forms.DataGridViewCheckBoxCell])) {
                Invoke-HooterAutoAdvanceFromCell -Grid $Sender -RowIndex $EventArgs.RowIndex -ColumnName $name
            }
        })
        $grid.Add_CurrentCellDirtyStateChanged({
            param($Sender, $EventArgs)
            if ($Sender.IsCurrentCellDirty -and $null -ne $Sender.CurrentCell -and ($Sender.CurrentCell -is [System.Windows.Forms.DataGridViewCheckBoxCell])) {
                [void]$Sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            }
        })
        $grid.Add_DataError({
            param($Sender, $EventArgs)
            $EventArgs.ThrowException = $false
        })
    }
    $missedTreeGrid.Add_CellValueChanged({
        param($Sender, $EventArgs)
        if ($script:SuppressGridEvents) { return }
        if ($EventArgs.RowIndex -lt 0) { return }
        if ($EventArgs.ColumnIndex -lt 0) { return }
        if ($Sender.Rows.Count -gt 0 -and $script:Ui.ContainsKey("CrewMissedTreeCheck")) {
            try {
                $script:SuppressMissedTreeCheckEvent = $true
                $script:Ui.CrewMissedTreeCheck.Checked = $true
            }
            finally {
                $script:SuppressMissedTreeCheckEvent = $false
            }
        }
        Show-HooterAutoSaveNotice -Text "Updating missed tree check..."
        Request-HooterMissedTreeRefresh
    })
    $missedTreeGrid.Add_DataError({
        param($Sender, $EventArgs)
        $EventArgs.ThrowException = $false
    })
    $executionGrid.Add_CellClick({
        param($Sender, $EventArgs)
        if ($EventArgs.RowIndex -lt 0 -or $EventArgs.ColumnIndex -lt 0) { return }
        if ($Sender.Columns[$EventArgs.ColumnIndex].Name -eq "Rating") {
            [void](Start-HooterGridCellEdit -Grid $Sender -RowIndex $EventArgs.RowIndex -ColumnName "Rating")
        }
    })
    $executionGrid.Add_CellValueChanged({
        param($Sender, $EventArgs)
        if ($script:SuppressGridEvents) { return }
        if ($EventArgs.RowIndex -lt 0) { return }
        if ($EventArgs.ColumnIndex -lt 0) { return }
        $name = $Sender.Columns[$EventArgs.ColumnIndex].Name
        if ($name -in @("Rating", "Notes")) {
            Update-HooterExecutionRowStatus -Row $Sender.Rows[$EventArgs.RowIndex]
            Update-HooterOverallLabel
        }
        if ($name -eq "Rating" -and $Sender.ContainsFocus) {
            Invoke-HooterAutoAdvanceFromCell -Grid $Sender -RowIndex $EventArgs.RowIndex -ColumnName $name
        }
    })
    $executionGrid.Add_CurrentCellDirtyStateChanged({
        if ($executionGrid.IsCurrentCellDirty) {
            [void]$executionGrid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $executionGrid.Add_DataError({
        param($Sender, $EventArgs)
        $EventArgs.ThrowException = $false
    })
    $stemToleranceGrid.Add_CellValueChanged({
        param($Sender, $EventArgs)
        if ($script:SuppressGridEvents) { return }
        if ($EventArgs.RowIndex -lt 0) { return }
        if ($EventArgs.ColumnIndex -lt 0) { return }
        if ($Sender.Columns[$EventArgs.ColumnIndex].Name -ne "ToleranceValue") { return }
        Update-HooterStemToleranceBandsFromGrid
        Refresh-HooterAllStatuses
    })
    $stemToleranceGrid.Add_DataError({
        param($Sender, $EventArgs)
        $EventArgs.ThrowException = $false
    })
    $settingsGrid.Add_CellValueChanged({
        param($Sender, $EventArgs)
        if ($script:SuppressGridEvents) { return }
        if ($EventArgs.RowIndex -lt 0) { return }
        Update-HooterSettingsFromGrid
        $name = if ($EventArgs.ColumnIndex -ge 0) { $Sender.Columns[$EventArgs.ColumnIndex].Name } else { "" }
        if ($name -eq "PointValue" -and $Sender.Columns.Contains("FieldKey")) {
            Set-HooterWorkingPointOverride -FieldKey $Sender.Rows[$EventArgs.RowIndex].Cells["FieldKey"].Value -PointValue $Sender.Rows[$EventArgs.RowIndex].Cells["PointValue"].Value | Out-Null
        }
        elseif ($name -eq "UseField") {
            Show-HooterAutoSaveNotice -Text "Updating field list..."
            Apply-HooterFieldOrderToLoadedPlot
            Request-HooterDeferredSettingsSave
        }
        elseif ($name -eq "Mode") {
            Update-HooterSettingsGridRowStyle -Row $Sender.Rows[$EventArgs.RowIndex]
            Refresh-HooterAllStatuses
        }
        else {
            Refresh-HooterAllStatuses
        }
    })
    $settingsGrid.Add_CurrentCellDirtyStateChanged({
        if ($settingsGrid.IsCurrentCellDirty) {
            [void]$settingsGrid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $settingsGrid.Add_DataError({
        param($Sender, $EventArgs)
        $EventArgs.ThrowException = $false
    })

    if (-not [string]::IsNullOrWhiteSpace($Database) -and (Test-Path -LiteralPath $Database)) {
        $form.Add_Shown({ Connect-HooterDatabase -CompletedOnly })
    }
    if ($UiSmokeTest) {
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 800
        $timer.Add_Tick({
            $timer.Stop()
            $script:Ui.Form.Close()
        })
        $form.Add_Shown({ $timer.Start() })
    }

    [void][System.Windows.Forms.Application]::Run($form)
    }
    catch {
        Close-HooterSplashScreen -SplashForm $splash
        throw
    }
}

function Assert-HooterToleranceStatus {
    param(
        [string]$Name,
        [object]$CrewValue,
        [object]$QaValue,
        [object]$Tolerance,
        [string]$ExpectedStatus
    )

    $result = Test-HooterFieldPass -CrewValue $CrewValue -QaValue $QaValue -Tolerance $Tolerance
    if ((ConvertTo-HooterText $result.Status) -ne $ExpectedStatus) {
        throw "$Name expected '$ExpectedStatus' but got '$($result.Status)': $($result.Detail)"
    }
    return $result
}

function Invoke-HooterToleranceRuleSelfTest {
    $exactRule = [pscustomobject]@{ FieldKey = "Self|Tolerance|Species"; Group = "Tree"; TableName = "Self"; FieldName = "Species"; Label = "Species"; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
    [void](Assert-HooterToleranceStatus -Name "Exact pass" -CrewValue "202" -QaValue "202" -Tolerance $exactRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Exact fail" -CrewValue "202" -QaValue "201" -Tolerance $exactRule -ExpectedStatus "Fail")
    [void](Assert-HooterToleranceStatus -Name "Blank QA not checked" -CrewValue "202" -QaValue "" -Tolerance $exactRule -ExpectedStatus "")

    $rangeRule = [pscustomobject]@{ FieldKey = "Self|Tolerance|Elevation"; Group = "Plot"; TableName = "Self"; FieldName = "Elevation"; Label = "Elevation"; Mode = "Range"; Value = "100"; PointValue = "1"; CriticalFail = $false }
    [void](Assert-HooterToleranceStatus -Name "Range pass" -CrewValue "5000" -QaValue "5100" -Tolerance $rangeRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Range fail" -CrewValue "5000" -QaValue "5101" -Tolerance $rangeRule -ExpectedStatus "Fail")

    $dbhRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "IDBH"
    [void](Assert-HooterToleranceStatus -Name "DBH decimal pass" -CrewValue "10.0" -QaValue "10.08" -Tolerance $dbhRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "DBH decimal fail" -CrewValue "10.0" -QaValue "10.25" -Tolerance $dbhRule -ExpectedStatus "Fail")
    [void](Assert-HooterToleranceStatus -Name "IDBH tenths pass" -CrewValue "60" -QaValue "62" -Tolerance $dbhRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "IDBH tenths fail" -CrewValue "60" -QaValue "63" -Tolerance $dbhRule -ExpectedStatus "Fail")
    $utmRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "UTMEasting" -Label "UTM Easting Coordinate"
    if ((ConvertTo-HooterText $utmRule.Mode) -ne "Range" -or (ConvertTo-HooterText $utmRule.Value) -ne "30") { throw "UTM range default self-test failed." }
    $regenIdbhRule = Get-DefaultToleranceForField -Group "Regen" -TableName "RegenMeasurements" -FieldName "IDBH"
    if ((ConvertTo-HooterText $regenIdbhRule.Mode) -ne "Exact") { throw "Regen IDBH exact default self-test failed." }

    $heightRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "TotalHeight"
    [void](Assert-HooterToleranceStatus -Name "Relative percent pass" -CrewValue "100" -QaValue "105" -Tolerance $heightRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Relative percent fail" -CrewValue "100" -QaValue "106" -Tolerance $heightRule -ExpectedStatus "Fail")
    $negativePercentRule = [pscustomobject]@{ FieldKey = "Self|Tolerance|TotalHeight"; Group = "Tree"; TableName = "Self"; FieldName = "TotalHeight"; Label = "Total Height"; Mode = "Percent"; Value = "-5"; PointValue = "1"; CriticalFail = $false }
    [void](Assert-HooterToleranceStatus -Name "Negative relative percent normalized" -CrewValue "100" -QaValue "105" -Tolerance $negativePercentRule -ExpectedStatus "Pass")

    $slopePercentRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "SlopePercent" -Label "Slope Percent"
    [void](Assert-HooterToleranceStatus -Name "Percentage-point pass" -CrewValue "40" -QaValue "45" -Tolerance $slopePercentRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Percentage-point fail" -CrewValue "40" -QaValue "46" -Tolerance $slopePercentRule -ExpectedStatus "Fail")
    $crownRatioRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "CrownRatio" -Label "Crown Ratio"
    [void](Assert-HooterToleranceStatus -Name "Ratio percentage-point pass" -CrewValue "30" -QaValue "40" -Tolerance $crownRatioRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Ratio percentage-point fail" -CrewValue "30" -QaValue "41" -Tolerance $crownRatioRule -ExpectedStatus "Fail")

    $classRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "CrownClass"
    [void](Assert-HooterToleranceStatus -Name "Class pass" -CrewValue "2" -QaValue "3" -Tolerance $classRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Class fail" -CrewValue "2" -QaValue "4" -Tolerance $classRule -ExpectedStatus "Fail")

    $passFailRule = [pscustomobject]@{ FieldKey = "Self|Tolerance|TreeFound"; Group = "Tree"; TableName = "Self"; FieldName = "TreeFound"; Label = "Tree Found"; Mode = "PassFail"; Value = ""; PointValue = "1"; CriticalFail = $true }
    [void](Assert-HooterToleranceStatus -Name "PassFail equivalent pass" -CrewValue "Yes" -QaValue "Pass" -Tolerance $passFailRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "PassFail fail" -CrewValue "Pass" -QaValue "Fail" -Tolerance $passFailRule -ExpectedStatus "Fail")
    $passFailScore = Get-HooterRowScoreResult -Status "Fail" -Tolerance $passFailRule -CrewValue "Pass" -QaValue "Fail"
    if (-not $passFailScore.CriticalFailure -or (ConvertTo-HooterText $passFailScore.Display) -ne "Critical fail") { throw "PassFail critical score self-test failed." }

    $filledRule = Get-DefaultToleranceForField -Group "Tree" -TableName "Trees" -FieldName "TreeNumber"
    [void](Assert-HooterToleranceStatus -Name "Filled pass" -CrewValue "12" -QaValue "Yes" -Tolerance $filledRule -ExpectedStatus "Pass")
    [void](Assert-HooterToleranceStatus -Name "Filled blank fail" -CrewValue "" -QaValue "Yes" -Tolerance $filledRule -ExpectedStatus "Fail")
    [void](Assert-HooterToleranceStatus -Name "Filled marked no fail" -CrewValue "12" -QaValue "No" -Tolerance $filledRule -ExpectedStatus "Fail")
    [void](Assert-HooterToleranceStatus -Name "Filled unchecked" -CrewValue "12" -QaValue "" -Tolerance $filledRule -ExpectedStatus "")

    $savedStemToleranceBands = @($script:StemToleranceBands | ForEach-Object { Copy-HooterStemToleranceBand -Band $_ })
    $savedStemPercentPreference = [bool]$script:UseStemCountPercentageForScoring
    try {
        Reset-HooterStemToleranceBands
        $script:UseStemCountPercentageForScoring = $false
        $stemRule = Get-DefaultToleranceForField -Group "Regen" -TableName "RegenMeasurements" -FieldName "StemCount"
        $defaultBand = Get-HooterStemToleranceBandSetting -Key "100plus"
        if ((ConvertTo-HooterText $defaultBand.Tolerance) -ne "") { throw "Regen stem count blank default self-test failed." }
        [void](Assert-HooterToleranceStatus -Name "StemCount exact pass" -CrewValue "100" -QaValue "100" -Tolerance $stemRule -ExpectedStatus "Pass")
        [void](Assert-HooterToleranceStatus -Name "StemCount exact fail" -CrewValue "99" -QaValue "100" -Tolerance $stemRule -ExpectedStatus "Fail")
        [void](Set-HooterStemToleranceBand -Key "100plus" -Tolerance "20")
        [void](Assert-HooterToleranceStatus -Name "StemCount band pass" -CrewValue "80" -QaValue "100" -Tolerance $stemRule -ExpectedStatus "Pass")
        [void](Assert-HooterToleranceStatus -Name "StemCount band fail" -CrewValue "79" -QaValue "100" -Tolerance $stemRule -ExpectedStatus "Fail")

        $stemPercentRule = [pscustomobject]@{ FieldKey = "Self|Tolerance|StemCount"; Group = "Regen"; TableName = "Self"; FieldName = "StemCount"; Label = "Stem Count"; Mode = "StemPercent"; Value = "Allowance=10; Step=5"; PointValue = "2"; CriticalFail = $false }
        [void](Assert-HooterToleranceStatus -Name "StemPercent pass" -CrewValue "90" -QaValue "100" -Tolerance $stemPercentRule -ExpectedStatus "Pass")
        [void](Assert-HooterToleranceStatus -Name "StemPercent fail" -CrewValue "5" -QaValue "7" -Tolerance $stemPercentRule -ExpectedStatus "Fail")
        $stemPercentScore = Get-HooterRowScoreResult -Status "Fail" -Tolerance $stemPercentRule -CrewValue "5" -QaValue "7"
        if ($stemPercentScore.LostPoints -ne 8) { throw "StemPercent point-loss self-test failed." }

        $stemSwitchTableMode = ConvertTo-HooterText ((Get-HooterTolerance -Field "Regen|RegenMeasurements|StemCount").Mode)
        $script:UseStemCountPercentageForScoring = $true
        $stemSwitchPercentRule = Get-HooterTolerance -Field "Regen|RegenMeasurements|StemCount"
        if ($stemSwitchTableMode -ne "StemCount" -or
            (ConvertTo-HooterText $stemSwitchPercentRule.Mode) -ne "StemPercent" -or
            [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stemSwitchPercentRule.Value))) { throw "Regen stem scoring switch self-test failed." }
    }
    finally {
        $script:StemToleranceBands = @($savedStemToleranceBands | ForEach-Object { Copy-HooterStemToleranceBand -Band $_ })
        $script:UseStemCountPercentageForScoring = $savedStemPercentPreference
    }
}

function Invoke-HooterSelfTest {
    Load-HooterSettings
    Invoke-HooterToleranceRuleSelfTest
    if ((Get-HooterDatabaseDisplayName -Path "C:\Temp\ProjectName.accdb") -ne "ProjectName.accdb") { throw "Database display name self-test failed." }
    $dbhRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "IDBH"
    $dbhPass = Test-HooterFieldPass -CrewValue "10.0" -QaValue "10.08" -Tolerance $dbhRule
    $dbhFail = Test-HooterFieldPass -CrewValue "10.0" -QaValue "10.25" -Tolerance $dbhRule
    $dbhTenthsPass = Test-HooterFieldPass -CrewValue "60" -QaValue "62" -Tolerance $dbhRule
    $dbhTenthsFail = Test-HooterFieldPass -CrewValue "60" -QaValue "63" -Tolerance $dbhRule
    if ($dbhPass.Status -ne "Pass" -or $dbhFail.Status -ne "Fail" -or $dbhTenthsPass.Status -ne "Pass" -or $dbhTenthsFail.Status -ne "Fail") { throw "DBH tolerance self-test failed." }

    $heightRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "TotalHeight"
    $heightPass = Test-HooterFieldPass -CrewValue "100" -QaValue "105" -Tolerance $heightRule
    $heightFail = Test-HooterFieldPass -CrewValue "100" -QaValue "106" -Tolerance $heightRule
    if ($heightPass.Status -ne "Pass" -or $heightFail.Status -ne "Fail") { throw "Height tolerance self-test failed." }

    $slopePercentRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "SlopePercent" -Label "Slope Percent"
    $slopePercentPass = Test-HooterFieldPass -CrewValue "40" -QaValue "45" -Tolerance $slopePercentRule
    $slopePercentFail = Test-HooterFieldPass -CrewValue "40" -QaValue "46" -Tolerance $slopePercentRule
    if ($slopePercentPass.Status -ne "Pass" -or $slopePercentFail.Status -ne "Fail") { throw "Percent-point tolerance self-test failed." }

    $classRule = Get-DefaultToleranceForField -Group "Tree" -TableName "TreeMeasurements" -FieldName "CrownClass"
    $classPass = Test-HooterFieldPass -CrewValue "2" -QaValue "3" -Tolerance $classRule
    $classFail = Test-HooterFieldPass -CrewValue "2" -QaValue "4" -Tolerance $classRule
    if ($classPass.Status -ne "Pass" -or $classFail.Status -ne "Fail") { throw "Class tolerance self-test failed." }

    $savedStemToleranceBands = @($script:StemToleranceBands | ForEach-Object { Copy-HooterStemToleranceBand -Band $_ })
    $savedStemPercentPreference = [bool]$script:UseStemCountPercentageForScoring
    try {
        Reset-HooterStemToleranceBands
        $script:UseStemCountPercentageForScoring = $false
        $stemRule = Get-DefaultToleranceForField -Group "Regen" -TableName "RegenMeasurements" -FieldName "StemCount"
        $defaultBand = Get-HooterStemToleranceBandSetting -Key "100plus"
        if ((ConvertTo-HooterText $defaultBand.Tolerance) -ne "") { throw "Regen stem count blank default self-test failed." }
        $stemExactPass = Test-HooterFieldPass -CrewValue "100" -QaValue "100" -Tolerance $stemRule
        $stemExactFail = Test-HooterFieldPass -CrewValue "99" -QaValue "100" -Tolerance $stemRule
        [void](Set-HooterStemToleranceBand -Key "100plus" -Tolerance "20")
        $stemQaBandPass = Test-HooterFieldPass -CrewValue "99" -QaValue "100" -Tolerance $stemRule
        $stemQaBandFail = Test-HooterFieldPass -CrewValue "79" -QaValue "100" -Tolerance $stemRule
        if ((ConvertTo-HooterText $stemRule.Mode) -ne "StemCount" -or $stemExactPass.Status -ne "Pass" -or $stemExactFail.Status -ne "Fail" -or $stemQaBandPass.Status -ne "Pass" -or $stemQaBandFail.Status -ne "Fail") { throw "Regen stem count tolerance self-test failed." }
        $stemPercentRule = [pscustomobject]@{ Mode = "StemPercent"; Value = "Allowance=10; Step=5"; PointValue = "2"; CriticalFail = $false }
        $stemPercentPass = Test-HooterFieldPass -CrewValue "90" -QaValue "100" -Tolerance $stemPercentRule
        $stemPercentFail = Test-HooterFieldPass -CrewValue "5" -QaValue "7" -Tolerance $stemPercentRule
        $stemPercentScore = Get-HooterRowScoreResult -Status $stemPercentFail.Status -Tolerance $stemPercentRule -CrewValue "5" -QaValue "7"
        if ($stemPercentPass.Status -ne "Pass" -or $stemPercentFail.Status -ne "Fail" -or $stemPercentScore.LostPoints -ne 8) { throw "Regen stem percent scoring self-test failed." }
        $stemSwitchTableMode = ConvertTo-HooterText ((Get-HooterTolerance -Field "Regen|RegenMeasurements|StemCount").Mode)
        $script:UseStemCountPercentageForScoring = $true
        $stemSwitchPercentRule = Get-HooterTolerance -Field "Regen|RegenMeasurements|StemCount"
        if ($stemSwitchTableMode -ne "StemCount" -or
            (ConvertTo-HooterText $stemSwitchPercentRule.Mode) -ne "StemPercent" -or
            [string]::IsNullOrWhiteSpace((ConvertTo-HooterText $stemSwitchPercentRule.Value))) { throw "Regen stem scoring switch self-test failed." }
    }
    finally {
        $script:StemToleranceBands = @($savedStemToleranceBands | ForEach-Object { Copy-HooterStemToleranceBand -Band $_ })
        $script:UseStemCountPercentageForScoring = $savedStemPercentPreference
    }

    $filledRule = Get-DefaultToleranceForField -Group "Tree" -TableName "Trees" -FieldName "TreeNumber"
    $filledPass = Test-HooterFieldPass -CrewValue "12" -QaValue "Yes" -Tolerance $filledRule
    $filledBlankFail = Test-HooterFieldPass -CrewValue "" -QaValue "Yes" -Tolerance $filledRule
    $filledNoFail = Test-HooterFieldPass -CrewValue "12" -QaValue "No" -Tolerance $filledRule
    if ((ConvertTo-HooterText $filledRule.Mode) -ne "Filled" -or $filledPass.Status -ne "Pass" -or $filledBlankFail.Status -ne "Fail" -or $filledNoFail.Status -ne "Fail") { throw "Filled-in checkbox self-test failed." }
    $crewRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "Crew"
    $dateRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "MeasurementDate"
    $dayRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "MeasurementDay" -Label "Measurement Day"
    $remarksRule = Get-DefaultToleranceForField -Group "Plot" -TableName "PlotMeasurements" -FieldName "Remarks" -Label "Plot Remarks"
    $plotNumberRule = Get-DefaultToleranceForField -Group "Plot" -TableName "Plots" -FieldName "PlotNumber"
    if ((ConvertTo-HooterText $crewRule.Mode) -ne "Filled" -or (ConvertTo-HooterText $dateRule.Mode) -ne "Filled" -or (ConvertTo-HooterText $dayRule.Mode) -ne "Filled" -or (ConvertTo-HooterText $remarksRule.Mode) -ne "Filled" -or (ConvertTo-HooterText $plotNumberRule.Mode) -ne "Filled") { throw "Filled-in default self-test failed." }
    $filledGrid = New-HooterGrid 0 0 700 160
    Add-HooterCheckColumns -Grid $filledGrid -WithEntry $false
    Add-HooterCheckRow -Grid $filledGrid -Field ([pscustomobject]@{ FieldKey = "Plot|SelfTest|MeasurementDate"; Group = "Plot"; TableName = "SelfTest"; FieldName = "MeasurementDate"; Label = "Measurement Date" }) -CrewValue "2026-01-01"
    $filledCell = $filledGrid.Rows[0].Cells["QaValue"]
    if (-not ($filledCell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) -or $null -ne $filledCell.Value -or (ConvertTo-HooterText $filledCell.Style.NullValue) -ne "No") { throw "Filled checkbox blank display self-test failed." }
    $treePlotNumberField = [pscustomobject]@{ Group = "Tree"; TableName = "Trees"; FieldName = "PlotNumber"; Label = "Plot Number" }
    $treeRealDbhField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "RealDBH"; Label = "Real DBH" }
    $regenPlotNumberField = [pscustomobject]@{ Group = "Regen"; TableName = "RegenMeasurements"; FieldName = "PlotNumber"; Label = "Plot Number" }
    $plotPlotNumberField = [pscustomobject]@{ Group = "Plot"; TableName = "Plots"; FieldName = "PlotNumber"; Label = "Plot Number" }
    if (-not (Test-HooterHiddenTreeCheckField -Field $treePlotNumberField) -or -not (Test-HooterHiddenTreeCheckField -Field $treeRealDbhField) -or -not (Test-HooterHiddenTreeCheckField -Field $regenPlotNumberField) -or (Test-HooterHiddenTreeCheckField -Field $plotPlotNumberField)) { throw "Tree/regen hidden field self-test failed." }
    $savedToleranceMap = $script:Tolerances.Clone()
    $savedDatabaseFieldPoints = $script:DatabaseFieldPoints.Clone()
    try {
        $qaPointField = [pscustomobject]@{
            FieldKey = Get-HooterFieldKey -Group "Tree" -TableName "TreeMeasurements" -FieldName "Species"
            Group = "Tree"
            TableName = "TreeMeasurements"
            FieldName = "Species"
            Label = "Species"
            QAPoints = "7"
        }
        $script:DatabaseFieldPoints[$qaPointField.FieldKey] = "7"
        $script:Tolerances[$qaPointField.FieldKey] = [pscustomobject]@{ FieldKey = $qaPointField.FieldKey; Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "Species"; Label = "Species"; Mode = "Exact"; Value = ""; PointValue = ""; CriticalFail = $false }
        $qaPointRule = Get-HooterTolerance -Field $qaPointField
        if ((Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $qaPointRule)) -ne "7") { throw "Database QApoints fill self-test failed." }
        $script:Tolerances[$qaPointField.FieldKey].PointValue = "9"
        $qaPointOverrideRule = Get-HooterTolerance -Field $qaPointField
        if ((Format-HooterScoreNumber (Get-HooterPointValue -Tolerance $qaPointOverrideRule)) -ne "9") { throw "Editable QApoints override self-test failed." }
    }
    finally {
        $script:Tolerances = $savedToleranceMap
        $script:DatabaseFieldPoints = $savedDatabaseFieldPoints
    }

    $queryOrderSelfTest = @{
        Tree = [pscustomobject]@{ QueryName = "SelfTest"; Order = @{ species = 0; dbh = 1 } }
    }
    $dbhField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "DBH"; Label = "DBH"; FieldKey = (Get-HooterFieldKey -Group "Tree" -TableName "TreeMeasurements" -FieldName "DBH"); DefaultOrder = 1 }
    $speciesField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "Species"; Label = "Species"; FieldKey = (Get-HooterFieldKey -Group "Tree" -TableName "TreeMeasurements" -FieldName "Species"); DefaultOrder = 0 }
    $treeSpeciesLabelField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "SpeciesCode"; Label = "Tree Species"; FieldKey = (Get-HooterFieldKey -Group "Tree" -TableName "TreeMeasurements" -FieldName "SpeciesCode"); DefaultOrder = 99 }
    $sortedQueryFields = @(@($dbhField, $speciesField) | Sort-Object @{ Expression = { Get-HooterFieldQueryOrder -Field $_ -QueryOrders $queryOrderSelfTest } }, @{ Expression = { $_.FieldName } })
    if ((ConvertTo-HooterText $sortedQueryFields[0].FieldName) -ne "Species") { throw "Measurement query field order self-test failed." }
    if ((Get-HooterFieldQueryOrder -Field $treeSpeciesLabelField -QueryOrders $queryOrderSelfTest) -ne 0) { throw "Measurement query alias order self-test failed." }
    $sourceOrderDbhField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "DBH"; Label = "DBH"; FieldKey = $dbhField.FieldKey; SourceOrder = 2 }
    $sourceOrderSpeciesField = [pscustomobject]@{ Group = "Tree"; TableName = "TreeMeasurements"; FieldName = "Species"; Label = "Species"; FieldKey = $speciesField.FieldKey; SourceOrder = 1 }
    $sourceOrderFallbackFields = @(@($sourceOrderDbhField, $sourceOrderSpeciesField) | Sort-Object @{ Expression = { Get-HooterFieldQueryOrder -Field $_ -QueryOrders @{} } }, @{ Expression = { if ($null -ne $_.PSObject.Properties["SourceOrder"]) { [int]$_.SourceOrder } else { 100000 } } }, @{ Expression = { $_.FieldName } })
    if ((ConvertTo-HooterText $sourceOrderFallbackFields[0].FieldName) -ne "Species") { throw "AppColumns source order fallback self-test failed." }
    $savedFieldOrder = $script:FieldOrder
    try {
        $script:FieldOrder = @{}
        $script:FieldOrder[$dbhField.FieldKey] = 0
        $script:FieldOrder[$speciesField.FieldKey] = 1
        $customSortedFields = @(Sort-HooterFieldsBySavedOrder -Fields @($speciesField, $dbhField))
        if ((ConvertTo-HooterText $customSortedFields[0].FieldName) -ne "DBH") { throw "Saved field order self-test failed." }
        $script:FieldOrder = @{}
        $defaultSortedFields = @(Sort-HooterFieldsBySavedOrder -Fields @($dbhField, $speciesField))
        if ((ConvertTo-HooterText $defaultSortedFields[0].FieldName) -ne "Species") { throw "Default field order self-test failed." }
        $script:FieldOrder = @{}
        $importedOrderCount = Import-HooterFieldOrderRows -Rows @(
            [pscustomobject]@{ FieldKey = $speciesField.FieldKey; FieldOrder = "3" },
            [pscustomobject]@{ FieldKey = $dbhField.FieldKey; FieldOrder = "2" }
        )
        if ($importedOrderCount -ne 2 -or [int]$script:FieldOrder[$dbhField.FieldKey] -ne 2 -or [int]$script:FieldOrder[$speciesField.FieldKey] -ne 3) { throw "CSV field order import self-test failed." }
        $script:FieldOrder = @{}
        $script:FieldOrder[$dbhField.FieldKey] = 0
        $script:FieldOrder[$speciesField.FieldKey] = 1
        $moveGrid = New-HooterGrid 0 0 700 220
        Add-HooterCheckColumns -Grid $moveGrid -WithEntry $true
        Add-HooterCheckRow -Grid $moveGrid -Field $speciesField -CrewValue "WP" -EntryNumber 1 -CrewRecord "Tree 1"
        Add-HooterCheckRow -Grid $moveGrid -Field $dbhField -CrewValue "10.0" -EntryNumber 1 -CrewRecord "Tree 1"
        Add-HooterCheckRow -Grid $moveGrid -Field $speciesField -CrewValue "DF" -EntryNumber 2 -CrewRecord "Tree 2"
        Add-HooterCheckRow -Grid $moveGrid -Field $dbhField -CrewValue "11.0" -EntryNumber 2 -CrewRecord "Tree 2"
        $moveGrid.Rows[0].Cells["QaValue"].Value = "WP"
        $moveGrid.Rows[1].Cells["QaValue"].Value = "10.0"
        $moveGrid.Rows[0].Cells["Status"].Value = "Pass"
        if (-not (Apply-HooterDataEntryFieldOrderInPlace -Grid $moveGrid -Scope "Tree" -SourceKey $dbhField.FieldKey -TargetKey $speciesField.FieldKey -EntryNumber "1")) { throw "Fast field move self-test failed." }
        if ((ConvertTo-HooterText $moveGrid.Rows[0].Cells["FieldName"].Value) -ne "DBH" -or
            (ConvertTo-HooterText $moveGrid.Rows[1].Cells["FieldName"].Value) -ne "Species" -or
            (ConvertTo-HooterText $moveGrid.Rows[2].Cells["FieldName"].Value) -ne "DBH" -or
            (ConvertTo-HooterText $moveGrid.Rows[3].Cells["FieldName"].Value) -ne "Species") { throw "Fast field move order self-test failed." }
        if ((ConvertTo-HooterText $moveGrid.Rows[1].Cells["QaValue"].Value) -ne "WP" -or
            (ConvertTo-HooterText $moveGrid.Rows[1].Cells["Status"].Value) -ne "Pass") { throw "Fast field move preserve self-test failed." }
    }
    finally {
        $script:FieldOrder = $savedFieldOrder
    }

    $script:ScorePassPercent = "90"
    $script:MaxPointLoss = "25"
    $scoredPass = Get-HooterOverallStatusFromStats -Total 30 -Checked 30 -Failed 1 -EarnedPoints 29 -PossiblePoints 30 -LostPoints 1 -CriticalFailures 0
    if ($scoredPass.Status -ne "Pass") { throw "Scored pass self-test failed." }
    $pointLossFail = Get-HooterOverallStatusFromStats -Total 30 -Checked 30 -Failed 26 -EarnedPoints 4 -PossiblePoints 30 -LostPoints 26 -CriticalFailures 0
    if ($pointLossFail.Status -ne "Fail") { throw "Point-loss fail self-test failed." }
    $criticalFail = Get-HooterOverallStatusFromStats -Total 30 -Checked 1 -Failed 1 -EarnedPoints 0 -PossiblePoints 30 -LostPoints 1 -CriticalFailures 1
    if ($criticalFail.Status -ne "Fail") { throw "Critical fail self-test failed." }
    $savedPlots = $script:Plots
    $savedSessions = $script:QaSessions
    $savedTargetPercent = $script:InventoryQaTargetPercent
    $savedDatabasePath = $script:DatabasePath
    $savedProjectName = $script:ProjectName
    try {
        $script:InventoryQaTargetPercent = 5.0
        $script:DatabasePath = ""
        $script:ProjectName = ""
        $script:Plots = @(1..20 | ForEach-Object { [pscustomobject]@{ PlotNumber = $_ } })
        $script:QaSessions = New-Object System.Collections.Generic.List[object]
        [void]$script:QaSessions.Add([pscustomobject]@{ PlotNumber = "1"; Database = ""; ProjectName = "" })
        $inventoryProgressTest = Get-HooterInventoryQaProgress
        if ($inventoryProgressTest.TargetPercent -ne 10.0 -or $inventoryProgressTest.RequiredPlots -ne 2 -or $inventoryProgressTest.TargetMet) { throw "BIA inventory 10 percent target self-test failed." }
    }
    finally {
        $script:Plots = $savedPlots
        $script:QaSessions = $savedSessions
        $script:InventoryQaTargetPercent = $savedTargetPercent
        $script:DatabasePath = $savedDatabasePath
        $script:ProjectName = $savedProjectName
    }
    $savedReportSessions = $script:QaSessions
    $savedReportPlots = $script:Plots
    $savedReportProjectName = $script:ProjectName
    try {
        $script:ProjectName = "Self Test Project"
        $script:Plots = @(1..20 | ForEach-Object { [pscustomobject]@{ PlotNumber = $_ } })
        $script:QaSessions = New-Object System.Collections.Generic.List[object]
        [void]$script:QaSessions.Add([pscustomobject]@{
            SessionID = "report-pass"
            SavedAt = "2026-01-01T08:00:00"
            Database = ""
            ProjectName = "Self Test Project"
            CheckCruiserName = "Self Test Cruiser"
            CheckCruiseDate = "2026-01-02"
            PlotNumber = "1"
            UTMEasting = "500000"
            UTMNorthing = "5100000"
            UTMZone = "12N"
            OverallStatus = "Pass"
            CheckedCount = 1
            FailedCount = 0
            UncheckedCount = 0
            MissedTreeCount = 0
            ScoreEarned = "1"
            ScorePossible = "1"
            ScoreLost = "0"
            ScorePercent = "100"
            ScorePassPercent = "90"
            MaxPointLoss = "25"
            PlotPointTotal = "1"
            TreePointTotal = "0"
            RegenPointTotal = "0"
            ExecutionPointTotal = "0"
            PlotPointLoss = "0"
            TreePointLoss = "0"
            RegenPointLoss = "0"
            ExecutionPointLoss = "0"
            CriticalFailCount = 0
            OverallNotes = "Pass note"
            Rows = @([pscustomobject]@{ Scope = "Plot"; EntryNumber = "1"; CrewRecord = "1"; TableName = "Plots"; FieldName = "Crew"; FieldLabel = "Crew"; CrewValue = "ABC"; QaValue = "Yes"; RuleMode = "Filled"; ToleranceValue = ""; PointValue = "1"; CriticalFail = $false; Status = "Pass"; EarnedPoints = "1"; PossiblePoints = "1"; LostPoints = "0"; CriticalFailure = $false; FieldNotes = "" })
        })
        [void]$script:QaSessions.Add([pscustomobject]@{
            SessionID = "report-fail"
            SavedAt = "2026-01-01T09:00:00"
            Database = ""
            ProjectName = "Self Test Project"
            CheckCruiserName = "Self Test Cruiser"
            CheckCruiseDate = "2026-01-03"
            PlotNumber = "2"
            UTMEasting = "500100"
            UTMNorthing = "5100100"
            UTMZone = "12N"
            OverallStatus = "Fail"
            CheckedCount = 2
            FailedCount = 2
            UncheckedCount = 0
            MissedTreeCount = 0
            ScoreEarned = "0"
            ScorePossible = "2"
            ScoreLost = "2"
            ScorePercent = "0"
            ScorePassPercent = "90"
            MaxPointLoss = "25"
            PlotPointTotal = "0"
            TreePointTotal = "2"
            RegenPointTotal = "0"
            ExecutionPointTotal = "0"
            PlotPointLoss = "0"
            TreePointLoss = "2"
            RegenPointLoss = "0"
            ExecutionPointLoss = "0"
            CriticalFailCount = 0
            OverallNotes = "Fail note"
            Rows = @([pscustomobject]@{ Scope = "Tree"; EntryNumber = "1"; CrewRecord = "Tree 1"; TableName = "TreeMeasurements"; FieldName = "Species"; FieldLabel = "Species"; CrewValue = "RO"; QaValue = "WA"; RuleMode = "Exact"; ToleranceValue = ""; PointValue = "2"; CriticalFail = $false; Status = "Fail"; EarnedPoints = "0"; PossiblePoints = "2"; LostPoints = "2"; CriticalFailure = $false; FieldNotes = "Species mismatch" })
        })
        $reportHtml = New-HooterQaReportHtml -Sessions @($script:QaSessions.ToArray())
        if ($reportHtml -notmatch "PlotHoot QA Report" -or $reportHtml -notmatch "Failed Plots and Reasons" -or $reportHtml -notmatch "Audited Tree Data" -or $reportHtml -notmatch "Species mismatch" -or $reportHtml -notmatch "Total errors" -or $reportHtml -notmatch "Self Test Cruiser" -or $reportHtml -notmatch "2026-01-03") {
            throw "QA report HTML self-test failed."
        }
        $workbookPath = Join-Path ([System.IO.Path]::GetTempPath()) ("PlotHoot_Workbook_SelfTest_{0}.xlsx" -f ([guid]::NewGuid().ToString("N")))
        try {
            New-HooterQaWorkbookXlsx -Sessions @($script:QaSessions.ToArray()) -Path $workbookPath
            if (-not (Test-Path -LiteralPath $workbookPath) -or (Get-Item -LiteralPath $workbookPath).Length -lt 1000) {
                throw "QA workbook self-test did not create a valid file."
            }
            $zip = [System.IO.Compression.ZipFile]::OpenRead($workbookPath)
            try {
                $entry = $zip.GetEntry("xl/workbook.xml")
                if ($null -eq $entry) { throw "QA workbook self-test could not find workbook.xml." }
                $reader = New-Object System.IO.StreamReader($entry.Open())
                try { $workbookXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                foreach ($sheetName in @("Summary", "Passed Plots", "Failed Plots", "Audited Tree Data", "Setup Settings")) {
                    if ($workbookXml -notmatch [regex]::Escape($sheetName)) {
                        throw "QA workbook self-test missing sheet '$sheetName'."
                    }
                }
                $tabColorChecks = @{
                    "xl/worksheets/sheet2.xml" = "FF7E57C2"
                    "xl/worksheets/sheet3.xml" = "FFC62828"
                    "xl/worksheets/sheet4.xml" = "FFF2C94C"
                    "xl/worksheets/sheet5.xml" = "FF2E7D32"
                    "xl/worksheets/sheet6.xml" = "FF1E88E5"
                    "xl/worksheets/sheet7.xml" = "FFEF8A24"
                }
                foreach ($entryName in @($tabColorChecks.Keys)) {
                    $colorEntry = $zip.GetEntry($entryName)
                    if ($null -eq $colorEntry) { throw "QA workbook self-test could not find $entryName." }
                    $colorReader = New-Object System.IO.StreamReader($colorEntry.Open())
                    try { $colorXml = $colorReader.ReadToEnd() } finally { $colorReader.Dispose() }
                    if ($colorXml -notmatch [regex]::Escape($tabColorChecks[$entryName])) {
                        throw "QA workbook self-test missing expected tab color $($tabColorChecks[$entryName]) in $entryName."
                    }
                }
            }
            finally {
                $zip.Dispose()
            }
        }
        finally {
            if (Test-Path -LiteralPath $workbookPath) { Remove-Item -LiteralPath $workbookPath -Force }
        }
    }
    finally {
        $script:QaSessions = $savedReportSessions
        $script:Plots = $savedReportPlots
        $script:ProjectName = $savedReportProjectName
    }
    Add-HooterAssemblies
    $missedTreeTestGrid = New-HooterGrid 0 0 700 220
    Add-HooterMissedTreeColumns -Grid $missedTreeTestGrid
    $missedTreeTestCheck = New-Object System.Windows.Forms.CheckBox
    $script:Ui["MissedTreeGrid"] = $missedTreeTestGrid
    $script:Ui["CrewMissedTreeCheck"] = $missedTreeTestCheck
    $missedTreeEmptyStats = Get-HooterMissedTreeStats
    if ($missedTreeEmptyStats.CriticalFailures -ne 0) { throw "Missed tree empty self-test failed." }
    $missedTreeTestCheck.Checked = $true
    $missedTreeFlagStats = Get-HooterMissedTreeStats
    if ($missedTreeFlagStats.CriticalFailures -ne 1 -or $missedTreeFlagStats.Failed -ne 1) { throw "Missed tree checkbox self-test failed." }
    [void]$missedTreeTestGrid.Rows.Add("1", "10.2", "14", "230", "Self-test", "Critical fail")
    $missedTreeRowStats = Get-HooterMissedTreeStats
    if ($missedTreeRowStats.MissedTreeCount -ne 1 -or $missedTreeRowStats.CriticalFailures -ne 1) { throw "Missed tree row self-test failed." }
    $script:Ui.Remove("MissedTreeGrid")
    $script:Ui.Remove("CrewMissedTreeCheck")
    $treeFoundRule = Get-HooterTolerance -Field (Get-HooterTreeFoundField)
    if ((ConvertTo-HooterText $treeFoundRule.Mode) -ne "PassFail" -or -not (Get-HooterCriticalFail -Tolerance $treeFoundRule)) { throw "Tree found critical default self-test failed." }

    $savedAutoAdvanceTolerances = $script:Tolerances.Clone()
    try {
        $autoAdvanceGrid = New-HooterGrid 0 0 700 220
        Add-HooterCheckColumns -Grid $autoAdvanceGrid -WithEntry $false
        $autoFieldOne = [pscustomobject]@{ FieldKey = "Plot|SelfTest|Species"; Group = "Plot"; TableName = "SelfTest"; FieldName = "Species"; Label = "Species" }
        $autoFieldTwo = [pscustomobject]@{ FieldKey = "Plot|SelfTest|TreeNumber"; Group = "Plot"; TableName = "SelfTest"; FieldName = "TreeNumber"; Label = "Tree Number" }
        $autoFieldThree = [pscustomobject]@{ FieldKey = "Plot|SelfTest|DBH"; Group = "Plot"; TableName = "SelfTest"; FieldName = "DBH"; Label = "DBH" }
        $autoFieldFour = [pscustomobject]@{ FieldKey = "Plot|SelfTest|ObserverInitials"; Group = "Plot"; TableName = "SelfTest"; FieldName = "ObserverInitials"; Label = "Observer Initials" }
        $script:Tolerances[$autoFieldOne.FieldKey] = [pscustomobject]@{ FieldKey = $autoFieldOne.FieldKey; Group = "Plot"; TableName = "SelfTest"; FieldName = "Species"; Label = "Species"; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
        $script:Tolerances[$autoFieldTwo.FieldKey] = [pscustomobject]@{ FieldKey = $autoFieldTwo.FieldKey; Group = "Plot"; TableName = "SelfTest"; FieldName = "TreeNumber"; Label = "Tree Number"; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
        $script:Tolerances[$autoFieldThree.FieldKey] = [pscustomobject]@{ FieldKey = $autoFieldThree.FieldKey; Group = "Plot"; TableName = "SelfTest"; FieldName = "DBH"; Label = "DBH"; Mode = "Range"; Value = "0.2"; PointValue = "1"; CriticalFail = $false }
        $script:Tolerances[$autoFieldFour.FieldKey] = [pscustomobject]@{ FieldKey = $autoFieldFour.FieldKey; Group = "Plot"; TableName = "SelfTest"; FieldName = "ObserverInitials"; Label = "Observer Initials"; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
        Add-HooterCheckRow -Grid $autoAdvanceGrid -Field $autoFieldOne -CrewValue "WP"
        Add-HooterCheckRow -Grid $autoAdvanceGrid -Field $autoFieldTwo -CrewValue "12"
        Add-HooterCheckRow -Grid $autoAdvanceGrid -Field $autoFieldThree -CrewValue "10.2"
        Add-HooterCheckRow -Grid $autoAdvanceGrid -Field $autoFieldFour -CrewValue "CL"
        $autoAdvanceGrid.CurrentCell = $autoAdvanceGrid.Rows[0].Cells["QaValue"]
        if ((Get-HooterGridInputScope -Grid $autoAdvanceGrid) -ne "Number") { throw "Species numeric keyboard self-test failed." }
        $autoAdvanceGrid.CurrentCell = $autoAdvanceGrid.Rows[2].Cells["QaValue"]
        if ((Get-HooterGridInputScope -Grid $autoAdvanceGrid) -ne "Number") { throw "DBH numeric keyboard self-test failed." }
        $autoAdvanceGrid.CurrentCell = $autoAdvanceGrid.Rows[3].Cells["QaValue"]
        if ((Get-HooterGridInputScope -Grid $autoAdvanceGrid) -ne "Text") { throw "Text keyboard self-test failed." }
        if (-not (Start-HooterGridCellEdit -Grid $autoAdvanceGrid -RowIndex 0 -ColumnName "QaValue" -SelectAll)) { throw "QA single-tap edit self-test failed." }
        if ($autoAdvanceGrid.CurrentCell.RowIndex -ne 0 -or $autoAdvanceGrid.Columns[$autoAdvanceGrid.CurrentCell.ColumnIndex].Name -ne "QaValue") { throw "QA single-tap edit target self-test failed." }
        $autoAdvanceGrid.Rows[0].Cells["QaValue"].Value = "WP"
        if (-not (Focus-HooterGridEntryCell -Grid $autoAdvanceGrid -RowIndex 0 -ColumnName "QaValue")) { throw "QA auto-advance focus self-test failed." }
        if (-not (Move-HooterToNextEntryCell -Grid $autoAdvanceGrid -FromRowIndex 0 -ColumnName "QaValue")) { throw "QA auto-advance movement self-test failed." }
        if ($autoAdvanceGrid.CurrentCell.RowIndex -ne 1 -or $autoAdvanceGrid.Columns[$autoAdvanceGrid.CurrentCell.ColumnIndex].Name -ne "QaValue") { throw "QA auto-advance target self-test failed." }
    }
    finally {
        $script:Tolerances = $savedAutoAdvanceTolerances
    }

    $script:TreePointTotal = "3"
    $script:CurrentPlot = [pscustomobject]@{ TreeRecords = @(1..10 | ForEach-Object { [pscustomobject]@{ Display = "Tree $_" } }) }
    $treeScoringGrid = New-HooterGrid 0 0 700 220
    Add-HooterCheckColumns -Grid $treeScoringGrid -WithEntry $true
    $treeTestField = [pscustomobject]@{ FieldKey = "Tree|SelfTest|TreeScore"; Group = "Tree"; TableName = "SelfTest"; FieldName = "TreeScore"; Label = "Tree score" }
    $script:Tolerances[$treeTestField.FieldKey] = [pscustomobject]@{ FieldKey = $treeTestField.FieldKey; Group = "Tree"; TableName = "SelfTest"; FieldName = "TreeScore"; Label = "Tree score"; Mode = "Exact"; Value = ""; PointValue = "1"; CriticalFail = $false }
    for ($i = 0; $i -lt 30; $i++) {
        Add-HooterCheckRow -Grid $treeScoringGrid -Field $treeTestField -CrewValue "A" -EntryNumber (($i % 10) + 1) -CrewRecord ("Tree {0}" -f (($i % 10) + 1))
        $treeScoringGrid.Rows[$treeScoringGrid.Rows.Count - 1].Cells["QaValue"].Value = "B"
        $treeScoringGrid.Rows[$treeScoringGrid.Rows.Count - 1].Cells["Status"].Value = "Fail"
    }
    $treeStatsAtMax = Get-HooterGridCheckStats -Grid $treeScoringGrid -Scope "Tree"
    if ($treeStatsAtMax.PossiblePoints -ne 30 -or $treeStatsAtMax.LostPoints -ne 30 -or $treeStatsAtMax.MaxExceeded) { throw "Tree max/tree multiplier pass self-test failed." }
    Add-HooterCheckRow -Grid $treeScoringGrid -Field $treeTestField -CrewValue "A" -EntryNumber 1 -CrewRecord "Tree 1"
    $treeScoringGrid.Rows[$treeScoringGrid.Rows.Count - 1].Cells["QaValue"].Value = "B"
    $treeScoringGrid.Rows[$treeScoringGrid.Rows.Count - 1].Cells["Status"].Value = "Fail"
    $treeStatsOverMax = Get-HooterGridCheckStats -Grid $treeScoringGrid -Scope "Tree"
    if ($treeStatsOverMax.PossiblePoints -ne 30 -or $treeStatsOverMax.LostPoints -ne 31 -or -not $treeStatsOverMax.MaxExceeded) { throw "Tree max/tree multiplier fail self-test failed." }

    $script:TreePointTotal = "22"
    $script:CurrentPlot = [pscustomobject]@{ TreeRecords = @([pscustomobject]@{ Display = "Tree 1" }) }
    $treeSummaryGrid = New-HooterGrid 0 0 700 220
    Add-HooterCheckColumns -Grid $treeSummaryGrid -WithEntry $true
    $script:Ui["TreeGrid"] = $treeSummaryGrid
    Add-HooterCheckRow -Grid $treeSummaryGrid -Field $treeTestField -CrewValue "A" -EntryNumber 1 -CrewRecord "Tree 1"
    Add-HooterCheckRow -Grid $treeSummaryGrid -Field $treeTestField -CrewValue "A" -EntryNumber 2 -CrewRecord "Tree 2"
    $treeSummaryStats = Get-HooterGridCheckStats -Grid $treeSummaryGrid -Scope "Tree"
    if ($treeSummaryStats.PossiblePoints -ne 44) { throw "Tree summary max loss multiplier self-test failed." }

    $script:MaxPointLoss = ""
    $script:ExecutionPointTotal = "0"
    $plotSummaryGrid = New-HooterGrid 0 0 700 120
    Add-HooterCheckColumns -Grid $plotSummaryGrid -WithEntry $false
    $regenSummaryGrid = New-HooterGrid 0 0 700 120
    Add-HooterCheckColumns -Grid $regenSummaryGrid -WithEntry $true
    $executionSummaryGrid = New-HooterGrid 0 0 700 120
    Add-HooterExecutionColumns -Grid $executionSummaryGrid
    $errorSummaryGrid = New-HooterGrid 0 0 700 160
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "Item" -Header "Item" -Width 430 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "TotalError" -Header "Total Error" -Width 100 -ReadOnly $true)
    [void](Add-GridTextColumn -Grid $errorSummaryGrid -Name "MaxPointLoss" -Header "Max Point Loss" -Width 125 -ReadOnly $true)
    $script:Ui["PlotGrid"] = $plotSummaryGrid
    $script:Ui["RegenGrid"] = $regenSummaryGrid
    $script:Ui["ExecutionGrid"] = $executionSummaryGrid
    $script:Ui["ErrorSummaryGrid"] = $errorSummaryGrid
    Update-HooterErrorSummary
    $totalSummaryRows = @($errorSummaryGrid.Rows | Where-Object { -not $_.IsNewRow -and (ConvertTo-HooterText $_.Cells["Item"].Value) -eq "Total max point loss (section max sum)" })
    if ($totalSummaryRows.Count -ne 1 -or (ConvertTo-HooterText $totalSummaryRows[0].Cells["MaxPointLoss"].Value) -ne "44") { throw "Summary total max point loss self-test failed." }
    $thresholdRows = @($errorSummaryGrid.Rows | Where-Object { -not $_.IsNewRow -and (ConvertTo-HooterText $_.Cells["Item"].Value) -eq "Failure threshold (total error greater than this fails)" })
    if ($thresholdRows.Count -ne 1 -or (ConvertTo-HooterText $thresholdRows[0].Cells["MaxPointLoss"].Value) -ne "Not set") { throw "Summary failure threshold self-test failed." }
    $script:Ui.Remove("PlotGrid")
    $script:Ui.Remove("TreeGrid")
    $script:Ui.Remove("RegenGrid")
    $script:Ui.Remove("ExecutionGrid")
    $script:Ui.Remove("ErrorSummaryGrid")
    $script:CurrentPlot = $null

    $executionTestGrid = New-HooterGrid 0 0 700 220
    Add-HooterExecutionColumns -Grid $executionTestGrid
    $script:Ui["ExecutionGrid"] = $executionTestGrid
    $savedExecutionTolerances = $script:Tolerances.Clone()
    try {
        foreach ($field in Get-HooterExecutionSettingFields) {
            if ($script:Tolerances.ContainsKey($field.FieldKey)) { $script:Tolerances.Remove($field.FieldKey) }
        }
        Populate-HooterExecutionGrid
        $executionTestGrid.Rows[0].Cells["Rating"].Value = "Fair"
        $executionTestGrid.Rows[5].Cells["Rating"].Value = "Poor"
        Refresh-HooterExecutionGridStatuses
        $executionStats = Get-HooterExecutionGridStats
        if ($executionStats.LostPoints -ne 1 -or $executionStats.CriticalFailures -ne 1) { throw "Execution checklist self-test failed." }
        $defaultExecutionRule = Get-HooterExecutionScoreRule -ItemKey "StartingPointDistanceAzimuth"
        if ((ConvertTo-HooterText $defaultExecutionRule.ToleranceValue) -ne "Good=0; Fair=1; Poor=2") { throw "Execution Good/Fair/Poor display self-test failed." }
        $startingPointKey = Get-HooterExecutionFieldKey -ItemKey "StartingPointDistanceAzimuth"
        $script:Tolerances[$startingPointKey] = [pscustomobject]@{
            FieldKey = $startingPointKey
            Group = "Execution"
            TableName = "TableD"
            FieldName = "StartingPointDistanceAzimuth"
            Label = "Starting Point distance and azimuth provided/correct"
            Mode = "GoodFairPoor"
            Value = "Good=0; Fair=0; Poor=2"
            PointValue = "2"
            CriticalFail = $false
        }
        Update-HooterExecutionRowStatus -Row $executionTestGrid.Rows[0]
        $customExecutionResult = Get-HooterExecutionRowResult -Row $executionTestGrid.Rows[0]
        if ($customExecutionResult.Status -ne "Pass" -or $customExecutionResult.PointLoss -ne 0) { throw "Execution custom Fair pass self-test failed." }
    }
    finally {
        $script:Tolerances = $savedExecutionTolerances
        $script:Ui.Remove("ExecutionGrid")
    }

    if (-not [string]::IsNullOrWhiteSpace($Database)) {
        $connection = Open-AccessConnection -ConnectionString (New-AccessConnectionString -Path $Database)
        try {
            $catalog = Get-HooterFieldCatalog -Connection $connection
            $plots = @(Get-HooterPlots -Connection $connection)
            "Database smoke test: plots=$($plots.Count), plotFields=$($catalog.Plot.Count), treeFields=$($catalog.Tree.Count), regenFields=$($catalog.Regen.Count)"
        }
        finally {
            if ($connection.State -eq "Open") { $connection.Close() }
            $connection.Dispose()
        }
    }

    "PlotHoot self-test passed."
}

if ($SelfTest) {
    Invoke-HooterSelfTest
    return
}

if ($RepairShortcuts) {
    Repair-HooterLaunchShortcuts -CreateDesktop
    return
}

if ($NoUi) {
    return
}

Show-HooterMainForm
