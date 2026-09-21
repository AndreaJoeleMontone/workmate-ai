param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("status","metadata","query")]
    [string]$Mode,
    [string]$QueryFile = "",
    [string]$ReportName = "YourReport.pbix"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Find-Dscmd {
    $command = Get-Command "dscmd.exe" -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $candidates = @(
        "$env:ProgramFiles\DAX Studio\dscmd.exe",
        "${env:ProgramFiles(x86)}\DAX Studio\dscmd.exe",
        "$env:LOCALAPPDATA\Programs\DAX Studio\dscmd.exe",
        "$env:LOCALAPPDATA\DAX Studio\dscmd.exe"
    )

    foreach ($path in $candidates) {
        if ($path -and (Test-Path $path)) { return $path }
    }

    foreach ($root in @(
        (Join-Path $env:USERPROFILE "Downloads"),
        (Join-Path $env:USERPROFILE "Desktop")
    )) {
        if (Test-Path $root) {
            $found = Get-ChildItem $root -Recurse -Filter "dscmd.exe" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) { return $found.FullName }
        }
    }

    throw "DAX Studio dscmd.exe was not found."
}

function Invoke-DscmdRows {
    param([string]$Dscmd,[string]$Server,[string]$Dax)

    $temp = Join-Path $env:TEMP ("WorkMate_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $temp -Force | Out-Null
    $query = Join-Path $temp "query.dax"
    $output = Join-Path $temp "result.csv"

    try {
        [System.IO.File]::WriteAllText($query,$Dax,(New-Object System.Text.UTF8Encoding($false)))
        $console = @(& $Dscmd csv $output --server $Server --file $query --filetype UTF8CSV 2>&1)

        if ($LASTEXITCODE -ne 0) {
            throw (($console | ForEach-Object { "$_" }) -join [Environment]::NewLine)
        }
        if (-not (Test-Path $output)) {
            throw "DAX Studio did not create the CSV result."
        }

        $first = Get-Content $output -Encoding UTF8 | Select-Object -First 1
        if (-not $first) { return @() }

        if ($first -match ';') { $delimiter = ';' } else { $delimiter = ',' }
        return @(Import-Csv -Path $output -Encoding UTF8 -Delimiter $delimiter)
    }
    finally {
        Remove-Item $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Server-Candidates {
    param([string]$Name)
    $result = New-Object System.Collections.Generic.List[string]

    if ($Name) {
        $result.Add($Name)
        if (-not $Name.EndsWith(".pbix")) { $result.Add("$Name.pbix") }
        if ($Name.EndsWith(".pbix")) {
            $result.Add([System.IO.Path]::GetFileNameWithoutExtension($Name))
        }
    }
    return @($result)
}

function Find-Server {
    param([string]$Dscmd,[string]$Name)
    $errors = @()

    foreach ($candidate in (Server-Candidates $Name)) {
        try {
            Invoke-DscmdRows -Dscmd $Dscmd -Server $candidate -Dax 'EVALUATE ROW("ConnectionTest",1)' | Out-Null
            return $candidate
        }
        catch {
            $errors += "$candidate => $($_.Exception.Message)"
        }
    }

    throw "Could not connect to an open Power BI report. " + ($errors -join " | ")
}

$dscmd = Find-Dscmd
$server = Find-Server -Dscmd $dscmd -Name $ReportName

if ($Mode -eq "status") {
    [pscustomobject]@{
        ok = $true
        provider = "DAX Studio DSCMD"
        database = $server
        report = $server
        dscmd = $dscmd
    } | ConvertTo-Json -Compress
    exit 0
}

if ($Mode -eq "metadata") {
    $tablesDax = @'
EVALUATE
SELECTCOLUMNS(
    INFO.VIEW.TABLES(),
    "Name", [Name],
    "IsHidden", [IsHidden]
)
ORDER BY [Name]
'@

    $columnsDax = @'
EVALUATE
SELECTCOLUMNS(
    INFO.VIEW.COLUMNS(),
    "TableName", [Table],
    "Name", [Name],
    "DataType", [DataType],
    "IsHidden", [IsHidden]
)
ORDER BY [TableName], [Name]
'@

    $measuresDax = @'
EVALUATE
SELECTCOLUMNS(
    INFO.VIEW.MEASURES(),
    "TableName", [Table],
    "Name", [Name],
    "Expression", [Expression],
    "IsHidden", [IsHidden]
)
ORDER BY [TableName], [Name]
'@

    $tables = @(Invoke-DscmdRows -Dscmd $dscmd -Server $server -Dax $tablesDax)
    $columns = @(Invoke-DscmdRows -Dscmd $dscmd -Server $server -Dax $columnsDax)
    $measures = @(Invoke-DscmdRows -Dscmd $dscmd -Server $server -Dax $measuresDax)

    [pscustomobject]@{
        ok = $true
        provider = "DAX Studio DSCMD"
        database = $server
        tables = $tables
        columns = $columns
        measures = $measures
    } | ConvertTo-Json -Depth 12 -Compress
    exit 0
}

if ($Mode -eq "query") {
    if (-not $QueryFile -or -not (Test-Path $QueryFile)) {
        throw "Query file was not found."
    }

    $dax = Get-Content $QueryFile -Raw -Encoding UTF8
    $rows = @(Invoke-DscmdRows -Dscmd $dscmd -Server $server -Dax $dax)

    [pscustomobject]@{
        ok = $true
        database = $server
        row_count = $rows.Count
        rows = $rows
    } | ConvertTo-Json -Depth 12 -Compress
    exit 0
}
