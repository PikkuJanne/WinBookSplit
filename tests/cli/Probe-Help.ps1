[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$help = Get-Help -Name $env:WBS_CLI_HELP_APP -Full -ErrorAction Stop
@{
    host_version = $PSVersionTable.PSVersion.ToString()
    parameters = @($help.parameters.parameter | ForEach-Object { $_.name })
    text = ($help | Out-String -Width 180)
} | ConvertTo-Json -Depth 8 -Compress
