# Compatible builder for older PS2EXE versions (no -ProductVersion).
# If unsure of supported params, run: Get-Command Invoke-ps2exe -Syntax
$In = Join-Path $PSScriptRoot 'TechnitiumDHCP-GUI_WPF_Bulk_Validate_fixed.ps1'
$Out = Join-Path $PSScriptRoot 'TechnitiumDHCP-GUI_WPF_Bulk_Validate.exe'

Import-Module PS2EXE -ErrorAction Stop
Invoke-PS2EXE -InputFile $In -OutputFile $Out -NoConsole -STA -Title 'Technitium DHCP Bulk Tool' -Company 'Circle B Wireless' -Product 'Technitium DHCP GUI'
Write-Host "Built: $Out"
