# Technitium DHCP GUI — WPF bulk tool (PS5-safe)
# - Add / Remove reserved leases
# - Bulk import from Excel (.xlsx/.xls) or CSV
# - Dry Run validation (MAC, IP, IP inside scope Start–End)
# - Results grid + Save CSV
# - "List Leases (quick)" populates the grid with live leases from server
# Endpoints used (token in query string):
#   /api/dhcp/scopes/list
#   /api/dhcp/scopes/get?name={scope}
#   /api/dhcp/scopes/addReservedLease
#   /api/dhcp/scopes/removeReservedLease
#   /api/dhcp/leases/list

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ---------------- XAML UI ----------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Technitium DHCP – Reserved Leases" Height="780" Width="1120" WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="2*"/>
      <ColumnDefinition Width="2*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <!-- Connection -->
    <StackPanel Grid.Row="0" Grid.ColumnSpan="3" Orientation="Horizontal" Margin="0,0,0,10">
      <StackPanel Width="420" Margin="0,0,20,0">
        <TextBlock Text="Base URL (e.g. http://38.80.65.250:5380)" />
        <TextBox Name="BaseBox" />
      </StackPanel>
      <StackPanel Width="420" Margin="0,0,20,0">
        <TextBlock Text="API Token (passed in URL)" />
        <TextBox Name="TokenBox" />
      </StackPanel>
      <StackPanel Width="320">
        <TextBlock Text="Scope Name (exact, e.g. Public IP Space)" />
        <TextBox Name="ScopeBox" />
      </StackPanel>
      <StackPanel Margin="20,16,0,0" Orientation="Horizontal" VerticalAlignment="Bottom">
        <CheckBox Name="DryRunChk" Content="Dry Run (validate only)" Margin="0,0,12,0"/>
        <Button Name="TestConnBtn" Width="140" Height="28" Margin="0,0,8,0">Test Connection</Button>
        <Button Name="ListBtn" Width="170" Height="28">List Leases (quick)</Button>
      </StackPanel>
    </StackPanel>

    <!-- Single add/remove -->
    <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Vertical" Margin="0,0,10,10">
      <TextBlock FontWeight="Bold" Text="Single Entry" Margin="0,0,0,8"/>
      <Grid Margin="0,0,0,10">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="160"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Grid.Column="0" Text="Hostname"/>
        <TextBox   Grid.Row="0" Grid.Column="1" Name="HostBox"/>

        <TextBlock Grid.Row="1" Grid.Column="0" Text="MAC (AA:BB:CC:DD:EE:FF)"/>
        <TextBox   Grid.Row="1" Grid.Column="1" Name="MacBox"/>

        <TextBlock Grid.Row="2" Grid.Column="0" Text="IP Address"/>
        <TextBox   Grid.Row="2" Grid.Column="1" Name="IpBox"/>

        <TextBlock Grid.Row="3" Grid.Column="0" Text="Comment (optional)"/>
        <TextBox   Grid.Row="3" Grid.Column="1" Name="CommentBox"/>
      </Grid>

      <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
        <Button Name="AddBtn" Width="160" Height="32" Margin="0,0,10,0">Add Reserved Lease</Button>
        <Button Name="DelBtn" Width="180" Height="32">Remove Reserved Lease</Button>
      </StackPanel>
    </StackPanel>

    <!-- Bulk import -->
    <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Vertical" Margin="10,0,10,10">
      <TextBlock FontWeight="Bold" Text="Bulk Import" Margin="0,0,0,8"/>
      <TextBlock Text="Excel columns accepted (any case): name/scope, hardwareAddress/mac, ipAddress/ip, hostName/hostname, comments/comment" Margin="0,0,0,6"/>
      <TextBlock Text="CSV supported too (same headers). First sheet used for Excel." Margin="0,0,0,10"/>
      <StackPanel Orientation="Horizontal">
        <Button Name="PickFileBtn" Width="160" Height="32" Margin="0,0,10,0">Pick Excel/CSV...</Button>
        <TextBlock Name="PickedFileText" VerticalAlignment="Center"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
        <Button Name="ImportBtn" Width="160" Height="32" IsEnabled="False">Import (Bulk)</Button>
        <Button Name="SaveCsvBtn" Width="160" Height="32" Margin="10,0,0,0">Save Results → CSV</Button>
      </StackPanel>
    </StackPanel>

    <!-- Results grid -->
    <DataGrid Grid.Row="2" Grid.ColumnSpan="3" Name="ResultsGrid" AutoGenerateColumns="True" IsReadOnly="True" Margin="0,10,0,10"
              AlternatingRowBackground="#FFF5F5F5" HeadersVisibility="Column" />

    <!-- Log -->
    <TextBox Grid.Row="3" Grid.ColumnSpan="3" Name="LogBox" Margin="0,10,0,0"
             Height="180" VerticalScrollBarVisibility="Auto" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap"/>
  </Grid>
</Window>
"@

# --------------- Load XAML ---------------
[xml]$xml = $xaml
$reader  = (New-Object System.Xml.XmlNodeReader $xml)
$window  = [Windows.Markup.XamlReader]::Load($reader)

# Elements
$BaseBox   = $window.FindName('BaseBox')
$TokenBox  = $window.FindName('TokenBox')
$ScopeBox  = $window.FindName('ScopeBox')
$HostBox   = $window.FindName('HostBox')
$MacBox    = $window.FindName('MacBox')
$IpBox     = $window.FindName('IpBox')
$CommentBox= $window.FindName('CommentBox')
$AddBtn    = $window.FindName('AddBtn')
$DelBtn    = $window.FindName('DelBtn')
$PickFileBtn = $window.FindName('PickFileBtn')
$PickedFileText = $window.FindName('PickedFileText')
$ImportBtn = $window.FindName('ImportBtn')
$TestConnBtn = $window.FindName('TestConnBtn')
$ListBtn   = $window.FindName('ListBtn')
$LogBox    = $window.FindName('LogBox')
$DryRunChk = $window.FindName('DryRunChk')
$ResultsGrid = $window.FindName('ResultsGrid')
$SaveCsvBtn  = $window.FindName('SaveCsvBtn')

function Log([string]$msg){ $LogBox.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $msg`r`n"); $LogBox.ScrollToEnd() }

# --------------- Results collection ---------------
$Results = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ResultsGrid.ItemsSource = $Results
function Add-Result {
  param([string]$Action,[string]$Scope,[string]$Hostname,[string]$MAC,[string]$IP,[string]$Result,[string]$Message)
  $Results.Add([pscustomobject]@{
    Time    = (Get-Date).ToString("HH:mm:ss")
    Action  = $Action
    Scope   = $Scope
    Host    = $Hostname
    MAC     = $MAC
    IP      = $IP
    Result  = $Result
    Message = $Message
  })
}

# NEW: Fill grid with live leases
function Show-LeasesInGrid {
  param($leases)
  $Results.Clear()
  foreach($l in $leases){
    $Results.Add([pscustomobject]@{
      Time     = (Get-Date).ToString("HH:mm:ss")
      Action   = "Lease"
      Scope    = $l.scope
      Type     = $l.type
      MAC      = $l.hardwareAddress
      IP       = $l.address
      Host     = $l.hostName
      Obtained = $l.leaseObtained
      Expires  = $l.leaseExpires
      Result   = ""
      Message  = ""
    })
  }
}

# --------------- HTTP Helpers ---------------
function Get-ScopeDetails {
  param([string]$Base,[string]$Token,[string]$ScopeName)
  try {
    Invoke-RestMethod -Method Get -Uri "$Base/api/dhcp/scopes/get?token=$Token&name=$([uri]::EscapeDataString($ScopeName))" -ErrorAction Stop
  } catch { return $null }
}

function Add-ReservedLease {
  param([string]$Base,[string]$Token,[string]$ScopeName,[string]$Mac,[string]$Ip,[string]$Hostname,[string]$Comment,[bool]$DryRun=$false)

  # Validation
  $macOk = $Mac -match '^(?:[0-9A-Fa-f]{2}([-:]))(?:[0-9A-Fa-f]{2}\1){4}[0-9A-Fa-f]{2}$'
  if(-not $macOk){ return @{ok=$false; err="Invalid MAC format"} }

  $ipObj = $null; [void][System.Net.IPAddress]::TryParse($Ip,[ref]$ipObj)
  if($ipObj -eq $null){ return @{ok=$false; err="Invalid IP address"} }

  $scopeDet = Get-ScopeDetails -Base $Base -Token $Token -ScopeName $ScopeName
  if($scopeDet -and $scopeDet.response){
    $start = $scopeDet.response.startingAddress
    $end   = $scopeDet.response.endingAddress
    if($start -and $end){
      if(-not (Test-IpInRange -Ip $Ip -Start $start -End $end)){
        return @{ok=$false; err=("IP not in scope range (" + $start + " - " + $end + ")")}
      }
    }
  }

  if($DryRun){ return @{ok=$true; data=@{dryRun=$true}} }

  $uri = "$Base/api/dhcp/scopes/addReservedLease?token=$Token&name=$([uri]::EscapeDataString($ScopeName))&hardwareAddress=$Mac&ipAddress=$Ip"
  if($Hostname){ $uri += "&hostName=$([uri]::EscapeDataString($Hostname))" }
  if($Comment){  $uri += "&comments=$([uri]::EscapeDataString($Comment))" }

  try {
    $r = Invoke-RestMethod -Method Post -Uri $uri -ErrorAction Stop
    return @{ok=$true; data=$r}
  } catch { return @{ok=$false; err=$_.Exception.Message} }
}

function Remove-ReservedLease {
  param([string]$Base,[string]$Token,[string]$ScopeName,[string]$Mac,[bool]$DryRun=$false)

  $macOk = $Mac -match '^(?:[0-9A-Fa-f]{2}([-:]))(?:[0-9A-Fa-f]{2}\1){4}[0-9A-Fa-f]{2}$'
  if(-not $macOk){ return @{ok=$false; err="Invalid MAC format"} }
  if($DryRun){ return @{ok=$true; data=@{dryRun=$true}} }

  $uri = "$Base/api/dhcp/scopes/removeReservedLease?token=$Token&name=$([uri]::EscapeDataString($ScopeName))&hardwareAddress=$Mac"
  try {
    $r = Invoke-RestMethod -Method Post -Uri $uri -ErrorAction Stop
    return @{ok=$true; data=$r}
  } catch { return @{ok=$false; err=$_.Exception.Message} }
}

function List-Leases {
  param([string]$Base,[string]$Token)
  try {
    Invoke-RestMethod -Method Get -Uri "$Base/api/dhcp/leases/list?token=$Token" -ErrorAction Stop
  } catch { @{ response = @{}; status = ("error: " + $_.Exception.Message) } }
}

function Test-ConnectionApi {
  param([string]$Base,[string]$Token)
  try {
    Invoke-RestMethod -Method Get -Uri "$Base/api/dhcp/scopes/list?token=$Token" -ErrorAction Stop
  } catch { @{ response = @{}; status = ("error: " + $_.Exception.Message) } }
}

# --------------- IP helpers ---------------
function IPToUInt32([string]$ip){
  $addr = [System.Net.IPAddress]::Parse($ip).GetAddressBytes()
  [Array]::Reverse($addr)
  return [System.BitConverter]::ToUInt32($addr,0)
}
function Test-IpInRange {
  param([string]$Ip,[string]$Start,[string]$End)
  try {
    $x = IPToUInt32 $Ip
    $a = IPToUInt32 $Start
    $b = IPToUInt32 $End
    return ($x -ge $a -and $x -le $b)
  } catch { return $false }
}

# --------------- Excel/CSV Loader ---------------
function Load-TableFromFile {
  param([string]$Path)

  $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
  if($ext -eq ".csv"){
    try { return (Import-Csv -Path $Path) }
    catch { throw ("CSV read error: " + $_.Exception.Message) }
  }

  if($ext -in @(".xlsx",".xls")){
    try {
      $excel = New-Object -ComObject Excel.Application
    } catch {
      throw "Excel is not installed (needed to read .xlsx/.xls). Either install Excel or save the file as CSV."
    }
    try {
      $excel.Visible = $false
      $wb = $excel.Workbooks.Open($Path)
      $ws = $wb.Worksheets.Item(1)
      $used = $ws.UsedRange
      $vals = $used.Value2
      $headers = @()
      for($c=1; $c -le $vals.GetLength(1); $c++){ $headers += [string]$vals[1,$c] }

      $rows = @()
      for($r=2; $r -le $vals.GetLength(0); $r++){
        $obj = [ordered]@{}
        for($c=1; $c -le $vals.GetLength(1); $c++){
          $obj[$headers[$c-1]] = $vals[$r,$c]
        }
        $rows += (New-Object psobject -Property $obj)
      }
      $wb.Close($false)
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)   | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null

      return $rows
    } catch {
      try { $excel.Quit() | Out-Null } catch {}
      throw ("Excel import error: " + $_.Exception.Message)
    }
  }

  throw ("Unsupported file type: " + $ext + ". Use .xlsx, .xls, or .csv")
}

function Normalize-Row {
  param($row)
  # Accept multiple header variants
  $scope   = $row.name; if(-not $scope){ $scope = $row.scope }
  $mac     = $row.hardwareAddress; if(-not $mac){ $mac = $row.mac }
  $ip      = $row.ipAddress; if(-not $ip){ $ip = $row.ip }
  $hostname= $row.hostName; if(-not $hostname){ $hostname = $row.hostname }
  $comment = $row.comments; if(-not $comment){ $comment = $row.comment }

  # Trim
  foreach($n in 'scope','mac','ip','hostname','comment'){
    $val = (Get-Variable $n -ValueOnly -ErrorAction SilentlyContinue)
    if($null -ne $val){ Set-Variable -Name $n -Value ([string]$val).Trim() }
  }

  [pscustomobject]@{ scope=$scope; mac=$mac; ip=$ip; host=$hostname; comment=$comment }
}

# --------------- Button Actions ---------------
$AddBtn.Add_Click({
  $base  = $BaseBox.Text;  $token = $TokenBox.Text; $scope = $ScopeBox.Text
  $mac   = $MacBox.Text;   $ip    = $IpBox.Text;    $hostname = $HostBox.Text; $com = $CommentBox.Text
  $dry   = [bool]$DryRunChk.IsChecked

  if([string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($scope) -or
     [string]::IsNullOrWhiteSpace($mac)  -or [string]::IsNullOrWhiteSpace($ip)){
    Log "Missing required fields (Base/Token/Scope/MAC/IP)."; return
  }

  $res = Add-ReservedLease -Base $base -Token $token -ScopeName $scope -Mac $mac -Ip $ip -Hostname $hostname -Comment $com -DryRun:$dry
  if($res.ok){
    $msg=""; $resultText=""
    if($dry){ $msg="VALID (dry run)"; $resultText="Valid" } else { $msg="ADD OK"; $resultText="Added" }
    Log "$($msg): $hostname ($mac) -> $ip"
    Add-Result -Action "Add" -Scope $scope -Host $hostname -MAC $mac -IP $ip -Result $resultText -Message ""
  } else {
    Log ("ADD FAIL: " + $res.err)
    Add-Result -Action "Add" -Scope $scope -Host $hostname -MAC $mac -IP $ip -Result "Error" -Message $res.err
  }
})

$DelBtn.Add_Click({
  $base  = $BaseBox.Text;  $token = $TokenBox.Text; $scope = $ScopeBox.Text
  $mac   = $MacBox.Text;   $dry   = [bool]$DryRunChk.IsChecked
  if([string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($scope) -or [string]::IsNullOrWhiteSpace($mac)){
    Log "Need Base/Token/Scope/MAC to delete."; return
  }
  $res = Remove-ReservedLease -Base $base -Token $token -ScopeName $scope -Mac $mac -DryRun:$dry
  if($res.ok){
    $msg=""; $resultText=""
    if($dry){ $msg="VALID (dry run)"; $resultText="Valid" } else { $msg="DELETE OK"; $resultText="Deleted" }
    Log "$($msg): $mac from scope '$scope'"
    Add-Result -Action "Delete" -Scope $scope -Host "" -MAC $mac -IP "" -Result $resultText -Message ""
  } else {
    Log ("DELETE FAIL: " + $res.err)
    Add-Result -Action "Delete" -Scope $scope -Host "" -MAC $mac -IP "" -Result "Error" -Message $res.err
  }
})

$PickFileBtn.Add_Click({
  $ofd = New-Object Microsoft.Win32.OpenFileDialog
  $ofd.Filter = "Excel/CSV|*.xlsx;*.xls;*.csv|All files|*.*"
  if($ofd.ShowDialog()){
    $PickedFileText.Text = $ofd.FileName
    $ImportBtn.IsEnabled = $true
    Log ("Selected file: " + $ofd.FileName)
  }
})

$ImportBtn.Add_Click({
  $base   = $BaseBox.Text
  $token  = $TokenBox.Text
  $uiScope= $ScopeBox.Text
  $dry    = [bool]$DryRunChk.IsChecked

  if([string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($token)){
    Log "Base URL and Token are required before import."; return
  }

  $path = $PickedFileText.Text
  if(-not (Test-Path $path)){ Log "Pick a file first."; return }

  try { $raw = Load-TableFromFile -Path $path } catch { Log $_; return }

  $count=0; $ok=0; $fail=0
  foreach($r in $raw){
    $n = Normalize-Row $r
    if([string]::IsNullOrWhiteSpace($n.scope)){ $scope = $uiScope } else { $scope = $n.scope }

    if([string]::IsNullOrWhiteSpace($scope) -or [string]::IsNullOrWhiteSpace($n.mac) -or [string]::IsNullOrWhiteSpace($n.ip)){
      $msg="Missing scope/mac/ip"
      Log "Skipped: $($msg) :: $($n | ConvertTo-Json -Compress)"
      Add-Result -Action "Add" -Scope $scope -Host $n.host -MAC $n.mac -IP $n.ip -Result "Skipped" -Message $msg
      $fail++; continue
    }

    $res = Add-ReservedLease -Base $base -Token $token -ScopeName $scope -Mac $n.mac -Ip $n.ip -Hostname $n.host -Comment $n.comment -DryRun:$dry
    if($res.ok){
      if($dry){ $msg="VALID (dry run)"; $resultText="Valid" } else { $msg="ADD OK"; $resultText="Added" }
      Log "$($msg): $($n.host) ($($n.mac)) -> $($n.ip) [scope: $scope]"
      Add-Result -Action "Add" -Scope $scope -Host $n.host -MAC $n.mac -IP $n.ip -Result $resultText -Message ""
      $ok++
    } else {
      Log ("ADD FAIL: " + $n.mac + " -> " + $n.ip + " :: " + $res.err)
      Add-Result -Action "Add" -Scope $scope -Host $n.host -MAC $n.mac -IP $n.ip -Result "Error" -Message $res.err
      $fail++
    }
    $count++
  }

  if($dry){ $modeText="Dry Run" } else { $modeText="Live" }
  Log ("Import done. Processed: " + $count + " | OK: " + $ok + " | Fail: " + $fail + " | Mode: " + $modeText)
})

$SaveCsvBtn.Add_Click({
  if($Results.Count -eq 0){ Log "No results to save."; return }
  $sfd = New-Object Microsoft.Win32.SaveFileDialog
  $sfd.Filter = "CSV files (*.csv)|*.csv"
  $sfd.FileName = "Technitium_Results_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
  if($sfd.ShowDialog()){
    $path = $sfd.FileName
    try {
      $Results | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
      Log ("Saved results to " + $path)
    } catch { Log ("Failed to save CSV: " + $_.Exception.Message) }
  }
})

$TestConnBtn.Add_Click({
  $res = Test-ConnectionApi -Base $BaseBox.Text -Token $TokenBox.Text
  $txt = ($res | ConvertTo-Json -Depth 6)
  Log ("Test Connection result: " + $txt)
})

$ListBtn.Add_Click({
  $base  = $BaseBox.Text
  $token = $TokenBox.Text

  if([string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($token)){
    Log "Base URL and Token are required."
    return
  }

  try {
    $res = List-Leases -Base $base -Token $token
    if($null -eq $res -or $null -eq $res.response -or $null -eq $res.response.leases){
      Log "No leases found."
      return
    }

    $leases = $res.response.leases

    # To show only reserved leases, uncomment:
    # $leases = $leases | Where-Object { $_.type -eq 'Reserved' }

    Show-LeasesInGrid -leases $leases
    Log ("Leases loaded: " + ($leases | Measure-Object).Count)
  } catch {
    Log ("List leases failed: " + $_.Exception.Message)
  }
})

# Show
$window.ShowDialog() | Out-Null
