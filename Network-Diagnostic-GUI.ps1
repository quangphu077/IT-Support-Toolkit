# Network Diagnostic Tool - GUI Edition
# Compatible with Windows PowerShell 5.1 and PowerShell 7 on Windows.
# Run from PowerShell:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Network-Diagnostic-GUI.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-NetworkDiagnosticReport {
    $lines = New-Object System.Collections.Generic.List[string]
    $now = Get-Date

    function Add-Title {
        param([string]$Title)
        $lines.Add("")
        $lines.Add(("=" * 72))
        $lines.Add($Title)
        $lines.Add(("=" * 72))
    }

    function Add-Text {
        param([string]$Text)
        $lines.Add($Text)
    }

    Add-Title "NETWORK DIAGNOSTIC REPORT"
    Add-Text "Computer Name : $env:COMPUTERNAME"
    Add-Text "User          : $env:USERDOMAIN\$env:USERNAME"
    Add-Text "Date          : $($now.ToString('yyyy-MM-dd HH:mm:ss'))"

    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        Add-Text "Operating Sys : $($os.Caption)"
        Add-Text "OS Version    : $($os.Version)"
    } catch {
        Add-Text "Operating Sys : Unable to retrieve ($($_.Exception.Message))"
    }

    Add-Title "NETWORK ADAPTERS"
    try {
        $adapters = @(Get-NetAdapter -ErrorAction Stop)
        if ($adapters.Count -eq 0) {
            Add-Text "No network adapters found."
        } else {
            foreach ($adapter in $adapters) {
                Add-Text "Name        : $($adapter.Name)"
                Add-Text "Description : $($adapter.InterfaceDescription)"
                Add-Text "Status      : $($adapter.Status)"
                Add-Text "Link Speed  : $($adapter.LinkSpeed)"
                Add-Text "MAC Address : $($adapter.MacAddress)"
                Add-Text "Interface   : $($adapter.ifIndex)"
                Add-Text ("-" * 50)
            }
        }
    } catch {
        Add-Text "Could not read adapters: $($_.Exception.Message)"
    }

    Add-Title "IP CONFIGURATION"
    $activeConfigs = @()
    try {
        $allConfigs = @(Get-NetIPConfiguration -ErrorAction Stop)
        foreach ($config in $allConfigs) {
            Add-Text "Interface : $($config.InterfaceAlias)"
            Add-Text "IPv4      : $(($config.IPv4Address | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', ')"
            Add-Text "IPv6      : $(($config.IPv6Address | ForEach-Object { $_.IPAddress }) -join ', ')"
            Add-Text "Gateway   : $(($config.IPv4DefaultGateway | ForEach-Object { $_.NextHop }) -join ', ')"
            Add-Text "DNS       : $(($config.DNSServer.ServerAddresses) -join ', ')"
            Add-Text ("-" * 50)

            if ($config.NetAdapter.Status -eq "Up" -and $config.IPv4Address) {
                $activeConfigs += $config
            }
        }
    } catch {
        Add-Text "Could not read IP configuration: $($_.Exception.Message)"
    }

    Add-Title "DHCP INFORMATION"
    try {
        $dhcpItems = @(Get-CimInstance Win32_NetworkAdapterConfiguration -ErrorAction Stop |
            Where-Object { $_.IPEnabled })
        foreach ($item in $dhcpItems) {
            Add-Text "Adapter     : $($item.Description)"
            Add-Text "DHCP Enabled: $($item.DHCPEnabled)"
            Add-Text "DHCP Server : $($item.DHCPServer)"
            Add-Text "Lease Start : $($item.DHCPLeaseObtained)"
            Add-Text "Lease End   : $($item.DHCPLeaseExpires)"
            Add-Text ("-" * 50)
        }
        if ($dhcpItems.Count -eq 0) { Add-Text "No active IP-enabled adapters found." }
    } catch {
        Add-Text "Could not read DHCP information: $($_.Exception.Message)"
    }

    Add-Title "CONNECTIVITY TESTS"

    $gateways = @($activeConfigs |
        ForEach-Object { $_.IPv4DefaultGateway | ForEach-Object { $_.NextHop } } |
        Where-Object { $_ } | Select-Object -Unique)

    if ($gateways.Count -gt 0) {
        foreach ($gateway in $gateways) {
            try {
                $ok = Test-Connection -ComputerName $gateway -Count 2 -Quiet -ErrorAction SilentlyContinue
                Add-Text "Gateway $gateway : $(if ($ok) { 'PASS - reachable' } else { 'FAIL - no reply' })"
            } catch {
                Add-Text "Gateway $gateway : ERROR - $($_.Exception.Message)"
            }
        }
    } else {
        Add-Text "Default Gateway : NOT CONFIGURED"
    }

    try {
        $internetOk = Test-Connection -ComputerName "1.1.1.1" -Count 2 -Quiet -ErrorAction SilentlyContinue
        Add-Text "Internet test (1.1.1.1): $(if ($internetOk) { 'PASS - reachable' } else { 'FAIL - no reply' })"
    } catch {
        Add-Text "Internet test (1.1.1.1): ERROR - $($_.Exception.Message)"
    }

    try {
        $dnsResult = Resolve-DnsName -Name "www.microsoft.com" -Type A -ErrorAction Stop |
            Where-Object { $_.IPAddress } | Select-Object -First 3 -ExpandProperty IPAddress
        Add-Text "DNS resolution (www.microsoft.com): PASS"
        Add-Text "Resolved IP(s): $($dnsResult -join ', ')"
    } catch {
        Add-Text "DNS resolution (www.microsoft.com): FAIL - $($_.Exception.Message)"
    }

    Add-Title "IPCONFIG /ALL (REFERENCE)"
    try {
        $ipconfigOutput = & ipconfig.exe /all 2>&1
        foreach ($line in $ipconfigOutput) { Add-Text ([string]$line) }
    } catch {
        Add-Text "ipconfig command failed: $($_.Exception.Message)"
    }

    Add-Text ""
    Add-Text "End of report."
    return ($lines -join [Environment]::NewLine)
}

# ---------------- GUI ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Network Diagnostic Tool"
$form.Size = New-Object System.Drawing.Size(1000, 720)
$form.MinimumSize = New-Object System.Drawing.Size(800, 560)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$header = New-Object System.Windows.Forms.Label
$header.Text = "NETWORK DIAGNOSTIC TOOL"
$header.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$header.ForeColor = [System.Drawing.Color]::FromArgb(31, 78, 121)
$header.AutoSize = $true
$header.Location = New-Object System.Drawing.Point(20, 16)
$form.Controls.Add($header)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Collect adapter, IP, gateway, DNS, DHCP and connectivity information"
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(22, 49)
$subtitle.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($subtitle)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Diagnostic"
$runButton.Location = New-Object System.Drawing.Point(22, 82)
$runButton.Size = New-Object System.Drawing.Size(145, 38)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(31, 119, 180)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = "Flat"
$runButton.FlatAppearance.BorderSize = 0
$form.Controls.Add($runButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export Report"
$exportButton.Location = New-Object System.Drawing.Point(178, 82)
$exportButton.Size = New-Object System.Drawing.Size(145, 38)
$exportButton.Enabled = $false
$exportButton.BackColor = [System.Drawing.Color]::FromArgb(40, 140, 90)
$exportButton.ForeColor = [System.Drawing.Color]::White
$exportButton.FlatStyle = "Flat"
$exportButton.FlatAppearance.BorderSize = 0
$form.Controls.Add($exportButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear"
$clearButton.Location = New-Object System.Drawing.Point(334, 82)
$clearButton.Size = New-Object System.Drawing.Size(90, 38)
$clearButton.FlatStyle = "Flat"
$form.Controls.Add($clearButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Ready. Click Run Diagnostic to start."
$statusLabel.Location = New-Object System.Drawing.Point(22, 132)
$statusLabel.Size = New-Object System.Drawing.Size(930, 22)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$form.Controls.Add($statusLabel)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(22, 160)
$outputBox.Size = New-Object System.Drawing.Size(940, 490)
$outputBox.Anchor = "Top, Bottom, Left, Right"
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Both"
$outputBox.WordWrap = $false
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($outputBox)

$script:LatestReport = ""

$runButton.Add_Click({
    $runButton.Enabled = $false
    $exportButton.Enabled = $false
    $statusLabel.Text = "Status: Collecting network information..."
    $outputBox.Text = "Running diagnostics. Please wait..."
    $form.Refresh()

    try {
        $script:LatestReport = Get-NetworkDiagnosticReport
        $outputBox.Text = $script:LatestReport
        $statusLabel.Text = "Status: Diagnostic completed. Review the report, then click Export Report."
        $exportButton.Enabled = $true
    } catch {
        $outputBox.Text = "An unexpected error occurred:`r`n$($_.Exception.Message)"
        $statusLabel.Text = "Status: Diagnostic failed."
    } finally {
        $runButton.Enabled = $true
    }
})

$exportButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:LatestReport)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please run the diagnostic before exporting.",
            "No report available",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = "Export Network Diagnostic Report"
    $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.FileName = "Network_Report_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $saveDialog.OverwritePrompt = $true

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:LatestReport | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            $statusLabel.Text = "Status: Report exported to $($saveDialog.FileName)"
            [System.Windows.Forms.MessageBox]::Show(
                "Report exported successfully.",
                "Export complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not export report:`r`n$($_.Exception.Message)",
                "Export failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
    $saveDialog.Dispose()
})

$clearButton.Add_Click({
    $outputBox.Clear()
    $script:LatestReport = ""
    $exportButton.Enabled = $false
    $statusLabel.Text = "Status: Ready. Click Run Diagnostic to start."
})

[void]$form.ShowDialog()
$form.Dispose()
