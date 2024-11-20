# Requires running PowerShell as Administrator for Event Log creation

# Add sleep prevention
Add-Type @'
using System;
using System.Runtime.InteropServices;

public class Prevention {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);

    public enum ExecutionState : uint {
        EsAwaymodeRequired = 0x00000040,
        EsContinuous = 0x80000000,
        EsDisplayRequired = 0x00000002,
        EsSystemRequired = 0x00000001
    }
}
'@

# Function to setup Event Log source
function Initialize-EventLogging {
    $logName = "Application"
    $sourceName = "NetworkMonitor"
    
    if (-not [System.Diagnostics.EventLog]::SourceExists($sourceName)) {
        try {
            New-EventLog -LogName $logName -Source $sourceName
            Write-Host "Created Event Log source: $sourceName"
        }
        catch {
            Write-Host "Failed to create Event Log source. Are you running as Administrator?" -ForegroundColor Red
        }
    }
    return $sourceName
}

# Function to update heartbeat file
function Update-Heartbeat {
    param (
        [string]$heartbeatPath
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $status = @{
        LastUpdate = $timestamp
        LastConnectivityTest = $script:lastConnectivityTime
        LastSpeedTest = $script:lastSpeedTestTime
        ScriptRunning = $true
    }
    $status | ConvertTo-Json | Set-Content -Path $heartbeatPath
}

# Function to get log filename
function Get-LogFilename {
    param (
        [string]$type
    )
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $logsPath = Join-Path -Path $scriptPath -ChildPath "NetworkLogs"
    
    if (-not (Test-Path $logsPath)) {
        New-Item -ItemType Directory -Path $logsPath -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyy-MMM-dd-HH"
    $filename = "$timestamp-$type.csv"
    return Join-Path $logsPath $filename
}

# Function to test network connectivity
function Test-NetworkConnectivity {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    $targets = @(
        "8.8.8.8",         # Google DNS
        "1.1.1.1",         # Cloudflare DNS
        "208.67.222.222"   # OpenDNS
    )
    
    $results = @()
    foreach ($target in $targets) {
        $pingResult = Test-Connection -ComputerName $target -Count 1 -ErrorAction SilentlyContinue
        if ($pingResult) {
            $results += [PSCustomObject]@{
                Target = $target
                Success = $true
                ResponseTime = $pingResult.ResponseTime
            }
        } else {
            $results += [PSCustomObject]@{
                Target = $target
                Success = $false
                ResponseTime = 0
            }
        }
    }
    
    $successfulPings = ($results | Where-Object Success -eq $true).Count
    $averageResponseTime = ($results | Where-Object Success -eq $true | Measure-Object -Property ResponseTime -Average).Average
    
    # Create object for CSV export
    $connectivityData = [PSCustomObject]@{
        Timestamp = $timestamp
        SuccessfulPings = $successfulPings
        TotalTargets = $targets.Count
        AverageResponseTime = [math]::Round($averageResponseTime, 2)
        Status = if ($successfulPings -gt 0) { "Connected" } else { "Disconnected" }
    }
    
    # Export immediately
    $connectionFile = Get-LogFilename -type "connectivity"
    $connectivityData | Export-Csv -Path $connectionFile -NoTypeInformation -Append
    
    # Display results
    Write-Host "[$timestamp] Network Test Results:"
    Write-Host "Status: $($connectivityData.Status)"
    Write-Host "Successful Connections: $successfulPings/$($targets.Count)"
    if ($averageResponseTime) {
        Write-Host "Average Response Time: $([math]::Round($averageResponseTime, 2)) ms"
    }
    Write-Host "------------------------------------------"
    
    return $successfulPings -gt 0
}

# Function to test network speed
function Test-NetworkSpeed {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] Starting Speed Test..."
    
    try {
        # Initialize variables for speed test
        $downloadSpeeds = @()
        $testUrls = @(
            "http://proof.ovh.net/files/1Mb.dat"
            "http://speedtest.tele2.net/1MB.zip"
        )
        
        foreach ($url in $testUrls) {
            try {
                $start = Get-Date
                $wc = New-Object System.Net.WebClient
                $data = $wc.DownloadData($url)
                $end = Get-Date
                $duration = ($end - $start).TotalSeconds
                $fileSize = $data.Length
                $speed = [math]::Round(($fileSize * 8) / ($duration * 1000000), 2) # Convert to Mbps
                $downloadSpeeds += $speed
            }
            catch {
                Write-Host "Failed to test download speed with $url" -ForegroundColor Yellow
            }
        }
        
        $averageSpeed = if ($downloadSpeeds.Count -gt 0) {
            ($downloadSpeeds | Measure-Object -Average).Average
        } else {
            0
        }
        
        # Create object for CSV export
        $speedData = [PSCustomObject]@{
            Timestamp = $timestamp
            DownloadSpeed = $averageSpeed
            TestsCompleted = $downloadSpeeds.Count
            TotalTests = $testUrls.Count
        }
        
        # Export immediately
        $speedFile = Get-LogFilename -type "speed"
        $speedData | Export-Csv -Path $speedFile -NoTypeInformation -Append
        
        Write-Host "[$timestamp] Speed Test Results:"
        Write-Host "Average Download Speed: $averageSpeed Mbps"
        Write-Host "Completed Tests: $($downloadSpeeds.Count)/$($testUrls.Count)"
        Write-Host "------------------------------------------"
        
        return $averageSpeed -gt 0
    }
    catch {
        Write-Host "[$timestamp] Speed Test failed: $_" -ForegroundColor Red
        return $false
    }
}

# Main script
Clear-Host
Write-Host "Network Monitoring Script Started"
Write-Host "------------------------------------------"

try {
    # Prevent system sleep
    [Prevention]::SetThreadExecutionState([Prevention+ExecutionState]::EsContinuous -bor 
        [Prevention+ExecutionState]::EsSystemRequired -bor 
        [Prevention+ExecutionState]::EsDisplayRequired)
    Write-Host "Sleep prevention enabled"

    # Setup paths
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $logsPath = Join-Path -Path $scriptPath -ChildPath "NetworkLogs"
    $heartbeatFile = Join-Path -Path $logsPath -ChildPath "heartbeat.json"

    # Create logs directory if it doesn't exist
    if (-not (Test-Path -Path $logsPath)) {
        New-Item -ItemType Directory -Path $logsPath -Force | Out-Null
    }

    # Create empty heartbeat file if it doesn't exist
    if (-not (Test-Path -Path $heartbeatFile)) {
        $null = New-Item -ItemType File -Path $heartbeatFile -Force
    }

    # Initialize Event Log
    $eventLogSource = Initialize-EventLogging

    # Initialize variables
    $connectivityInterval = 30    # Test connectivity every 30 seconds
    $speedTestInterval = 300      # Test speed every 5 minutes
    $heartbeatInterval = 30       # Update heartbeat every 30 seconds
    $lastConnectivityTest = 0
    $lastSpeedTest = 0
    $lastHeartbeat = 0

    # Initialize script variables
    $script:lastConnectivityTime = "Never"
    $script:lastSpeedTestTime = "Never"

    # Log script start
    Write-EventLog -LogName Application -Source $eventLogSource -EventId 1000 -EntryType Information -Message "Network monitoring script started"
    Write-Host "Log files will be stored in: $logsPath"
    Write-Host "Heartbeat file: $heartbeatFile"
    Write-Host "------------------------------------------"

    # Main loop
    while ($true) {
        $currentTime = [int](Get-Date -UFormat %s)
        
        try {
            # Check if it's time for a connectivity test
            if ($currentTime - $lastConnectivityTest -ge $connectivityInterval) {
                Test-NetworkConnectivity
                $lastConnectivityTest = $currentTime
                $script:lastConnectivityTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            # Check if it's time for a speed test
            if ($currentTime - $lastSpeedTest -ge $speedTestInterval) {
                Test-NetworkSpeed
                $lastSpeedTest = $currentTime
                $script:lastSpeedTestTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            # Update heartbeat file
            if ($currentTime - $lastHeartbeat -ge $heartbeatInterval) {
                Update-Heartbeat -heartbeatPath $heartbeatFile
                $lastHeartbeat = $currentTime
            }
        }
        catch {
            $errorMessage = "Error in main loop: $_"
            Write-Host $errorMessage -ForegroundColor Red
            Write-EventLog -LogName Application -Source $eventLogSource -EventId 1001 -EntryType Error -Message $errorMessage
        }
        
        Start-Sleep -Seconds 1
    }
}
finally {
    # Restore default power settings
    [Prevention]::SetThreadExecutionState([Prevention+ExecutionState]::EsContinuous)
    Write-Host "Restored default power settings"
    
    # Log script end
    if ($eventLogSource) {
        Write-EventLog -LogName Application -Source $eventLogSource -EventId 1000 -EntryType Information -Message "Network monitoring script ended"
    }
}