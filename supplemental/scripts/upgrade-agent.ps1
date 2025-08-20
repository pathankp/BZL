param (
    [switch]$Elevated
)

# Stop on first error
$ErrorActionPreference = "Stop"

#region Utility Functions

# Function to check if running as admin
function Test-Admin {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check if a command exists
function Test-CommandExists {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Command
    )
    return (Get-Command $Command -ErrorAction SilentlyContinue)
}

# Function to find serversentry-agent in common installation locations
function Find-ServerSentryAgent {
    # First check if it's in PATH
    $agentCmd = Get-Command "serversentry-agent" -ErrorAction SilentlyContinue
    if ($agentCmd) {
        return $agentCmd.Source
    }
    
    # Common installation paths to check
    $commonPaths = @(
        "$env:USERPROFILE\scoop\apps\serversentry-agent\current\serversentry-agent.exe",
        "$env:ProgramData\scoop\apps\serversentry-agent\current\serversentry-agent.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\nak-ventures.serversentry-agent*\serversentry-agent.exe",
        "$env:ProgramFiles\WinGet\Packages\nak-ventures.serversentry-agent*\serversentry-agent.exe",
        "${env:ProgramFiles(x86)}\WinGet\Packages\nak-ventures.serversentry-agent*\serversentry-agent.exe",
        "$env:ProgramFiles\serversentry-agent\serversentry-agent.exe",
        "$env:ProgramFiles(x86)\serversentry-agent\serversentry-agent.exe",
        "$env:SystemDrive\Users\*\scoop\apps\serversentry-agent\current\serversentry-agent.exe"
    )
    
    foreach ($path in $commonPaths) {
        # Handle wildcard paths
        if ($path.Contains("*")) {
            $foundPaths = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            if ($foundPaths) {
                return $foundPaths[0].FullName
            }
        } else {
            if (Test-Path $path) {
                return $path
            }
        }
    }
    
    return $null
}

# Function to find NSSM in common installation locations
function Find-NSSM {
    # First check if it's in PATH
    $nssmCmd = Get-Command "nssm" -ErrorAction SilentlyContinue
    if ($nssmCmd) {
        return $nssmCmd.Source
    }
    
    # Common installation paths to check
    $commonPaths = @(
        "$env:USERPROFILE\scoop\apps\nssm\current\nssm.exe",
        "$env:ProgramData\scoop\apps\nssm\current\nssm.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\NSSM.NSSM*\nssm.exe",
        "$env:ProgramFiles\WinGet\Packages\NSSM.NSSM*\nssm.exe",
        "${env:ProgramFiles(x86)}\WinGet\Packages\NSSM.NSSM*\nssm.exe",
        "$env:SystemDrive\Users\*\scoop\apps\nssm\current\nssm.exe"
    )
    
    foreach ($path in $commonPaths) {
        # Handle wildcard paths
        if ($path.Contains("*")) {
            $foundPaths = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            if ($foundPaths) {
                return $foundPaths[0].FullName
            }
        } else {
            if (Test-Path $path) {
                return $path
            }
        }
    }
    
    return $null
}

#endregion

#region Upgrade Functions

# Function to upgrade serversentry-agent with Scoop
function Upgrade-ServerSentryAgentWithScoop {
    Write-Host "Upgrading serversentry-agent with Scoop..."
    scoop update serversentry-agent
    
    if (-not (Test-CommandExists "serversentry-agent")) {
        throw "Failed to upgrade serversentry-agent with Scoop"
    }
    
    return $(Join-Path -Path $(scoop prefix serversentry-agent) -ChildPath "serversentry-agent.exe")
}

# Function to upgrade serversentry-agent with WinGet
function Upgrade-ServerSentryAgentWithWinGet {
    Write-Host "Upgrading serversentry-agent with WinGet..."
    
    # Temporarily change ErrorActionPreference to allow WinGet to complete and show output
    $originalErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    
    # Use call operator (&) and capture exit code properly
    & winget upgrade --exact --id nak-ventures.serversentry-agent --accept-source-agreements --accept-package-agreements | Out-Null
    $wingetExitCode = $LASTEXITCODE
    
    # Restore original ErrorActionPreference
    $ErrorActionPreference = $originalErrorActionPreference
    
    # WinGet exit codes:
    # 0 = Success
    # -1978335212 (0x8A150014) = No applicable upgrade found (package is up to date)
    # -1978335189 (0x8A15002B) = Another "no upgrade needed" variant
    # Other codes indicate actual errors
    if ($wingetExitCode -eq -1978335212 -or $wingetExitCode -eq -1978335189) {
        Write-Host "Package is already up to date." -ForegroundColor Green
    } elseif ($wingetExitCode -ne 0)  {
        Write-Host "WinGet exit code: $wingetExitCode" -ForegroundColor Yellow
    }
    
    # Refresh PATH environment variable to make serversentry-agent available in current session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    
    # Find the path to the serversentry-agent executable
    $agentPath = (Get-Command serversentry-agent -ErrorAction SilentlyContinue).Source
    
    if (-not $agentPath) {
        # Try to find it using our search function
        $agentPath = Find-ServerSentryAgent
        if (-not $agentPath) {
            throw "Could not find serversentry-agent executable path after upgrade"
        }
    }
    
    return $agentPath
}

# Function to get current service configuration
function Get-ServiceConfiguration {
    param (
        [string]$NSSMPath = ""
    )
    
    # Determine the NSSM executable to use
    $nssmCommand = "nssm"
    if ($NSSMPath -and (Test-Path $NSSMPath)) {
        $nssmCommand = $NSSMPath
    } elseif (-not (Test-CommandExists "nssm")) {
        throw "NSSM is not available in PATH and no valid NSSMPath was provided"
    }
    
    # Check if service exists
    $existingService = Get-Service -Name "serversentry-agent" -ErrorAction SilentlyContinue
    if (-not $existingService) {
        throw "serversentry-agent service does not exist. Please run the installation script first."
    }
    
    # Get current service configuration
    $config = @{}
    
    try {
        # Get current application path
        $currentPath = & $nssmCommand get serversentry-agent Application
        if ($LASTEXITCODE -eq 0) {
            $config.CurrentPath = $currentPath.Trim()
        }
        
        # Get environment variables
        $envVars = & $nssmCommand get serversentry-agent AppEnvironmentExtra
        if ($LASTEXITCODE -eq 0 -and $envVars) {
            $config.EnvironmentVars = $envVars
        }
        
        Write-Host "Current service configuration retrieved successfully."
        Write-Host "Current agent path: $($config.CurrentPath)"
        
        return $config
    }
    catch {
        throw "Failed to retrieve current service configuration: $($_.Exception.Message)"
    }
}

# Function to update service path
function Update-ServicePath {
    param (
        [Parameter(Mandatory=$true)]
        [string]$NewAgentPath,
        [string]$NSSMPath = ""
    )
    
    Write-Host "Updating serversentry-agent service path..."
    
    # Determine the NSSM executable to use
    $nssmCommand = "nssm"
    if ($NSSMPath -and (Test-Path $NSSMPath)) {
        $nssmCommand = $NSSMPath
        Write-Host "Using NSSM from: $NSSMPath"
    } elseif (-not (Test-CommandExists "nssm")) {
        throw "NSSM is not available in PATH and no valid NSSMPath was provided"
    }
    
    # Update the application path
    & $nssmCommand set serversentry-agent Application $NewAgentPath
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to update serversentry-agent service path"
    }
    
    Write-Host "Service path updated to: $NewAgentPath"
    
    # Start the service
    Start-ServerSentryAgentService -NSSMPath $nssmCommand
}

# Function to start and monitor the service
function Start-ServerSentryAgentService {
    param (
        [string]$NSSMPath = ""
    )
    
    Write-Host "Starting serversentry-agent service..."
    
    # Determine the NSSM executable to use
    $nssmCommand = "nssm"
    if ($NSSMPath -and (Test-Path $NSSMPath)) {
        $nssmCommand = $NSSMPath
    } elseif (-not (Test-CommandExists "nssm")) {
        throw "NSSM is not available in PATH and no valid NSSMPath was provided"
    }
    
    & $nssmCommand start serversentry-agent
    $startResult = $LASTEXITCODE
    
    # Only enter the status check loop if the NSSM start command failed
    if ($startResult -ne 0) {
        Write-Host "NSSM start command returned error code: $startResult" -ForegroundColor Yellow
        Write-Host "This could be due to 'SERVICE_START_PENDING' state. Checking service status..."
        
        # Allow up to 10 seconds for the service to start, checking every second
        $maxWaitTime = 10 # seconds
        $elapsedTime = 0
        $serviceStarted = $false
        
        while (-not $serviceStarted -and $elapsedTime -lt $maxWaitTime) {
            Start-Sleep -Seconds 1
            $elapsedTime += 1

            $serviceStatus = & $nssmCommand status serversentry-agent
            
            if ($serviceStatus -eq "SERVICE_RUNNING") {
                $serviceStarted = $true
                Write-Host "Success! The serversentry-agent service is now running." -ForegroundColor Green
            }
            elseif ($serviceStatus -like "*PENDING*") {
                Write-Host "Service is still starting (status: $serviceStatus)... waiting" -ForegroundColor Yellow
            }
            else {
                Write-Host "Warning: The service status is '$serviceStatus' instead of 'SERVICE_RUNNING'." -ForegroundColor Yellow
                Write-Host "You may need to troubleshoot the service installation." -ForegroundColor Yellow
                break
            }
        }
        
        if (-not $serviceStarted) {
            Write-Host "Service did not reach running state." -ForegroundColor Yellow
            Write-Host "You can check status manually with 'nssm status serversentry-agent'" -ForegroundColor Yellow
        }
    } else {
        # NSSM start command was successful
        Write-Host "Success! The serversentry-agent service is running properly." -ForegroundColor Green
    }
}

#endregion

#region Main Script Execution

# Check if we're running as admin
$isAdmin = Test-Admin

try {
    Write-Host "ServerSentry Agent Upgrade Script" -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    
    # First: Check if service exists (doesn't require admin)
    $existingService = Get-Service -Name "serversentry-agent" -ErrorAction SilentlyContinue
    if (-not $existingService) {
        Write-Host "ERROR: serversentry-agent service does not exist." -ForegroundColor Red
        Write-Host "Please run the installation script first before attempting to upgrade." -ForegroundColor Red
        exit 1
    }
    
    # Find current NSSM and agent paths
    $nssmPath = Find-NSSM
    if (-not $nssmPath -and (Test-CommandExists "nssm")) {
        $nssmPath = (Get-Command "nssm" -ErrorAction SilentlyContinue).Source
    }
    
    if (-not $nssmPath) {
        Write-Host "ERROR: NSSM not found. Cannot manage the service without NSSM." -ForegroundColor Red
        exit 1
    }
    
    # Get current service configuration (doesn't require admin)
    Write-Host "Retrieving current service configuration..."
    $currentConfig = Get-ServiceConfiguration -NSSMPath $nssmPath
    
    # Stop the service before upgrade
    Write-Host "Stopping serversentry-agent service..."
    & $nssmPath stop serversentry-agent
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Failed to stop service, continuing anyway..." -ForegroundColor Yellow
    }
    
    # Upgrade the agent (doesn't require admin)
    Write-Host "Upgrading serversentry-agent..."
    $newAgentPath = $null
    
    if (Test-CommandExists "scoop") {
        Write-Host "Using Scoop for upgrade..."
        $newAgentPath = Upgrade-ServerSentryAgentWithScoop
    }
    elseif (Test-CommandExists "winget") {
        Write-Host "Using WinGet for upgrade..."
        $newAgentPath = Upgrade-ServerSentryAgentWithWinGet
    }
    else {
        Write-Host "ERROR: Neither Scoop nor WinGet is available for upgrading." -ForegroundColor Red
        exit 1
    }
    
    if (-not $newAgentPath) {
        $newAgentPath = Find-ServerSentryAgent
        if (-not $newAgentPath) {
            throw "Could not find serversentry-agent executable after upgrade."
        }
    }
    
    Write-Host "New agent path: $newAgentPath"
    
    # Check if the path has changed
    if ($currentConfig.CurrentPath -eq $newAgentPath) {
        Write-Host "Agent path has not changed. Restarting service..." -ForegroundColor Green
        Start-ServerSentryAgentService -NSSMPath $nssmPath
        Write-Host "Upgrade completed successfully!" -ForegroundColor Green
        exit 0
    }
    
    Write-Host "Agent path has changed from:" -ForegroundColor Yellow
    Write-Host "  Old: $($currentConfig.CurrentPath)" -ForegroundColor Yellow
    Write-Host "  New: $newAgentPath" -ForegroundColor Yellow
    Write-Host ""
    
    # If we need admin rights for service update and we don't have them, relaunch
    if (-not $isAdmin -and -not $Elevated) {
        Write-Host "Admin privileges required for service path update. Relaunching as admin..." -ForegroundColor Yellow
        
        # Prepare arguments for the elevated script
        $argumentList = @(
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$PSCommandPath`"",
            "-Elevated"
        )
        
        # Relaunch the script with the -Elevated switch
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argumentList
        exit
    }
    
    # Update service path (requires admin)
    if ($isAdmin -or $Elevated) {
        Update-ServicePath -NewAgentPath $newAgentPath -NSSMPath $nssmPath
        
        Write-Host ""
        Write-Host "Upgrade completed successfully!" -ForegroundColor Green
        Write-Host "The serversentry-agent service has been updated to use the new executable path." -ForegroundColor Green
        
        # Pause to see results if this is an elevated window
        if ($Elevated) {
            Write-Host ""
            Write-Host "Press any key to exit..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Upgrade failed. Please check the error message above." -ForegroundColor Red
    
    # Pause if this is likely a new window
    if ($Elevated -or (-not $isAdmin)) {
        Write-Host "Press any key to exit..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    exit 1
}

#endregion 
