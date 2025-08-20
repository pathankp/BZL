param (
    [switch]$Elevated,
    [Parameter(Mandatory=$true)]
    [string]$Key,
    [string]$Token = "",
    [string]$Url = "",
    [int]$Port = 45876,
    [string]$AgentPath = "",
    [string]$NSSMPath = ""
)

# Check if required parameters are provided
if ([string]::IsNullOrWhiteSpace($Key)) {
    Write-Host "ERROR: SSH Key is required." -ForegroundColor Red
    Write-Host "Usage: .\install-agent.ps1 -Key 'your-ssh-key-here' [-Token 'your-token-here'] [-Url 'your-hub-url-here'] [-Port port-number]" -ForegroundColor Yellow
    Write-Host "Note: Token and Url are optional for backwards compatibility with older hub versions." -ForegroundColor Yellow
    exit 1
}

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

#region Installation Methods

# Function to install Scoop
function Install-Scoop {
    Write-Host "Installing Scoop..."
    
    # Check if running as admin - Scoop should not be installed as admin
    if (Test-Admin) {
        throw "Scoop cannot be installed with administrator privileges. Please run this script as a regular user first to install Scoop and serversentry-agent, then run as admin to configure the service."
    }
    
    try {
        Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression
        
        if (-not (Test-CommandExists "scoop")) {
            throw "Failed to install Scoop - command not available after installation"
        }
        Write-Host "Scoop installed successfully."
    }
    catch {
        throw "Failed to install Scoop: $($_.Exception.Message)"
    }
}

# Function to install Git via Scoop
function Install-Git {
    if (Test-CommandExists "git") {
        Write-Host "Git is already installed."
        return
    }
    
    Write-Host "Installing Git..."
    scoop install git
    
    if (-not (Test-CommandExists "git")) {
        throw "Failed to install Git"
    }
}

# Function to install NSSM
function Install-NSSM {
    param (
        [string]$Method = "Scoop" # Default to Scoop method
    )
    
    if (Test-CommandExists "nssm") {
        Write-Host "NSSM is already installed."
        return
    }
    
    Write-Host "Installing NSSM..."
    if ($Method -eq "Scoop") {
        scoop install nssm
    }
    elseif ($Method -eq "WinGet") {
        winget install -e --id NSSM.NSSM --accept-source-agreements --accept-package-agreements
        
        # Refresh PATH environment variable to make NSSM available in current session
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    }
    else {
        throw "Unsupported installation method: $Method"
    }
    
    if (-not (Test-CommandExists "nssm")) {
        throw "Failed to install NSSM"
    }
}

# Function to install serversentry-agent with Scoop
function Install-ServerSentryAgentWithScoop {
    Write-Host "Adding serversentry bucket..."
    scoop bucket add serversentry https://github.com/nak-ventures/serversentry-scoops | Out-Null
    
    Write-Host "Installing / updating serversentry-agent..."
    scoop install serversentry-agent | Out-Null
    
    if (-not (Test-CommandExists "serversentry-agent")) {
        throw "Failed to install serversentry-agent"
    }
    
    return $(Join-Path -Path $(scoop prefix serversentry-agent) -ChildPath "serversentry-agent.exe")
}

# Function to install serversentry-agent with WinGet
function Install-ServerSentryAgentWithWinGet {
    Write-Host "Installing / updating serversentry-agent..."
    
    # Temporarily change ErrorActionPreference to allow WinGet to complete and show output
    $originalErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    
    # Use call operator (&) and capture exit code properly
    & winget install --exact --id nak-ventures.serversentry-agent --accept-source-agreements --accept-package-agreements | Out-Null
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
        throw "Could not find serversentry-agent executable path after installation"
    }
    
    return $agentPath
}

# Function to install using Scoop
function Install-WithScoop {
    param (
        [string]$Key,
        [int]$Port
    )
    
    try {
        # Ensure Scoop is installed
        if (-not (Test-CommandExists "scoop")) {
            Install-Scoop | Out-Null
        }
        else {
            Write-Host "Scoop is already installed."
        }
        
        # Install Git (required for Scoop buckets)
        Install-Git | Out-Null
        
        # Install NSSM
        Install-NSSM -Method "Scoop" | Out-Null
        
        # Install serversentry-agent
        $agentPath = Install-ServerSentryAgentWithScoop
        
        return $agentPath
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Installation failed. Please check the error message above." -ForegroundColor Red
        Write-Host "Press any key to exit..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

# Function to install using WinGet
function Install-WithWinGet {
    param (
        [string]$Key,
        [int]$Port
    )
    
    try {
        # Install NSSM
        Install-NSSM -Method "WinGet" | Out-Null
        
        # Install serversentry-agent
        $agentPath = Install-ServerSentryAgentWithWinGet
        
        return $agentPath
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Installation failed. Please check the error message above." -ForegroundColor Red
        Write-Host "Press any key to exit..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

#endregion

#region Service Configuration

# Function to install and configure the NSSM service
function Install-NSSMService {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AgentPath,
        [Parameter(Mandatory=$true)]
        [string]$Key,
        [string]$Token = "",
        [string]$HubUrl = "",
        [Parameter(Mandatory=$true)]
        [int]$Port,
        [string]$NSSMPath = ""
    )
    
    Write-Host "Installing serversentry-agent service..."
    
    # Determine the NSSM executable to use
    $nssmCommand = "nssm"
    if ($NSSMPath -and (Test-Path $NSSMPath)) {
        $nssmCommand = $NSSMPath
        Write-Host "Using NSSM from: $NSSMPath"
    } elseif (-not (Test-CommandExists "nssm")) {
        throw "NSSM is not available in PATH and no valid NSSMPath was provided"
    }
    
    # Check if service already exists
    $existingService = Get-Service -Name "serversentry-agent" -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-Host "Service already exists. Checking if path update is needed..."
        
        # Get current service path 
        try {
            $currentPath = & $nssmCommand get serversentry-agent Application
            if ($LASTEXITCODE -eq 0 -and $currentPath.Trim() -eq $AgentPath) {
                Write-Host "Service already configured with correct path. Skipping service recreation." -ForegroundColor Green
                return
            }
            
            Write-Host "Service path needs updating. Stopping and removing existing service..."
            Write-Host "  Current path: $($currentPath.Trim())"
            Write-Host "  New path: $AgentPath"
        } catch {
            Write-Host "Could not retrieve current service path, will recreate service: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Service path needs updating. Stopping and removing existing service..."
        }
        
        try {
            & $nssmCommand stop serversentry-agent
            & $nssmCommand remove serversentry-agent confirm
        } catch {
            Write-Host "Warning: Failed to remove existing service: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    & $nssmCommand install serversentry-agent $AgentPath
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install serversentry-agent service"
    }
    
    Write-Host "Configuring service environment variables..."
    & $nssmCommand set serversentry-agent AppEnvironmentExtra "+KEY=$Key"
    & $nssmCommand set serversentry-agent AppEnvironmentExtra "+TOKEN=$Token"
    & $nssmCommand set serversentry-agent AppEnvironmentExtra "+HUB_URL=$HubUrl"
    & $nssmCommand set serversentry-agent AppEnvironmentExtra "+PORT=$Port"
    
    # Configure log files
    $logDir = "$env:ProgramData\serversentry-agent\logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $logFile = "$logDir\serversentry-agent.log"
    & $nssmCommand set serversentry-agent AppStdout $logFile
    & $nssmCommand set serversentry-agent AppStderr $logFile
}

# Function to configure firewall rules
function Configure-Firewall {
    param (
        [Parameter(Mandatory=$true)]
        [int]$Port
    )
    
    # Create a firewall rule if it doesn't exist
    $ruleName = "Allow serversentry-agent"
    $existingRule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
    
    # Remove existing rule if found
    if ($existingRule) {
        Write-Host "Removing existing firewall rule..."
        try {
            Remove-NetFirewallRule -DisplayName $ruleName
            Write-Host "Existing firewall rule removed successfully."
        } catch {
            Write-Host "Warning: Failed to remove existing firewall rule: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Create new rule with current settings
    Write-Host "Creating firewall rule for serversentry-agent on port $Port..."
    try {
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port
        Write-Host "Firewall rule created successfully."
    } catch {
        Write-Host "Warning: Failed to create firewall rule: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "You may need to manually create a firewall rule for port $Port." -ForegroundColor Yellow
    }
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
    # First: Install the agent (doesn't require admin)
    if (-not $AgentPath) {
        # Check for problematic case: running as admin and need Scoop
        if ($isAdmin -and -not (Test-CommandExists "scoop") -and -not (Test-CommandExists "winget")) {
            Write-Host "ERROR: You're running as administrator but neither Scoop nor WinGet is available." -ForegroundColor Red
            Write-Host "Scoop should be installed without admin privileges." -ForegroundColor Red
            Write-Host "" 
            Write-Host "Please either:" -ForegroundColor Yellow
            Write-Host "1. Run this script again without administrator privileges" -ForegroundColor Yellow
            Write-Host "2. Install WinGet and run this script again" -ForegroundColor Yellow
            exit 1
        }
        
        if (Test-CommandExists "scoop") {
            Write-Host "Using Scoop for installation..."
            $AgentPath = Install-WithScoop -Key $Key -Port $Port
        }
        elseif (Test-CommandExists "winget") {
            Write-Host "Using WinGet for installation..."
            $AgentPath = Install-WithWinGet -Key $Key -Port $Port
        }
        else {
            Write-Host "Neither Scoop nor WinGet is installed. Installing Scoop..."
            $AgentPath = Install-WithScoop -Key $Key -Port $Port
        }
    }

    if (-not $AgentPath) {
        throw "Could not find serversentry-agent executable. Make sure it was properly installed."
    }
    
    # Find NSSM path if not already provided
    if (-not $NSSMPath) {
        $NSSMPath = Find-NSSM
        
        if (-not $NSSMPath -and (Test-CommandExists "nssm")) {
            $NSSMPath = (Get-Command "nssm" -ErrorAction SilentlyContinue).Source
        }
        
        # If we still don't have NSSM, try to install it if we have package managers
        if (-not $NSSMPath) {
            if (Test-CommandExists "winget") {
                Write-Host "NSSM not found. Attempting to install via WinGet..."
                try {
                    Install-NSSM -Method "WinGet"
                    $NSSMPath = Find-NSSM
                    if (-not $NSSMPath -and (Test-CommandExists "nssm")) {
                        $NSSMPath = (Get-Command "nssm" -ErrorAction SilentlyContinue).Source
                    }
                } catch {
                    Write-Host "Failed to install NSSM via WinGet: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            } elseif (Test-CommandExists "scoop") {
                Write-Host "NSSM not found. Attempting to install via Scoop..."
                try {
                    Install-NSSM -Method "Scoop"
                    $NSSMPath = Find-NSSM
                    if (-not $NSSMPath -and (Test-CommandExists "nssm")) {
                        $NSSMPath = (Get-Command "nssm" -ErrorAction SilentlyContinue).Source
                    }
                } catch {
                    Write-Host "Failed to install NSSM via Scoop: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # Final check - if we still don't have NSSM and we're admin, we have a problem
            if (-not $NSSMPath -and ($isAdmin -or $Elevated)) {
                throw "NSSM is required for service installation but was not found and could not be installed. Please install NSSM manually or run as a regular user to install it."
            }
        }
    }
    
    # Second: If we need admin rights for service installation and we don't have them, relaunch
    if (-not $isAdmin -and -not $Elevated) {
        Write-Host "Admin privileges required for service installation. Relaunching as admin..." -ForegroundColor Yellow
        Write-Host "Check service status with 'nssm status serversentry-agent'"
        Write-Host "Edit service configuration with 'nssm edit serversentry-agent'"
        
        # Prepare arguments for the elevated script
        $argumentList = @(
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$PSCommandPath`"",
            "-Elevated",
            "-Key", "`"$Key`"",
            "-Token", "`"$Token`"",
            "-Url", "`"$Url`"",
            "-Port", $Port,
            "-AgentPath", "`"$AgentPath`""
        )
        
        # Add NSSMPath if we found it
        if ($NSSMPath) {
            $argumentList += "-NSSMPath"
            $argumentList += "`"$NSSMPath`""
        }
        
        # Relaunch the script with the -Elevated switch and pass parameters
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argumentList
        exit
    }
    
    # Third: If we have admin rights, install service and configure firewall
    if ($isAdmin -or $Elevated) {
        # Install the service
        Install-NSSMService -AgentPath $AgentPath -Key $Key -Token $Token -HubUrl $Url -Port $Port -NSSMPath $NSSMPath
        
        # Configure firewall
        Configure-Firewall -Port $Port
        
        # Start the service
        Start-ServerSentryAgentService -NSSMPath $NSSMPath
        
        # Pause to see results if this is an elevated window
        if ($Elevated) {
            Write-Host "Press any key to exit..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Installation failed. Please check the error message above." -ForegroundColor Red
    
    # Pause if this is likely a new window
    if ($Elevated -or (-not $isAdmin)) {
        Write-Host "Press any key to exit..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    exit 1
}

#endregion
