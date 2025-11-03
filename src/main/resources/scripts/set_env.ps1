# Script to set environment variables from a .env file
#
# USAGE:
# 1. Save this script as 'Set-EnvVariables.ps1'
# 2. Create a '.env' file in the same directory with your key-value pairs.
# 3. Run the script from your PowerShell terminal: .\Set-EnvVariables.ps1
#
# To make the variables persistent, uncomment the alternative method in the script.

# Get the directory of the currently running script
$scriptPath = $PSScriptRoot

# Define the path to the .env file
$envFilePath = Join-Path -Path $scriptPath -ChildPath "..\.env"

# Check if the .env file exists
if (-not (Test-Path $envFilePath)) {
    Write-Host "Error: .env file not found at $envFilePath"
    exit
}

# Read the content of the .env file
Get-Content $envFilePath | ForEach-Object {
    # Trim whitespace from the beginning and end of the line
    $line = $_.Trim()

    # Ignore empty lines and comments (lines starting with #)
    if ($line -ne "" -and -not $line.StartsWith("#")) {
        # Split the line into key and value at the first '='
        $parts = $line -split '=', 2

        if ($parts.Length -eq 2) {
            $key = $parts[0].Trim()
            $value = $parts[1].Trim()

            Set-Item -Path Env:$key -Value $value
            Write-Host "Set environment variable: $key"
            # This will make the variable available in future PowerShell sessions for the current user.
            # To set it for all users, change 'User' to 'Machine' (requires administrator privileges).
            # [System.Environment]::SetEnvironmentVariable($key, $value, 'User')
            # Write-Host "Set persistent environment variable for the current user: $key"
        }
    }
}

Write-Host "`nEnvironment variables from .env file have been set for the current session."```
