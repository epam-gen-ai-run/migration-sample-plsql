#!/usr/bin/env pwsh

# Connect to PostgreSQL using Docker
# This script uses a PostgreSQL Docker container to run psql client

param(
    [string]$Query = "",
    [string]$ImportFile = "",
    [switch]$Interactive = $false
)

Write-Host "Connecting to PostgreSQL database..." -ForegroundColor Green

# Check if PostgreSQL container is running
$containerStatus = docker ps --filter "name=postgresql" --format "{{.Status}}" 2>$null
if (-not $containerStatus) {
    Write-Host "PostgreSQL container is not running!" -ForegroundColor Red
    Write-Host "Please start the container first using: .\start_postgresql.ps1" -ForegroundColor Yellow
    exit 1
}

Write-Host "PostgreSQL container status: $containerStatus" -ForegroundColor Cyan

# Connection details
$dbHost = "localhost"
$port = "5432"
$database = "employees"
$username = "demo"

if ($ImportFile) {
    # Import SQL file
    if (-not (Test-Path $ImportFile)) {
        Write-Host "Error: File '$ImportFile' not found!" -ForegroundColor Red
        exit 1
    }
    
    $absolutePath = Resolve-Path $ImportFile
    $fileName = Split-Path $ImportFile -Leaf
    
    Write-Host "Importing SQL file: $fileName" -ForegroundColor Yellow
    Write-Host "File path: $absolutePath" -ForegroundColor Gray
    
    # Copy file to container and execute
    Write-Host "Copying file to container..." -ForegroundColor Cyan
    docker cp "$absolutePath" postgresql:/tmp/$fileName
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Executing SQL file..." -ForegroundColor Cyan
        docker exec -it postgresql psql -h localhost -U $username -d $database -f "/tmp/$fileName"
        
        # Clean up temporary file in container
        docker exec postgresql rm "/tmp/$fileName"
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "SQL file imported successfully!" -ForegroundColor Green
        } else {
            Write-Host "Error occurred during SQL file execution" -ForegroundColor Red
        }
    } else {
        Write-Host "Error copying file to container" -ForegroundColor Red
        exit 1
    }
}
elseif ($Query) {
    # Execute a specific query
    Write-Host "Executing query: $Query" -ForegroundColor Yellow
    docker exec -it postgresql psql -h localhost -U $username -d $database -c "$Query"
}
elseif ($Interactive) {
    # Interactive mode
    Write-Host "Starting interactive psql session..." -ForegroundColor Cyan
    Write-Host "Connection details: $username@${dbHost}:$port/$database" -ForegroundColor Gray
    Write-Host "Type '\q' to quit" -ForegroundColor Gray
    docker exec -it postgresql psql -h localhost -U $username -d $database
}
else {
    # Default: show connection info and basic database info
    Write-Host "Getting database information..." -ForegroundColor Cyan
    Write-Host ""
    
    # Show connection info
    docker exec postgresql psql -h localhost -U $username -d $database -c "\conninfo"
    
    Write-Host ""
    Write-Host "Available tables:" -ForegroundColor Yellow
    docker exec postgresql psql -h localhost -U $username -d $database -c "\dt"
    
    Write-Host ""
    Write-Host "Database size:" -ForegroundColor Yellow
    docker exec postgresql psql -h localhost -U $username -d $database -c "SELECT pg_size_pretty(pg_database_size('$database')) as database_size;"
    
    Write-Host ""
    Write-Host "Usage examples:" -ForegroundColor Cyan
    Write-Host "  Interactive session:  .\connect_postgresql.ps1 -Interactive" -ForegroundColor White
    Write-Host "  Execute query:        .\connect_postgresql.ps1 -Query 'SELECT version();'" -ForegroundColor White
    Write-Host "  Import SQL file:      .\connect_postgresql.ps1 -ImportFile 'path\to\file.sql'" -ForegroundColor White
}