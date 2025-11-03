#!/usr/bin/env pwsh

# Start PostgreSQL Docker container
# This script starts the PostgreSQL container defined in docker/postgresql.yaml

Write-Host "Starting PostgreSQL Docker container..." -ForegroundColor Green

# Check if Docker is running
Write-Host "Checking Docker status..." -ForegroundColor Yellow
try {
    $dockerVersion = docker version --format json 2>&1 
    if ($LASTEXITCODE -ne 0) {
        throw "Docker command failed"
    }
    Write-Host "Docker is running" -ForegroundColor Green
    
    # Additional check for Docker daemon connectivity
    Write-Host "Testing Docker daemon connectivity..." -ForegroundColor Yellow
    $dockerInfo = docker info 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Docker daemon connection issues detected" -ForegroundColor Red
        Write-Host "Docker may still be starting up. Please wait and try again." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "Docker daemon is accessible" -ForegroundColor Green
}
catch {
    Write-Host "Error: Docker Desktop is not running properly" -ForegroundColor Red
    Write-Host "Please ensure Docker Desktop is:" -ForegroundColor Yellow
    Write-Host "  1. Installed and running" -ForegroundColor White
    Write-Host "  2. Fully started (not just starting up)" -ForegroundColor White
    Write-Host "  3. Not stuck in a startup loop" -ForegroundColor White
    Write-Host ""
    Write-Host "Try these steps:" -ForegroundColor Cyan
    Write-Host "  1. Close Docker Desktop completely" -ForegroundColor White
    Write-Host "  2. Restart Docker Desktop as Administrator" -ForegroundColor White
    Write-Host "  3. Wait for it to fully start (green status)" -ForegroundColor White
    Write-Host "  4. Run this script again" -ForegroundColor White
    exit 1
}

# Ensure the pg_data directory exists
$pgDataDir = "C:\git-codemie\ai_demo_plsql2java\db\pg_data"
if (-not (Test-Path $pgDataDir)) {
    Write-Host "Creating pg_data directory: $pgDataDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $pgDataDir -Force | Out-Null
}

# Change to the docker directory
$dockerDir = Join-Path $PSScriptRoot "../docker"
Push-Location $dockerDir

try {
    # Start the PostgreSQL container
    Write-Host "Starting PostgreSQL container using docker-compose..." -ForegroundColor Cyan
    docker-compose -f postgresql.yaml up -d
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "PostgreSQL container started successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Connection details:" -ForegroundColor Cyan
        Write-Host "  Host: localhost" -ForegroundColor White
        Write-Host "  Port: 5432" -ForegroundColor White
        Write-Host "  Database: demodb" -ForegroundColor White
        Write-Host "  Username: demo" -ForegroundColor White
        Write-Host "  Password: demo" -ForegroundColor White
        Write-Host ""
        Write-Host "Data directory: $pgDataDir" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "To connect using psql:" -ForegroundColor Yellow
        Write-Host "  psql -h localhost -U demo -d demodb" -ForegroundColor White
        Write-Host ""
        Write-Host "To stop the container:" -ForegroundColor Yellow
        Write-Host "  docker-compose -f ../docker/postgresql.yaml down" -ForegroundColor White
    }
    else {
        Write-Host "Failed to start PostgreSQL container" -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host "Error starting PostgreSQL container: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Return to original directory
    Pop-Location
}