#!/usr/bin/env pwsh

# Copy Excel Report from Docker Container to Host
# This script generates the report in the container and copies it to the host

Write-Host "Generating and Retrieving Excel Report from Docker Container..." -ForegroundColor Green

try {
    Write-Host "Step 1: Generating Excel CSV content..." -ForegroundColor Yellow

    $sqlQuery = "SELECT generate_salary_report_to_csv();"
    & .\connect_postgresql.ps1 -Query $sqlQuery
    
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "`nDone!" -ForegroundColor Green