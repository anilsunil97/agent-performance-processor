# GitHub Authentication Setup Script
# Run this script to configure your GitHub credentials

Write-Host "=== GitHub Authentication Setup ===" -ForegroundColor Green
Write-Host ""

# Get user information
$username = Read-Host "Enter your GitHub username"
$email = Read-Host "Enter your GitHub email"

# Configure Git with user information
Write-Host "Configuring Git with your information..." -ForegroundColor Yellow
& "C:\Program Files\Git\bin\git.exe" config --global user.name $username
& "C:\Program Files\Git\bin\git.exe" config --global user.email $email

Write-Host "Git configuration updated!" -ForegroundColor Green
Write-Host ""

# Instructions for Personal Access Token
Write-Host "=== IMPORTANT: Personal Access Token Required ===" -ForegroundColor Red
Write-Host ""
Write-Host "GitHub no longer accepts passwords for Git operations." -ForegroundColor Yellow
Write-Host "You need to create a Personal Access Token (PAT):" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Go to: https://github.com/settings/tokens" -ForegroundColor Cyan
Write-Host "2. Click 'Generate new token' -> 'Generate new token (classic)'" -ForegroundColor Cyan
Write-Host "3. Give it a name: 'Kiro IDE Access'" -ForegroundColor Cyan
Write-Host "4. Select scopes: repo, workflow, write:packages" -ForegroundColor Cyan
Write-Host "5. Click 'Generate token'" -ForegroundColor Cyan
Write-Host "6. COPY THE TOKEN (you won't see it again!)" -ForegroundColor Cyan
Write-Host ""
Write-Host "When you push to GitHub, use:" -ForegroundColor Green
Write-Host "Username: $username" -ForegroundColor Green
Write-Host "Password: [Your Personal Access Token]" -ForegroundColor Green
Write-Host ""

# Test connection
Write-Host "Testing Git configuration..." -ForegroundColor Yellow
& "C:\Program Files\Git\bin\git.exe" config --list | Select-String "user"

Write-Host ""
Write-Host "Setup complete! Now you can use Git commands." -ForegroundColor Green
Write-Host "Next: Create your Personal Access Token and try pushing to GitHub." -ForegroundColor Yellow