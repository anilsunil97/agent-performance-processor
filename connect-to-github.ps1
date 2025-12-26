# Connect to Existing GitHub Repository
# This script helps you connect your local project to your GitHub repository

Write-Host "=== Connect to GitHub Repository ===" -ForegroundColor Green
Write-Host ""

# Get repository information
$username = Read-Host "Enter your GitHub username"
$reponame = Read-Host "Enter your repository name (e.g., agent-performance-processor)"

# Construct repository URL
$repoUrl = "https://github.com/$username/$reponame.git"

Write-Host ""
Write-Host "Repository URL: $repoUrl" -ForegroundColor Cyan
Write-Host ""

# Remove existing remote if it exists
Write-Host "Removing any existing remote..." -ForegroundColor Yellow
& "C:\Program Files\Git\bin\git.exe" remote remove origin 2>$null

# Add the new remote
Write-Host "Adding GitHub repository as remote..." -ForegroundColor Yellow
& "C:\Program Files\Git\bin\git.exe" remote add origin $repoUrl

# Verify remote was added
Write-Host "Verifying remote repository..." -ForegroundColor Yellow
& "C:\Program Files\Git\bin\git.exe" remote -v

Write-Host ""
Write-Host "Repository connected successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Make sure you have your Personal Access Token ready" -ForegroundColor White
Write-Host "2. Try pushing your code:" -ForegroundColor White
Write-Host "   . .\git-commands.ps1" -ForegroundColor Cyan
Write-Host "   gp" -ForegroundColor Cyan
Write-Host ""
Write-Host "When prompted for credentials:" -ForegroundColor Yellow
Write-Host "Username: $username" -ForegroundColor Green
Write-Host "Password: [Your Personal Access Token]" -ForegroundColor Green