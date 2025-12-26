# Git Commands Helper for Kiro IDE
# Usage: . .\git-commands.ps1 (to load functions)

$GitPath = "C:\Program Files\Git\bin\git.exe"

function Git-Status {
    & $GitPath status
}

function Git-Add($files = ".") {
    & $GitPath add $files
}

function Git-Commit($message) {
    & $GitPath commit -m $message
}

function Git-Push($branch = "main") {
    & $GitPath push origin $branch
}

function Git-Pull($branch = "main") {
    & $GitPath pull origin $branch
}

function Git-Branch {
    & $GitPath branch -a
}

function Git-Checkout($branch) {
    & $GitPath checkout $branch
}

function Git-NewBranch($branchName) {
    & $GitPath checkout -b $branchName
}

function Git-Log {
    & $GitPath log --oneline -10
}

function Git-Diff {
    & $GitPath diff
}

# Aliases
Set-Alias gs Git-Status
Set-Alias ga Git-Add
Set-Alias gc Git-Commit
Set-Alias gp Git-Push
Set-Alias gl Git-Pull
Set-Alias gb Git-Branch
Set-Alias gco Git-Checkout
Set-Alias gnb Git-NewBranch
Set-Alias glog Git-Log
Set-Alias gd Git-Diff

Write-Host "Git commands loaded! Available commands:" -ForegroundColor Green
Write-Host "gs  - Git Status" -ForegroundColor Yellow
Write-Host "ga  - Git Add" -ForegroundColor Yellow
Write-Host "gc  - Git Commit" -ForegroundColor Yellow
Write-Host "gp  - Git Push" -ForegroundColor Yellow
Write-Host "gl  - Git Pull" -ForegroundColor Yellow
Write-Host "gb  - Git Branch" -ForegroundColor Yellow
Write-Host "gco - Git Checkout" -ForegroundColor Yellow
Write-Host "gnb - Git New Branch" -ForegroundColor Yellow
Write-Host "glog- Git Log" -ForegroundColor Yellow
Write-Host "gd  - Git Diff" -ForegroundColor Yellow