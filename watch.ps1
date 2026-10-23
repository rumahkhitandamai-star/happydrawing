# Happy Drawing (SIMS) - CLASP Auto Push & Deploy Watcher
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "  Happy Drawing (SIMS) - CLASP Auto Push & Deploy Watcher" -ForegroundColor Yellow
Write-Host "  Setiap perubahan pada Code.gs / index.html akan otomatis" -ForegroundColor White
Write-Host "  di-push dan di-deploy ke URL Production Web App:" -ForegroundColor White
Write-Host "  https://script.google.com/macros/s/AKfycbyEsKdyMBJzJl2E15bOnkm5A7MYLcBB4xJePka1mz9icDHspmxw6gWu2e8jesl_8QLa/exec" -ForegroundColor Green
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "Tekan Ctrl+C untuk menghentikan watcher.`n" -ForegroundColor Gray

node watch-deploy.js
