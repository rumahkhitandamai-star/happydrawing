@echo off
title Happy Drawing - CLASP Auto Push & Deploy Watcher
echo ==============================================================
echo   Happy Drawing (SIMS) - CLASP Auto Push & Deploy Watcher
echo   Setiap perubahan pada Code.gs / index.html akan otomatis
echo   di-push dan di-deploy ke URL Production Web App:
echo   https://script.google.com/macros/s/AKfycbyEsKdyMBJzJl2E15bOnkm5A7MYLcBB4xJePka1mz9icDHspmxw6gWu2e8jesl_8QLa/exec
echo ==============================================================
echo.
node watch-deploy.js
pause
