@echo off
setlocal

for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8765"') do taskkill /PID %%P /F >nul 2>nul
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8766"') do taskkill /PID %%P /F >nul 2>nul
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8767"') do taskkill /PID %%P /F >nul 2>nul
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8768"') do taskkill /PID %%P /F >nul 2>nul
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8769"') do taskkill /PID %%P /F >nul 2>nul

echo MICE Travel Bot detenido.
pause
