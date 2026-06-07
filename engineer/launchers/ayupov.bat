@echo off
setlocal
cd /d "%~dp0..\.."
set SK_ENGINEER_PROFILE=ayupov
echo SK-Reporter engineer profile: %SK_ENGINEER_PROFILE%
echo Open http://127.0.0.1:8010/engineer after starting the server.
start http://127.0.0.1:8010/engineer
call scripts\run-server.ps1
