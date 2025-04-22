@echo off
cd /d %~dp0
echo Executando extracao de horas por aplicacao...
python extrator_local.py
pause
