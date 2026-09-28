@echo off
chcp 65001 >nul
title Painel SPED - Compras e Vendas
REM Arraste para este arquivo a PASTA do mes ou os arquivos SPED (.txt).
REM Se abrir com duplo clique, aparece uma janela para escolher os arquivos.
python -c "import openpyxl" 2>nul || python -m pip install --user openpyxl
python "%~dp0sped_painel.py" %*
echo.
pause