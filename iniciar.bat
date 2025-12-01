@echo off
cd /d "C:\Users\Mucio P\Documents\Projeto Python 3.0 - Mucio\Aplicativo\Dash Trigo"
echo Iniciando Dashboard de Sementes...
echo Acesse: http://localhost:5005
echo.
start cmd /k "python app.py"
timeout /t 2
echo Pronto! Verifique a janela do Flask que abriu.
pause