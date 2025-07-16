@echo off
TITLE Compilador do Projeto VerificadorApp

echo.
echo --- Limpando builds anteriores... ---
rmdir /s /q build
rmdir /s /q dist
echo Limpeza concluida.

echo.
echo --- PASSO 1 de 2: Compilando o VerificadorApp.exe ---
pyinstaller VerificadorApp.spec --noconfirm
IF %ERRORLEVEL% NEQ 0 (
    echo !!! ERRO AO COMPILAR O VerificadorApp.exe !!!
    goto:FIM
)
echo VerificadorApp.exe compilado com sucesso.

echo.
echo --- PASSO 2 de 2: Compilando o Atualizador.exe ---
pyinstaller Atualizador.spec --noconfirm
IF %ERRORLEVEL% NEQ 0 (
    echo !!! ERRO AO COMPILAR O Atualizador.exe !!!
    goto:FIM
)
echo Atualizador.exe compilado com sucesso.

echo.
echo --- Finalizando: Movendo o Atualizador.exe para a pasta principal ---

REM Copia o Atualizador.exe para a pasta do VerificadorApp
move /Y dist\Atualizador\Atualizador.exe dist\VerificadorApp\

REM Limpa a pasta agora vazia do atualizador
rmdir /s /q dist\Atualizador

echo.
echo ==========================================================
echo      PROCESSO CONCLUIDO COM SUCESSO!
echo ==========================================================
echo Seus programas estao prontos na pasta 'dist\VerificadorApp'.

:FIM
echo.
pause