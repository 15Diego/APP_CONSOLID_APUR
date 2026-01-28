@echo off
REM Script de Deploy Automatizado para Streamlit Cloud

echo ========================================
echo  Deploy - Consolidador de Relatorios
echo ========================================
echo.

REM Verificar se Git esta inicializado
if not exist .git (
    echo [1/5] Inicializando Git...
    git init
    echo.
) else (
    echo [1/5] Git ja inicializado
    echo.
)

REM Adicionar arquivos
echo [2/5] Adicionando arquivos...
git add app.py
git add consolidate_relatorio_base.py
git add requirements.txt
git add README.md
git add .streamlit/config.toml
git add .gitignore
git add DEPLOY.md
git add Dockerfile
echo.

REM Verificar mudancas
echo [3/5] Verificando mudancas...
git status
echo.

REM Commit
echo [4/5] Criando commit...
set /p commit_msg="Digite a mensagem do commit (ou Enter para usar default): "
if "%commit_msg%"=="" set commit_msg=Deploy: Consolidador de Relatorios v2.0

git commit -m "%commit_msg%"
echo.

REM Instrucoes para remote
echo [5/5] Proximo passo: Configurar repositorio remoto
echo.
echo Para conectar ao GitHub, execute:
echo   git remote add origin https://github.com/SEU_USUARIO/consolidador-relatorios.git
echo   git branch -M main
echo   git push -u origin main
echo.
echo Depois, acesse: https://share.streamlit.io
echo.

pause
